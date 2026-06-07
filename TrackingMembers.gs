// ===============================================
// BC1890 - Mitgliederlisten-Tracking-System
// Erkennt Hinzufügungen, Löschungen und Änderungen
// Unabhängige Version (ohne CONFIG-Abhängigkeit)
// ===============================================

const TRACKING_CONFIG = {
  // Schalter für den Testmodus
  // true  = Der Bericht wird gesendet, aber es wird KEIN neuer Schnappschuss gespeichert.
  // false = Der Bericht wird gesendet und der Schnappschuss aktualisiert (Normalbetrieb).
  TEST_MODUS_AKTIV: false
};

/**
 * Hauptfunktion für das Tracking: Vergleicht den aktuellen Stand mit dem letzten Schnappschuss.
 */
function tracklistchanges() {
  Logger.log('=== STARTE MITGLIEDERLISTEN-TRACKING ===');
  
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Holt sich die IDs und Einstellungen direkt aus den Umgebungsvariablen des Skripts
  // (Diese wurden beim ersten Einrichten deines Hauptskripts dort hinterlegt)
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || 
                  (typeof CONFIG !== 'undefined' ? CONFIG.SHEET_CONFIG_ID : null);
                  
  const adminEmail = scriptProperties.getProperty('ADMIN_EMAIL') || 
                    (typeof CONFIG !== 'undefined' ? CONFIG.ADMIN_EMAIL : null);

  if (!sheetId || !adminEmail) {
    Logger.log('❌ FEHLER: Weder SHEET_CONFIG_ID noch ADMIN_EMAIL konnten in den ScriptProperties oder im CONFIG-Objekt gefunden werden.');
    Logger.log('Bitte stelle sicher, dass diese Werte im Hauptskript korrekt initialisiert wurden.');
    return;
  }

  // Letzten gespeicherten Schnappschuss laden
  const lastSnapshotRaw = scriptProperties.getProperty('MEMBER_LIST_SNAPSHOT');
  
  // Aktuelle Daten aus der Tabelle einlesen
  let currentSnapshot = {};
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // Holt ID, Vorname, Nachname, E-Mail, Mobile
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      
      for (let i = 0; i < data.length; i++) {
        const id = data[i][0] ? data[i][0].toString().trim() : '';
        if (!id) continue;
        
        currentSnapshot[id] = {
          vorname: data[i][1] ? data[i][1].toString().trim() : '',
          nachname: data[i][2] ? data[i][2].toString().trim() : '',
          email: data[i][3] ? data[i][3].toString().trim() : '',
          mobile: data[i][4] ? data[i][4].toString().trim() : ''
        };
      }
    }
  } catch (e) {
    Logger.log('❌ Fehler beim Einlesen der Tabelle: ' + e.message);
    return;
  }

  // Falls noch kein Schnappschuss existiert (Initialer Lauf)
  if (!lastSnapshotRaw) {
    Logger.log('Kein alter Schnappschuss vorhanden. Erstelle initialen Datenstand...');
    scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot));
    Logger.log('=== TRACKING BEENDET (Initialer Lauf) ===');
    return;
  }

  const lastSnapshot = JSON.parse(lastSnapshotRaw);
  
  let addedMembers = [];
  let removedMembers = [];
  let updatedMembers = [];

  // 1. Auf Hinzufügungen und Änderungen prüfen
  for (let id in currentSnapshot) {
    const current = currentSnapshot[id];
    const last = lastSnapshot[id];
    current.id = id;

    if (!last) {
      addedMembers.push(current);
    } else {
      let changedFields = [];
      let textDetails = [];
      
      if (current.vorname !== last.vorname) { changedFields.push('vorname'); textDetails.push(`Vorname: ${last.vorname} -> ${current.vorname}`); }
      if (current.nachname !== last.nachname) { changedFields.push('nachname'); textDetails.push(`Nachname: ${last.nachname} -> ${current.nachname}`); }
      if (current.email !== last.email) { changedFields.push('email'); textDetails.push(`E-Mail: ${last.email} -> ${current.email}`); }
      if (current.mobile !== last.mobile) { changedFields.push('mobile'); textDetails.push(`Mobil: ${last.mobile} -> ${current.mobile}`); }

      if (changedFields.length > 0) {
        updatedMembers.push({
          id: id,
          old: last,
          current: current,
          changedFields: changedFields,
          textDetails: textDetails
        });
      }
    }
  }

  // 2. Auf Löschungen prüfen
  for (let id in lastSnapshot) {
    if (!currentSnapshot[id]) {
      const removed = lastSnapshot[id];
      removed.id = id;
      removedMembers.push(removed);
    }
  }

  // 3. Auswertung: Gab es Änderungen?
  if (addedMembers.length > 0 || removedMembers.length > 0 || updatedMembers.length > 0) {
    Logger.log(`Änderungen erkannt! Neu: ${addedMembers.length}, Gelöscht: ${removedMembers.length}, Geändert: ${updatedMembers.length}`);
    
    // E-Mail-Bericht senden
    sendChangeReportMail(adminEmail, addedMembers, removedMembers, updatedMembers);
    
    if (!TRACKING_CONFIG.TEST_MODUS_AKTIV) {
      scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot));
      Logger.log('Der neue Schnappschuss wurde erfolgreich gespeichert.');
    } else {
      Logger.log('⚠️ HINWEIS: Im Testmodus wird der alte Schnappschuss NICHT überschrieben.');
    }
  } else {
    Logger.log('Keine Änderungen an der Mitgliederliste festgestellt.');
  }

  Logger.log('=== TRACKING BEENDET ===');
}

/**
 * Erstellt und versendet den detaillierten HTML-Bericht an den Vorstand
 */
function sendChangeReportMail(adminEmail, added, removed, updated) {
  let subject = `✅ Änderungsbericht: Mitgliederliste BC1890`;
  if (TRACKING_CONFIG.TEST_MODUS_AKTIV) subject = `[TEST] ` + subject;

  const tableStyle = 'width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 25px; font-size: 14px;';
  const thStyle = 'background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 10px; text-align: left; color: #334155; font-weight: bold;';
  const tdStyle = 'border: 1px solid #e2e8f0; padding: 10px; vertical-align: top; color: #475569;';

  let htmlBody = `
    <div style="font-family: sans-serif; color: #333; max-width: 750px; line-height: 1.5;">
      <h2 style="color: #1a365d; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 20px;">Bericht: Änderungen an der Mitgliederliste</h2>
      <p>Hallo Vorstand,<br>das automatisierte System hat Änderungen in der Mitglieder-Tabelle festgestellt. Nachfolgend findest du alle Details zu den betroffenen Personen:</p>
  `;

  if (added.length > 0) {
    htmlBody += `
      <h3 style="color: #2f855a; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #c6f6d5; padding-bottom: 4px;">➕ Neu hinzugefügte Mitglieder (${added.length})</h3>
      <table style="${tableStyle}">
        <tr>
          <th style="${thStyle} width: 10%;">ID</th>
          <th style="${thStyle} width: 25%;">Name</th>
          <th style="${thStyle} width: 40%;">E-Mail</th>
          <th style="${thStyle} width: 25%;">Mobile</th>
        </tr>
    `;
    added.forEach(m => {
      htmlBody += `
        <tr>
          <td style="${tdStyle}"><code>${m.id || ''}</code></td>
          <td style="${tdStyle}"><b>${m.vorname} ${m.nachname}</b></td>
          <td style="${tdStyle}">${m.email}</td>
          <td style="${tdStyle}">${m.mobile || '-'}</td>
        </tr>`;
    });
    htmlBody += `</table>`;
  }

  if (removed.length > 0) {
    htmlBody += `
      <h3 style="color: #9b2c2c; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #fed7d7; padding-bottom: 4px;">➖ Entfernte Mitglieder (${removed.length})</h3>
      <table style="${tableStyle}">
        <tr>
          <th style="${thStyle} width: 10%;">ID</th>
          <th style="${thStyle} width: 25%;">Name</th>
          <th style="${thStyle} width: 40%;">E-Mail</th>
          <th style="${thStyle} width: 25%;">Mobile</th>
        </tr>
    `;
    removed.forEach(m => {
      htmlBody += `
        <tr style="background-color: #fafafa;">
          <td style="${tdStyle} color: #94a3b8;"><code>${m.id || ''}</code></td>
          <td style="${tdStyle} color: #94a3b8;">${m.vorname} ${m.nachname}</td>
          <td style="${tdStyle} color: #94a3b8;">${m.email}</td>
          <td style="${tdStyle} color: #94a3b8;">${m.mobile || '-'}</td>
        </tr>`;
    });
    htmlBody += `</table>`;
  }

  if (updated.length > 0) {
    htmlBody += `
      <h3 style="color: #dd6b20; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #feebc8; padding-bottom: 4px;">⚠️ Aktualisierte Mitgliedsdaten (${updated.length})</h3>
    `;
    
    updated.forEach(m => {
      const vNameStyle = m.changedFields.includes('vorname') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
      const nNameStyle = m.changedFields.includes('nachname') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
      const emailStyle = m.changedFields.includes('email') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
      const mobilStyle = m.changedFields.includes('mobile') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';

      htmlBody += `
        <div style="margin-bottom: 25px; border-left: 4px solid #dd6b20; padding-left: 12px;">
          <span style="font-size: 15px; font-weight: bold; color: #2d3748;">Mitglied: ${m.current.vorname} ${m.current.nachname}</span> 
          <span style="font-size: 13px; color: #718096; margin-left: 10px;">(ID: <code>${m.id}</code>)</span>
          
          <table style="${tableStyle} margin-top: 6px; margin-bottom: 5px;">
            <tr style="background-color: #f8fafc;">
              <th style="${thStyle} width: 25%;">Feld</th>
              <th style="${thStyle} width: 37.5%;">Alter Wert (Schnappschuss)</th>
              <th style="${thStyle} width: 37.5%;">Neuer Wert (Tabelle)</th>
            </tr>
            <tr>
              <td style="${tdStyle} ${vNameStyle}">Vorname</td>
              <td style="${tdStyle} ${vNameStyle}">${m.old.vorname || '-'}</td>
              <td style="${tdStyle} ${vNameStyle}">${m.current.vorname || '-'}</td>
            </tr>
            <tr>
              <td style="${tdStyle} ${nNameStyle}">Nachname</td>
              <td style="${tdStyle} ${nNameStyle}">${m.old.nachname || '-'}</td>
              <td style="${tdStyle} ${nNameStyle}">${m.current.nachname || '-'}</td>
            </tr>
            <tr>
              <td style="${tdStyle} ${emailStyle}">E-Mail</td>
              <td style="${tdStyle} ${emailStyle}">${m.old.email || '-'}</td>
              <td style="${tdStyle} ${emailStyle}">${m.current.email || '-'}</td>
            </tr>
            <tr>
              <td style="${tdStyle} ${mobilStyle}">Mobile</td>
              <td style="${tdStyle} ${mobilStyle}">${m.old.mobile || '-'}</td>
              <td style="${tdStyle} ${mobilStyle}">${m.current.mobile || '-'}</td>
            </tr>
          </table>
        </div>
      `;
    });
  }

  htmlBody += `
      <hr style="border: 0; border-top: 1px solid #e2e8f0; margin-top: 40px;">
      <p style="font-size: 12px; color: #a0aec0;">Dieses Protokoll wurde automatisch generiert. Zeitstempel: ${new Date().toLocaleString('de-DE')}</p>
    </div>
  `;

  let plainBody = `Änderungsbericht Mitgliederliste BC1890\n\n`;
  if (added.length > 0) plainBody += `Neu (${added.length}):\n` + added.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}, Mail: ${m.email}, Mobil: ${m.mobile}`).join('\n') + `\n\n`;
  if (removed.length > 0) plainBody += `Entfernt (${removed.length}):\n` + removed.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}, Mail: ${m.email}, Mobil: ${m.mobile}`).join('\n') + `\n\n`;
  if (updated.length > 0) plainBody += `Geändert (${updated.length}):\n` + updated.map(m => `- ID: ${m.id}, Name: ${m.current.vorname} ${m.current.nachname}\n  Änderungen: ${m.textDetails.join(', ')}`).join('\n') + `\n`;

  try {
    GmailApp.sendEmail(adminEmail, subject, plainBody, {
      htmlBody: htmlBody
    });
    Logger.log('📧 Ausführlicher Änderungsbericht erfolgreich an den Vorstand versendet.');
  } catch (err) {
    Logger.log('❌ Fehler beim Senden des Änderungsberichts: ' + err.message);
  }
}

/**
 * Hilfsfunktion: Setzt das Tracking-Gedächtnis zurück.
 */
function resetTrackingSnapshot() {
  PropertiesService.getScriptProperties().deleteProperty('MEMBER_LIST_SNAPSHOT');
  Logger.log('Tracking-Schnappschuss wurde erfolgreich gelöscht.');
}
