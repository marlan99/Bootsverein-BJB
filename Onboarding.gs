// ===============================================
// BC1890 - Mitglieder Onboarding & Willkommens-System
// Google Apps Script für automatische Willkommens-Mails
// (Nutzt die globale CONFIG aus Reservierungssystem.gs)
// ===============================================

const ONBOARDING_CONFIG = {
  // <--- HIER DIE ID DEINER PDF-ANLEITUNG AUS GOOGLE DRIVE EINTRAGEN --->
  // Du findest die ID in der Drive-Link-URL: https://drive.google.com/file/d/hier_steht_die_id/view
  PDF_FILE_ID: 'HIER_DEINE_GOOGLE_DRIVE_FILE_ID_EINTRAGEN'
};

/**
 * Hauptfunktion: Prüft auf neue Mitglieder und versendet die Willkommens-Mail.
 * Richte hierfür einen täglichen Zeitgesteuerten Trigger (z.B. morgens) ein.
 */
function checkAndWelcomeNewMembers() {
  Logger.log('=== STARTE PRÜFUNG AUF NEUE MITGLIEDER ===');
  
  // Überprüfung, ob das CONFIG-Objekt aus dem Hauptskript existiert
  if (typeof CONFIG === 'undefined' || !CONFIG.SHEET_CONFIG_ID || !CONFIG.ADMIN_EMAIL) {
    Logger.log('❌ FEHLER: Das CONFIG-Objekt aus dem Hauptskript wurde nicht gefunden oder ist unvollständig.');
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Lädt die Liste der bereits begrüßten Mitglieder-IDs
  let welcomedMembersRaw = scriptProperties.getProperty('WELCOMED_MEMBER_IDS');
  let welcomedMemberIds = welcomedMembersRaw ? JSON.parse(welcomedMembersRaw) : [];
  
  // Wenn das System komplett neu gestartet wird, optional alle bestehenden einlesen ohne Mail-Versand
  const isInitialRun = welcomedMemberIds.length === 0;
  if (isInitialRun) {
    Logger.log('Erster Durchlauf erkannt. Bestehende Mitglieder werden erfasst, ohne E-Mails zu senden.');
  }

  try {
    // Greift direkt auf CONFIG.SHEET_CONFIG_ID aus deinem vorhandenen Script zu
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_CONFIG_ID);
    const sheet = ss.getSheets()[0]; 
    if (!sheet) {
      Logger.log('Fehler: Tabelle konnte nicht geöffnet werden.');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('Keine Mitgliederdaten in der Tabelle gefunden.');
      return;
    } 
    
    // Daten ab Zeile 2 einlesen (Spalten: ID, Vorname, Nachname, E-Mail, Mobile)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let newWelcomedIds = [...welcomedMemberIds];
    let mailsSentCount = 0;

    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      const memberId = row[0] ? row[0].toString().trim() : '';
      const vorname = row[1] ? row[1].toString().trim() : '';
      const nachname = row[2] ? row[2].toString().trim() : '';
      const email = row[3] ? row[3].toString().trim() : '';
      
      if (!memberId || !email) continue; // Überspringe unvollständige Zeilen

      // Prüfen, ob die ID bereits registriert/begrüßt wurde
      if (!welcomedMemberIds.includes(memberId)) {
        
        if (!isInitialRun) {
          // Neues Mitglied gefunden -> E-Mail mit PDF-Anhang senden
          sendWelcomeMail(email, vorname, nachname);
          mailsSentCount++;
        }
        
        // Zur Liste der bekannten IDs hinzufügen
        newWelcomedIds.push(memberId);
      }
    }

    // Aktualisierte Liste dauerhaft im Script speichern
    scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify(newWelcomedIds));
    Logger.log(`Prüfung abgeschlossen. ${mailsSentCount} neue(s) Mitglied(er) begrüßt.`);
    
  } catch (e) {
    Logger.log('Fehler im Onboarding-Script: ' + e.message);
  }
  
  Logger.log('=== PRÜFUNG BEENDET ===');
}

/**
 * Versendet die Willkommens-E-Mail mit der PDF-Anleitung im Anhang und setzt den Vorstand ins CC
 */
function sendWelcomeMail(toEmail, vorname, nachname) {
  const name = vorname ? vorname : 'Mitglied';
  const subject = 'Herzlich willkommen beim Bootsclub 1890! ⛵';
  
  // HTML-Inhalt der Mail (Anleitungstext gekürzt, da jetzt als PDF angehängt)
  const htmlBody = `
    Hallo ${name},<br><br>
    herzlich willkommen im <b>Bootsclub 1890</b>! Deine E-Mail-Adresse wurde erfolgreich für unser automatisiertes Reservierungssystem freigeschaltet.<br><br>
    
    Ab sofort kannst du Bootstermine direkt per E-Mail reservieren. <br>
    <b>Im Anhang dieser E-Mail findest du die detaillierte Anleitung als PDF-Datei.</b><br><br>
    
    Hier sind die wichtigsten Kernpunkte im Überblick:<br>
    • Sende Reservierungen an: <b>${CONFIG.ADMIN_EMAIL}</b><br>
    • Die E-Mail muss die Zeilen <b>Datum:</b>, <b>Slot:</b> (Vormittag/Nachmittag) und <b>Typ:</b> (Standard/Joker) enthalten.<br>
    • Für eine Stornierung sende einfach das Wort <b>"Stornierung"</b> im Betreff (bis max. 24 Stunden vor dem Termin).<br><br>
    
    Bitte lies dir die angehängte PDF-Anleitung aufmerksam durch, bevor du deine erste Reservierung vornimmst.<br><br>
    Bei Fragen steht dir der Vorstand jederzeit gerne zur Verfügung.<br><br>
    Allzeit gute Fahrt und viel Spass auf dem Wasser!<br><br>
    <b>Dein Vorstand</b>
  `;

  const plainBody = `Hallo ${name},\n\nherzlich willkommen beim Bootsclub 1890!\nDeine E-Mail wurde für das Reservierungssystem freigeschaltet.\n\nEine detaillierte Anleitung findest du im Anhang dieser E-Mail als PDF.\n\nBitte sende Reservierungen an ${CONFIG.ADMIN_EMAIL}.\n\nAllzeit gute Fahrt!\nDein Vorstand`;

  try {
    // 1. PDF-Datei aus Google Drive holen
    const fileId = ONBOARDING_CONFIG.PDF_FILE_ID;
    if (!fileId || fileId === 'HIER_DEINE_GOOGLE_DRIVE_FILE_ID_EINTRAGEN') {
      throw new Error('Es wurde keine gültige Google Drive File ID für das PDF konfiguriert.');
    }
    
    const pdfFile = DriveApp.getFileById(fileId);
    const attachmentBlob = pdfFile.getBlob(); // Holt die Datei als Binärobjekt (Blob)

    // 2. E-Mail mit Attachment versenden
    GmailApp.sendEmail(toEmail, subject, plainBody, {
      cc: CONFIG.ADMIN_EMAIL, // Vorstand aus Haupt-CONFIG im CC
      replyTo: CONFIG.ADMIN_EMAIL,
      htmlBody: htmlBody,
      attachments: [attachmentBlob] // Hier wird das PDF angehängt
    });
    
    Logger.log(`Willkommens-E-Mail inkl. PDF-Anleitung erfolgreich gesendet an: ${toEmail}`);
  } catch (error) {
    Logger.log(`❌ FEHLER beim Senden der Willkommens-Mail an ${toEmail}: ${error.message}`);
  }
}

/**
 * Erstellt den täglichen automatischen Trigger für die Prüfung
 */
function setupOnboardingTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkAndWelcomeNewMembers') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Läuft jeden Tag zwischen 08:00 und 09:00 Uhr morgens
  ScriptApp.newTrigger('checkAndWelcomeNewMembers')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  Logger.log('Täglicher Onboarding-Trigger wurde erfolgreich eingerichtet.');
}

/**
 * Hilfsfunktion: Setzt die Datenbank im Script zurück (nur für Testzwecke).
 */
function resetWelcomeDatabase() {
  PropertiesService.getScriptProperties().deleteProperty('WELCOMED_MEMBER_IDS');
  Logger.log('Datenbank zurückgesetzt.');
}
