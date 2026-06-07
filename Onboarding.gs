// ===============================================
// BC1890 - Mitglieder Onboarding & Willkommens-System
// Google Apps Script für automatische Willkommens-Mails
// (Nutzt die globale CONFIG aus Reservierungssystem.gs)
// ===============================================

const ONBOARDING_CONFIG = {
  // <--- SCHALTER FÜR DEN TESTMODUS --->
  // true  = Mails werden abgefangen und NUR an den Vorstand gesendet (Sicherer Testmodus)
  // false = Mails gehen direkt an die neuen Mitglieder und der Vorstand im CC (Live-Betrieb)
  TEST_MODUS_AKTIV: true, 

  // <--- HIER DIE ID DEINER PDF-ANLEITUNG AUS GOOGLE DRIVE EINTRAGEN --->
  PDF_FILE_ID: 'PDF-ID-HIER-EINFÜGEN' 
};

/**
 * Hauptfunktion: Prüft auf neue Mitglieder, bereinigt gelöschte IDs 
 * und versendet die Willkommens-Mail.
 */
function checkAndWelcomeNewMembers() {
  const modusText = ONBOARDING_CONFIG.TEST_MODUS_AKTIV ? '⚠️ TESTMODUS (AKTIV)' : '🚀 LIVE-BETRIEB';
  Logger.log(`=== STARTE PRÜFUNG AUF NEUE MITGLIEDER [Modus: ${modusText}] ===`);
  
  // Überprüfung der CONFIG
  if (typeof CONFIG === 'undefined' || !CONFIG.SHEET_CONFIG_ID || !CONFIG.ADMIN_EMAIL) {
    Logger.log('❌ FEHLER: Das CONFIG-Objekt aus dem Hauptskript wurde nicht gefunden oder ist unvollständig.');
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Lädt die Liste der bisher bekannten Mitglieder-IDs
  let welcomedMembersRaw = scriptProperties.getProperty('WELCOMED_MEMBER_IDS');
  let welcomedMemberIds = welcomedMembersRaw ? JSON.parse(welcomedMembersRaw) : [];
  
  // Flag für den allerersten Systemstart (wenn die Datenbank komplett leer ist)
  const isInitialRun = welcomedMemberIds.length === 0;
  if (isInitialRun) {
    Logger.log('Erster Durchlauf erkannt. Bestehende Mitglieder werden erfasst, ohne E-Mails zu senden.');
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_CONFIG_ID);
    const sheet = ss.getSheets()[0]; 
    if (!sheet) {
      Logger.log('Fehler: Tabelle konnte nicht geöffnet werden.');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('Keine Mitgliederdaten in der Tabelle gefunden.');
      // Falls die Tabelle komplett geleert wurde, löschen wir auch das Gedächtnis
      scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([]));
      return;
    } 
    
    // Daten ab Zeile 2 einlesen (Spalten: ID, Vorname, Nachname, E-Mail, Mobile)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    // 1. SCHRITT: Alle AKTUELLEN IDs aus der Tabelle sammeln
    let currentTableIds = [];
    for (let i = 0; i < dataRange.length; i++) {
      const memberId = dataRange[i][0] ? dataRange[i][0].toString().trim() : '';
      if (memberId) {
        currentTableIds.push(memberId);
      }
    }

    // 2. SCHRITT: Bereinigung (Gelöschte Mitglieder entfernen)
    // Wir filtern das alte Gedächtnis und behalten NUR IDs, die auch jetzt noch in der Tabelle existieren.
    // Beim "isInitialRun" ist welcomedMemberIds ohnehin leer, daher überspringen wir das dort logisch.
    let cleanedWelcomedIds = welcomedMemberIds.filter(id => currentTableIds.includes(id));
    
    let removedCount = welcomedMemberIds.length - cleanedWelcomedIds.length;
    if (removedCount > 0) {
      Logger.log(`🧹 BEREINIGUNG: ${removedCount} gelöschte(s) Mitglied(er) aus dem Skript-Gedächtnis entfernt.`);
    }

    // 3. SCHRITT: Neue Mitglieder prüfen und begrüssen
    let mailsSentCount = 0;

    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      const memberId = row[0] ? row[0].toString().trim() : '';
      const vorname = row[1] ? row[1].toString().trim() : '';
      const nachname = row[2] ? row[2].toString().trim() : '';
      const email = row[3] ? row[3].toString().trim() : '';
      
      if (!memberId || !email) continue; // Überspringe unvollständige Zeilen

      // Prüfen, ob diese ID in unserer bereinigten Liste fehlt
      if (!cleanedWelcomedIds.includes(memberId)) {
        
        if (!isInitialRun) {
          // Neues Mitglied gefunden -> Willkommens-Mail senden
          sendWelcomeMail(email, vorname, nachname);
          mailsSentCount++;
        }
        
        // Die neue ID direkt zur bereinigten Liste hinzufügen
        cleanedWelcomedIds.push(memberId);
      }
    }

    // 4. SCHRITT: Speicher aktualisieren
    scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify(cleanedWelcomedIds));
    Logger.log(`Prüfung abgeschlossen. ${mailsSentCount} neue(s) Mitglied(er) verarbeitet.`);
    
  } catch (e) {
    Logger.log('Fehler im Onboarding-Script: ' + e.message);
  }
  
  Logger.log('=== PRÜFUNG BEENDET ===');
}

/**
 * Versendet die Willkommens-E-Mail (Berücksichtigt den eingestellten Test- oder Livemodus)
 */
function sendWelcomeMail(toEmail, vorname, nachname) {
  const name = vorname ? vorname : 'Mitglied';
  
  // Variablen deklarieren, die sich je nach Modus ändern
  let finalReceiver = toEmail;
  let finalCc = CONFIG.ADMIN_EMAIL; // Im Live-Modus geht der Vorstand standardmäßig ins CC
  let subject = 'Herzlich willkommen beim Bootsclub 1890! ⛵';
  let testNoticeHtml = '';
  let testNoticePlain = '';

  // Logik umschalten, falls der Testmodus aktiv ist
  if (ONBOARDING_CONFIG.TEST_MODUS_AKTIV) {
    finalReceiver = CONFIG.ADMIN_EMAIL; // Mail an Vorstand umleiten
    finalCc = ''; // CC leeren, um Doppelversand an Vorstand zu vermeiden
    subject = `[TEST-MODUS für: ${toEmail}] Herzlich willkommen beim Bootsclub 1890! ⛵`;
    
    testNoticeHtml = `
      <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 12px; margin-bottom: 20px; color: #856404; font-family: sans-serif; border-radius: 4px;">
        ⚠️ <b>SYSTEM-HINWEIS (TEST-MODUS):</b> Diese E-Mail wurde automatisch abgefangen und an den Vorstand umgeleitet.<br>
        <b>Geplanter Empfänger im Live-Betrieb:</b> ${vorname} ${nachname} (&lt;${toEmail}&gt;)
      </div>
    `;
    testNoticePlain = `[⚠️ TEST-MODUS - Geplanter Empfänger im Live-Betrieb: ${vorname} ${nachname} (${toEmail})]\n\n`;
  }
  
  // Der eigentliche HTML-Inhalt der Willkommens-Mail
  const htmlBody = `
    ${testNoticeHtml}
    Hallo ${name},<br><br>
    herzlich willkommen im <b>Bootsclub 1890</b>! Deine E-Mail-Adresse wurde erfolgreich für unser automatisiertes Reservierungssystem freigeschaltet.<br><br>
    
    Ab sofort kannst du Bootstermine direkt per E-Mail reservieren. <br>
    <b>Im Anhang dieser E-Mail findest du die detaillierte Anleitung als PDF-Datei.</b><br><br>
    
    Hier sind die wichtigsten Kernpunkte im Überblick:<br>
    • Sende Reservierungen an: <b>${CONFIG.ADMIN_EMAIL}</b>. Die E-Mail muss das Wort <b>Reservierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text enthalten.<br><br>
    • Für eine Stornierung sende einfach das Wort <b>Stornierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text senden (bis max. 24 Stunden vor dem Termin).<br><br>
    
    Bitte lies dir die angehängte PDF-Anleitung aufmerksam durch, bevor du deine erste Reservierung vornimmst.<br><br>
    Bei Fragen steht dir der Vorstand jederzeit gerne zur Verfügung.<br><br>
    Allzeit gute Fahrt und viel Spass auf dem Wasser!<br><br>
    <b>Dein Vorstand</b><br>
  `;

  const plainBody = `${testNoticePlain}Hallo ${name},\n\nherzlich willkommen beim Bootsclub 1890!\nDeine E-Mail wurde für das Reservierungssystem freigeschaltet.\n\nEine detaillierte Anleitung findest du im Anhang dieser E-Mail als PDF.\n\nBitte sende Reservierungen an ${CONFIG.ADMIN_EMAIL}.\n\nAllzeit gute Fahrt!\nDein Vorstand`;

  try {
    // PDF-Anleitung aus Google Drive holen
    const fileId = ONBOARDING_CONFIG.PDF_FILE_ID;
    if (!fileId || fileId === 'HIER_DEINE_GOOGLE_DRIVE_FILE_ID_EINTRAGEN') {
      throw new Error('Es wurde keine gültige Google Drive File ID für das PDF konfiguriert.');
    }
    
    const pdfFile = DriveApp.getFileById(fileId);
    const attachmentBlob = pdfFile.getBlob();

    // E-Mail senden mit den dynamisch gesetzten Empfängern
    GmailApp.sendEmail(finalReceiver, subject, plainBody, {
      cc: finalCc,
      replyTo: CONFIG.ADMIN_EMAIL,
      htmlBody: htmlBody,
      attachments: [attachmentBlob]
    });
    
    if (ONBOARDING_CONFIG.TEST_MODUS_AKTIV) {
      Logger.log(`[TEST] Mail für ${toEmail} wurde erfolgreich an den Vorstand (${finalReceiver}) umgeleitet.`);
    } else {
      Logger.log(`[LIVE] Willkommens-E-Mail erfolgreich direkt gesendet an: ${toEmail}`);
    }
  } catch (error) {
    Logger.log(`❌ FEHLER beim Senden der E-Mail (Ziel: ${finalReceiver}): ${error.message}`);
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
