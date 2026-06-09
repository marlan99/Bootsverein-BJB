// =============================================================================
// BC1890 - Integriertes Gesamtskript: Reservierung, Onboarding & Tracking
// Google Apps Script für Google Calendar, Gmail, Drive & Google Spreadsheet
// =============================================================================

// Globale URL-Quelle für die PDF-Anleitung
const PDF_SOURCE_URL = 'https://github.com/marlan99/Bootsverein-BJB/blob/main/Anleitung%20Bootsreservation.pdf';

const CONFIG = {
  CALENDAR_ID: '',  // Hier die KALENDER ID eintragen, falls nicht der Standardkalender verwendet wird
  ADMIN_EMAIL: Session.getActiveUser().getEmail(),
  GMAIL_LABEL: 'Reservierung/Neu',
  SLOT_VORMITTAG: { start: '08:00', end: '14:00' },
  SLOT_NACHMITTAG: { start: '14:00', end: '20:00' },
  EXCEL_SUBJECT: 'Mitgliederliste',
  EXCEL_TARGET_LABEL: 'Reservierung/Mitgliederliste',
  TEST_MODUS_AKTIV: false,
  TRACKING_TEST_MODUS_AKTIV: false,
};

// =============================================================================
// 1. KERN-LOGIK: RESERVIERUNGEN & STORNIERUNGEN VERARBEITEN (OPTIMIERT)
// =============================================================================

function processReservationEmails() {
  let labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL) || createGmailLabelStructure(CONFIG.GMAIL_LABEL);

  // OPTIMIERUNG 1: Kombinierte Suchanfrage spart API-Quota und verhindert Doppelverarbeitung
  const emailThreads = GmailApp.search('in:inbox (subject:"Reservierung" OR subject:"Stornierung")');
  Logger.log(`Gefundene relevante Threads im Posteingang: ${emailThreads.length}`);
  
  // OPTIMIERUNG 2: Kalender-Instanz EINMALIG holen und wiederverwenden
  const calendar = CONFIG.CALENDAR_ID ?
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
  if (!calendar) {
    Logger.log('❌ KRITISCHER FEHLER: Kalender konnte nicht geladen werden.');
    return;
  }

  // Arrays für die Batch-Label-Verarbeitung (Optimierung 1)
  const threadsErledigt = [];
  const threadsAbgelehnt = [];
  const threadsStorniert = [];

  emailThreads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread()) {
        // Optimierung 4: Sofort als gelesen markieren, um Endlosschleifen bei Timeouts zu verhindern
        message.markRead();
        thread.addLabel(labelNeu);
        
        // Verarbeiten und den Thread anhand des Rückgabestatus kategorisieren
        const status = processSingleEmail(message, thread, calendar);
        
        if (status === 'ERLEDIGT') {
          threadsErledigt.push(thread);
        } else if (status === 'ABGELEHNT') {
          threadsAbgelehnt.push(thread);
        } else if (status === 'STORNIERT') {
          threadsStorniert.push(thread);
        }
      }
    });
  });

// Batch-Label-Zuweisung & Archivierung (FÜR BUCHUNGEN & STORNIERUNGEN)
  if (threadsErledigt.length > 0) {
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') ||
      GmailApp.createLabel('Reservierung/Erledigt');
    labelErledigt.addToThreads(threadsErledigt);
    if (labelNeu) {
      labelNeu.removeFromThreads(threadsErledigt);
    }
    threadsErledigt.forEach(thread => thread.moveToArchive());
  }
  
  if (threadsAbgelehnt.length > 0) {
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') ||
      GmailApp.createLabel('Reservierung/Abgelehnt');
    labelAbgelehnt.addToThreads(threadsAbgelehnt);
    if (labelNeu) {
      labelNeu.removeFromThreads(threadsAbgelehnt);
    }
    threadsAbgelehnt.forEach(thread => thread.moveToArchive());
  }

  // Erfolgreiche Stornierungen gehen nun ebenfalls in 'Reservierung/Erledigt'
  if (threadsStorniert.length > 0) {
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') ||
      GmailApp.createLabel('Reservierung/Erledigt');
    
    labelErledigt.addToThreads(threadsStorniert);
    if (labelNeu) {
      labelNeu.removeFromThreads(threadsStorniert);
    }
    threadsStorniert.forEach(thread => thread.moveToArchive());
  }

  if (emailThreads.length > 0) {
    GmailApp.refreshThreads(emailThreads);
  }
}

function processSingleEmail(message, thread, calendar) {
  const sender = message.getFrom().match(/[\w.-]+@[\w.-]+/)?.[0] || 'unbekannt';
  const subject = message.getSubject().toLowerCase();
  const body = message.getPlainBody();
  const data = parseEmailTemplate(body);
  
  // Optimierung 3: Fehlerhaften Zugriff CONFIG.CONFIG? korrigiert auf CONFIG.GMAIL_LABEL
  const labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);

  if (!data.valid) {
    sendRejectionEmail(sender, data.error, thread);
    return 'ABGELEHNT';
  }

  const userId = sender;
  // Erleichterte Erkennung von Stornierungen
  if (subject.includes('stornierung') || subject.includes('absage')) {
    const cancellationSuccess = executeCancellation(data, userId, thread, message);
    return cancellationSuccess ? 'STORNIERT' : 'ABGELEHNT';
  }

  // Kalender wird hier direkt übergeben
  const validation = validateRequest(data, userId, sender, calendar);
  if (!validation.valid) {
    sendRejectionEmail(sender, validation.error, thread);
    return 'ABGELEHNT';
  }

  // Kalender wird hier direkt übergeben
  const event = createCalendarEvent(data, userId, calendar);
  if (event) {
    sendConfirmationEmail(sender, event, data, thread);
    return 'ERLEDIGT';
  } else {
    sendRejectionEmail(sender, 'Fehler beim Erstellen des Termins im Google Kalender.', thread);
    return 'ABGELEHNT';
  }
}

function parseEmailTemplate(body) {
  // OPTIMIERUNG 3: Regex-Split fängt Windows-Zeilenumbrüche (\r\n) sauber ab
  const lines = body.split(/\r?\n/).map(l => l.trim());
  const data = { valid: false };

  const fields = {
    'Datum': 'date',
    'Slot': 'slot',
    'Typ': 'type',
    'Beschreibung': 'description',
    'Anlass': 'occasion'
  };
  lines.forEach(line => {
    for (const [key, prop] of Object.entries(fields)) {
      if (line.toLowerCase().startsWith(key.toLowerCase() + ':')) { // Tolerant gegenüber Groß-/Kleinschreibung beim Key
        data[prop] = line.substring(key.length + 1).trim();
      }
    }
  });
  if (!data.date || !data.slot) {
    data.error = 'Fehlende Pflichtfelder im Text: "Datum:" oder "Slot:" konnten nicht extrahiert werden.';
    return data;
  }

  data.parsedDate = parseEuropeanDate(data.date);

  if (!data.parsedDate || isNaN(data.parsedDate.getTime())) {
    data.error = 'Ungültiges Datum. Das Datum konnte nicht erkannt werden (Erlaubt z.B.: 05.06.2026, 5.6.2026, 5/6/2026 oder 5. Juni 2026).';
    return data;
  }

  data.parsedDate.setHours(0, 0, 0, 0);

  data.slot = data.slot.toLowerCase();
  if (!['vormittag', 'nachmittag'].includes(data.slot)) {
    data.error = 'Der angegebene Slot ist ungültig. Erlaubt ist: "Vormittag" oder "Nachmittag".';
    return data;
  }

  data.type = data.type ? data.type.toLowerCase() : 'standard';
  if (!['standard', 'joker'].includes(data.type)) {
    data.error = 'Der Typ kann nur "Standard" oder "Joker" sein.';
    return data;
  }

  data.valid = true;
  return data;
}

// Akzeptiert jetzt die bestehende Kalenderinstanz
function validateRequest(data, userId, sender, calendar) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const memberData = getAuthorizedUserData(userId);
  
  if (!memberData) {
    return { 
      valid: false, 
      error: `Deine E-Mail-Adresse (${userId}) ist nicht für das Reservierungssystem freigeschaltet. Bitte wende dich an den Vorstand.`
    };
  }
  
  data.memberId = memberData.id;
  data.memberMobile = memberData.mobile;
  if (memberData.name) data.name = memberData.name;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  // Prüfen des frühestmöglichen Startdatums
  const startDatumRaw = scriptProperties.getProperty('EARLIEST_BOOKING_DATE');
  if (startDatumRaw) {
    const parts = startDatumRaw.split('.');
    if (parts.length >= 3) {
      const startTag = parseInt(parts[0], 10);
      const startMonat = parseInt(parts[1], 10) - 1;
      const startJahr = parseInt(parts[2], 10);
      if (startJahr === today.getFullYear()) {
        const earliestAllowedDate = new Date(startJahr, startMonat, startTag, 0, 0, 0, 0);
        if (today < earliestAllowedDate) {
          const formatiertesStartDatum = `${String(startTag).padStart(2, '0')}.${String(startMonat + 1).padStart(2, '0')}.${startJahr}`;
          return { 
            valid: false, 
            error: `Das Reservierungssystem ist für das aktuelle Jahr noch nicht freigeschaltet. Buchungen sind erst ab dem ${formatiertesStartDatum} möglich.`
          };
        }
      }
    }
  }

  if (data.parsedDate < today) {
    return { valid: false, error: 'Das gewählte Datum liegt in der Vergangenheit.' };
  }

  const seasonStart = getCurrentSeasonStart();

  // JOKER-VALIDIERUNG
  if (data.type === 'joker') {
    if (data.parsedDate.getFullYear() !== today.getFullYear()) {
      return { valid: false, error: `Joker-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.` };
    }

    const seasonEnd = new Date(seasonStart);
    seasonEnd.setFullYear(seasonStart.getFullYear() + 1);

    const allEvents = calendar.getEvents(seasonStart, seasonEnd);
    const jokerEvents = allEvents.filter(e => {
      const desc = e.getDescription() || '';
      const title = e.getTitle() || '';
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && title.includes('JOKER');
    });

    if (jokerEvents.length >= 2) {
      return { valid: false, error: 'Du hast bereits das Maximum von 2 Joker-Terminen in dieser Saison erreicht.' };
    }
  }

  // STANDARD-VALIDIERUNG
  if (data.type === 'standard') {
    if (data.parsedDate.getFullYear() !== today.getFullYear()) {
      return { valid: false, error: `Standard-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.` };
    }

    const seasonEnd = new Date(today.getFullYear(), 11, 31, 23, 59, 59);
    const existingEvents = calendar.getEvents(seasonStart, seasonEnd);
    const activeStandardEvents = existingEvents.filter(e => {
      const desc = e.getDescription() || '';
      const title = e.getTitle() || '';
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && !title.includes('JOKER') && e.getStartTime() >= today;
    });

    if (activeStandardEvents.length > 0) {
      const bestehenderTermin = activeStandardEvents[0];
      return {
        valid: false,
        error: `Du hast bereits einen aktiven Standard-Termin gebucht (am ${formatDateDDMMYYYY(bestehenderTermin.getStartTime())}). Erst wenn dieser Termin vorbei ist, kannst du einen neuen Standard-Termin vereinbaren.`
      };
    }
  }

  const slotTime = data.slot === 'vormittag' ? CONFIG.SLOT_VORMITTAG : CONFIG.SLOT_NACHMITTAG;
  const startTime = new Date(data.parsedDate);
  const [sh, sm] = slotTime.start.split(':');
  startTime.setHours(sh, sm, 0, 0);

  const endTime = new Date(startTime);
  const [eh, em] = slotTime.end.split(':');
  endTime.setHours(eh, em, 0, 0);

  const conflicting = calendar.getEvents(startTime, endTime);
  if (conflicting.length > 0) {
    return { valid: false, error: 'Dieser Zeitraum (Slot) ist bereits von einem anderen Mitglied belegt.' };
  }

  data.startTime = startTime;
  data.endTime = endTime;

  return { valid: true };
}

// Akzeptiert jetzt die bestehende Kalenderinstanz
function createCalendarEvent(data, userId, calendar) {
  try {
    const myPrefix = 'Boot:';
    const title = data.type === 'joker' ? `JOKER - ${myPrefix} ${data.name}` : `${myPrefix} ${data.name}`;
    const description = [
      `Name: ${data.name}`,
      `Mitglieder-ID: ${data.memberId || 'Nicht hinterlegt'}`,
      `Kontakt: ${userId}`,
      `Mobile: ${data.memberMobile || 'Nicht hinterlegt'}`,
      `Slot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}`,
      `Typ: ${data.type.charAt(0).toUpperCase() + data.type.slice(1)}`,
      data.description ? `Beschreibung: ${data.description}` : '',
      data.occasion ? `Anlass: ${data.occasion}` : '',
      `Eingereicht per E-Mail`
    ].filter(Boolean).join('\n');
    const event = calendar.createEvent(title, data.startTime, data.endTime, { description: description });
    event.setColor(data.type === 'joker' ? CalendarApp.EventColor.RED : CalendarApp.EventColor.BLUE);

    return event;
  } catch (e) {
    Logger.log('Fehler beim Erstellen des Kalendereintrags: ' + e);
    return null;
  }
}

// =============================================================================
// 2. EXCEL-IMPORT SYSTEM (EXCEL -> GOOGLE SHEET) - OPTIMIERT & MIT ARCHIVIERUNG
// =============================================================================

function importExcelToSheets() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  const adminEmail = CONFIG.ADMIN_EMAIL;
  if (!sheetId || !adminEmail) {
    Logger.log("❌ KRITISCHER FEHLER: Tabellen-ID ('SHEET_CONFIG_ID') oder Admin-E-Mail konnte nicht ermittelt werden.");
    return;
  }
  
  const searchQuery = `subject:"${CONFIG.EXCEL_SUBJECT}" is:unread`;
  const threads = GmailApp.search(searchQuery);
  
  Logger.log(`Prüfe Posteingang auf neue Excel-Listen... Gefunden: ${threads.length}`);
  const adminEmailLower = adminEmail.toLowerCase();
  const targetLabel = GmailApp.getUserLabelByName(CONFIG.EXCEL_TARGET_LABEL) || createGmailLabelStructure(CONFIG.EXCEL_TARGET_LABEL);
  const errorLabel = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
  
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    let importErfolgreich = false;
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      if (!message.isUnread()) continue;
      // Nur ungelesene Nachrichten der Konversation betrachten

      // Sofort auf gelesen setzen, um Timeouts abzufangen
      message.markRead();
      const sender = message.getFrom().toLowerCase();
      const subject = message.getSubject();
      
      if (subject !== CONFIG.EXCEL_SUBJECT) continue;
      
      // Berechtigungsprüfung via String-Vergleich
      if (!sender.includes(adminEmailLower)) {
        Logger.log(`WARNUNG: E-Mail von unbefugtem Absender blockiert: ${sender}`);
        if (errorLabel) thread.addLabel(errorLabel);
        continue;
      }
      
      const attachments = message.getAttachments();
      for (let k = 0; k < attachments.length; k++) {
        const attachment = attachments[k];
        const isExcel = attachment.getContentType() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                        attachment.getName().toLowerCase().endsWith(".xlsx");
        
        if (!isExcel) continue;

        Logger.log(`Verarbeite Excel-Anhang: ${attachment.getName()}`);
        const fileBlob = attachment.copyBlob();
        let tempSheetFile = null;
        
        try {
          const resource = {
            title: "temp_mitgliederliste_import_" + new Date().getTime(),
            mimeType: MimeType.GOOGLE_SHEETS
          };
          // Temporäres Google Sheet aus Excel-Blob erstellen
          tempSheetFile = Drive.Files.create(resource, fileBlob);
          const tempSpreadsheet = SpreadsheetApp.openById(tempSheetFile.id);
          const tempSheet = tempSpreadsheet.getSheets()[0];
          const tempLastRow = tempSheet.getLastRow();
          const tempLastColumn = tempSheet.getLastColumn();
          if (tempLastRow <= 1) {
            Logger.log(`⚠️ Excel-Datei ${attachment.getName()} enthält keine Datenzeilen.`);
            continue;
          }
          
          // Daten im Speicher sichern
          const newValues = tempSheet.getRange(2, 1, tempLastRow - 1, tempLastColumn).getValues();
          // Zugriff auf Ziel-Tabelle erst JETZT, wenn Daten validiert sind
          const targetSpreadsheet = SpreadsheetApp.openById(sheetId);
          const targetSheet = targetSpreadsheet.getSheets()[0];
          const targetLastRow = targetSheet.getLastRow();
          
          // Erst bestehende Daten löschen (Ab Zeile 3)
          if (targetLastRow > 2) {
            targetSheet.getRange(3, 1, targetLastRow - 2, targetSheet.getLastColumn()).clearContent();
          }
          
          // Neue Daten reinschreiben
          targetSheet.getRange(3, 1, newValues.length, tempLastColumn).setValues(newValues);
          Logger.log(`✅ Mitgliederliste erfolgreich durch Excel-Mail aktualisiert (${newValues.length} Mitglieder).`);
          
          importErfolgreich = true;
          break; // Schleife für Anhänge abbrechen, da Import erfolgreich
          
        } catch (e) {
          Logger.log(`❌ Fehler beim Verarbeiten der Import-Datei: ${e.message}`);
        } finally {
          // Sicheres Löschen der temporären Datei
          if (tempSheetFile && tempSheetFile.id) {
            try { 
              DriveApp.getFileById(tempSheetFile.id).setTrashed(true);
            } catch(err) {
              Logger.log(`Hinweis beim Aufräumen: Temp-Datei konnte nicht gelöscht werden: ${err.message}`);
            }
          }
        }
      }
      if (importErfolgreich) break;
    }
    
    // E-Mail-Status finalisieren und aus der Inbox entfernen
    if (importErfolgreich) {
      if (targetLabel) thread.addLabel(targetLabel);
    } else {
      Logger.log(`⚠️ Thread [${thread.getFirstMessageSubject()}] wurde verarbeitet, konnte aber nicht erfolgreich importiert werden.`);
      if (errorLabel) thread.addLabel(errorLabel);
    }
    
    // NEU: Erzwingt das Verschieben ins Archiv (Entfernt das Posteingangs-Label)
    thread.moveToArchive();
  }

  // NEU: Zwingt Gmail dazu, die Posteingangsansicht sofort zu aktualisieren
  if (threads.length > 0) {
    GmailApp.refreshThreads(threads);
  }

  // ─── KETTENREAKTION: TRACKING WIRD BEI JEDEM DURCHLAUF GESTARTET ──────────
  if (typeof tracklistchanges === 'function') {
    Logger.log("🔎 Starte routinemässige Prüfung auf manuelle Änderungen (tracklistchanges)...");
    tracklistchanges();
  } else {
    Logger.log("Hinweis: Die Funktion tracklistchanges wurde nicht gefunden.");
  }
}

// =============================================================================
// 3. MITGLIEDERLISTEN-TRACKING-SYSTEM (DATENÄNDERUNGEN ERKENNEN) - OPTIMIERT
// =============================================================================

function tracklistchanges() {
  Logger.log('=== STARTE MITGLIEDERLISTEN-TRACKING ===');
  
  const scriptProperties = PropertiesService.getScriptProperties();
  // Alle Properties in EINEM einzigen Netzwerkaufruf holen
  const allProperties = scriptProperties.getProperties();
  const sheetId = allProperties['SHEET_CONFIG_ID'] || CONFIG.SHEET_CONFIG_ID;
  const adminEmail = allProperties['ADMIN_EMAIL'] || CONFIG.ADMIN_EMAIL;

  if (!sheetId || !adminEmail) {
    Logger.log('❌ FEHLER: Weder SHEET_CONFIG_ID noch ADMIN_EMAIL konnten gefunden werden.');
    return;
  }

  const lastSnapshotRaw = allProperties['MEMBER_LIST_SNAPSHOT'];
  const currentSnapshot = {};

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
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
    Logger.log('❌ Fehler beim Einlesen der Tabelle für Tracking: ' + e.message);
    return;
  }

  if (!lastSnapshotRaw) {
    Logger.log('Kein alter Schnappschuss vorhanden. Erstelle initialen Datenstand...');
    scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot));
    Logger.log('=== TRACKING BEENDET (Initialer Lauf) ===');
    
    if (typeof checkAnd === 'function') {
      checkAnd();
    }
    return;
  }

  const lastSnapshot = JSON.parse(lastSnapshotRaw);
  const addedMembers = [];
  const updatedMembers = [];

  // OPTIMIERUNG 1: Abgleich und direktes Erkennen von Updates & Neuzugängen
  for (const id in currentSnapshot) {
    const current = currentSnapshot[id];
    const last = lastSnapshot[id];
    current.id = id;

    if (!last) {
      addedMembers.push(current);
    } else {
      const changedFields = [];
      const textDetails = [];
      // Felder dynamisch prüfen statt 4x hartem "if"
      const fieldsToTrack = { vorname: 'Vorname', nachname: 'Nachname', email: 'E-Mail', mobile: 'Mobil' };
      for (const [field, label] of Object.entries(fieldsToTrack)) {
        if (current[field] !== last[field]) {
          changedFields.push(field);
          textDetails.push(`${label}: ${last[field] || '-'} -> ${current[field] || '-'}`);
        }
      }

      if (changedFields.length > 0) {
        updatedMembers.push({
          id: id,
          old: last,
          current: current,
          changedFields: changedFields,
          textDetails: textDetails
        });
      }
      
      // OPTIMIERUNG 2: Gefundene IDs aus dem alten Snapshot löschen.
      // Alles was am Ende übrig bleibt, wurde aus der Tabelle gelöscht!
      delete lastSnapshot[id];
    }
  }

  // Was jetzt noch im alten Snapshot ist, wurde entfernt
  const removedMembers = Object.keys(lastSnapshot).map(id => {
    const removed = lastSnapshot[id];
    removed.id = id;
    return removed;
  });

  if (addedMembers.length > 0 || removedMembers.length > 0 || updatedMembers.length > 0) {
    Logger.log(`Änderungen erkannt! Neu: ${addedMembers.length}, Gelöscht: ${removedMembers.length}, Geändert: ${updatedMembers.length}`);
    sendChangeReportMail(adminEmail, addedMembers, removedMembers, updatedMembers);
    
    if (!CONFIG.TRACKING_TEST_MODUS_AKTIV) {
      scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot));
      Logger.log('Der neue Schnappschuss wurde erfolgreich gespeichert.');
    } else {
      Logger.log('⚠️ HINWEIS: Im Tracking-Testmodus wird der alte Schnappschuss NICHT überschrieben.');
    }
  } else {
    Logger.log('Keine Änderungen an der Mitgliederliste festgestellt.');
  }

  // KETTENREAKTION: Am Ende des Trackings direkt das Onboarding triggern
  if (typeof checkAnd === 'function') {
    Logger.log("🚀 Starte automatische Prüfung auf neue Mitglieder (checkAnd)...");
    checkAnd();
  }

  Logger.log('=== TRACKING BEENDET ===');
}

function sendChangeReportMail(adminEmail, added, removed, updated) {
  let subject = `✅ Änderungsbericht: Mitgliederliste BC1890`;
  if (CONFIG.TRACKING_TEST_MODUS_AKTIV) subject = `[TEST] ` + subject;

  const tableStyle = 'width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 25px; font-size: 14px;';
  const thStyle = 'background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 10px; text-align: left; color: #334155; font-weight: bold;';
  const tdStyle = 'border: 1px solid #e2e8f0; padding: 10px; vertical-align: top; color: #475569;';
  
  // OPTIMIERUNG 3 / 5: Konsequent HTML-Push-Array statt träger String-Verkettung im inneren Loop
  const html = [
    '<div style="font-family: sans-serif; color: #333; max-width: 750px; line-height: 1.5;">',
    '<h2 style="color: #1a365d; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 20px;">Bericht: Änderungen an der Mitgliederliste</h2>',
    '<p>Hallo Vorstand,<br>das automatisierte System hat Änderungen in der Mitglieder-Tabelle festgestellt. Nachfolgend findest du alle Details:</p>'
  ];

  if (added.length > 0) {
    html.push(`<h3 style="color: #2f855a; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #c6f6d5; padding-bottom: 4px;">➕ Neu hinzugefügte Mitglieder (${added.length})</h3>`,
              `<table style="${tableStyle}"><tr><th style="${thStyle} width: 10%;">ID</th><th style="${thStyle} width: 25%;">Name</th><th style="${thStyle} width: 40%;">E-Mail</th><th style="${thStyle} width: 25%;">Mobile</th></tr>`);
    added.forEach(m => { 
      html.push(`<tr><td style="${tdStyle}"><code>${m.id || ''}</code></td><td style="${tdStyle}"><b>${m.vorname} ${m.nachname}</b></td><td style="${tdStyle}">${m.email}</td><td style="${tdStyle}">${m.mobile || '-'}</td></tr>`);
    });
    html.push('</table>');
  }

  if (removed.length > 0) {
    html.push(`<h3 style="color: #9b2c2c; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #fed7d7; padding-bottom: 4px;">➖ Entfernte Mitglieder (${removed.length})</h3>`,
              `<table style="${tableStyle}"><tr><th style="${thStyle} width: 10%;">ID</th><th style="${thStyle} width: 25%;">Name</th><th style="${thStyle} width: 40%;">E-Mail</th><th style="${thStyle} width: 25%;">Mobile</th></tr>`);
    removed.forEach(m => { 
      html.push(`<tr style="background-color: #fafafa;"><td style="${tdStyle} color: #94a3b8;"><code>${m.id || ''}</code></td><td style="${tdStyle} color: #94a3b8;">${m.vorname} ${m.nachname}</td><td style="${tdStyle} color: #94a3b8;">${m.email}</td><td style="${tdStyle} color: #94a3b8;">${m.mobile || '-'}</td></tr>`);
    });
    html.push('</table>');
  }

  if (updated.length > 0) {
    html.push(`<h3 style="color: #dd6b20; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #feebc8; padding-bottom: 4px;">⚠️ Aktualisierte Mitgliedsdaten (${updated.length})</h3>`);
    updated.forEach(m => {
      // Optimierung 5: Komplett auf Array-Pushes umgestellt, um String-Verkettungen zu vermeiden
      html.push('<div style="margin-bottom: 25px; border-left: 4px solid #dd6b20; padding-left: 12px;">');
      html.push(`<span style="font-size: 15px; font-weight: bold; color: #2d3748;">Mitglied: ${m.current.vorname} ${m.current.nachname}</span> <span style="font-size: 13px; color: #718096; margin-left: 10px;">(ID: <code>${m.id}</code>)</span>`);
      html.push(`<table style="${tableStyle} margin-top: 6px; margin-bottom: 5px;">`);
      html.push(`<tr style="background-color: #f8fafc;"><th style="${thStyle} width: 25%;">Feld</th><th style="${thStyle} width: 37.5%;">Alter Wert</th><th style="${thStyle} width: 37.5%;">Neuer Wert</th></tr>`);
      
      // OPTIMIERUNG 4: Schleife für die Tabellenzeilen spart massiven Code-Duplikat-Overhead
      const rows = [
        { label: 'Vorname', key: 'vorname' },
        { label: 'Nachname', key: 'nachname' },
        { label: 'E-Mail', key: 'email' },
        { label: 'Mobile', key: 'mobile' }
      ];
      rows.forEach(r => {
        const isChanged = m.changedFields.includes(r.key);
        const cellStyle = isChanged ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
        html.push(`<tr><td style="${tdStyle} ${cellStyle}">${r.label}</td><td style="${tdStyle} ${cellStyle}">${m.old[r.key] || '-'}</td><td style="${tdStyle} ${cellStyle}">${m.current[r.key] || '-'}</td></tr>`);
      });
      html.push('</table></div>');
    });
  }

  html.push(`<hr style="border: 0; border-top: 1px solid #e2e8f0; margin-top: 40px;"><p style="font-size: 12px; color: #a0aec0;">Generiert am: ${new Date().toLocaleString('de-DE')}</p></div>`);
  const plainBody = `Änderungsbericht Mitgliederliste BC1890\n\n` +
    (added.length > 0 ? `Neu (${added.length}):\n` + added.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}`).join('\n') + `\n\n` : '') +
    (removed.length > 0 ? `Entfernt (${removed.length}):\n` + removed.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}`).join('\n') + `\n\n` : '') +
    (updated.length > 0 ? `Geändert (${updated.length}):\n` + updated.map(m => `- ID: ${m.id}, Änderungen: ${m.textDetails.join(', ')}`).join('\n') + `\n` : '');
  try {
    GmailApp.sendEmail(adminEmail, subject, plainBody, { htmlBody: html.join('') });
  } catch (err) {
    Logger.log('❌ Fehler beim Senden des Änderungsberichts: ' + err.message);
  }
}

// =============================================================================
// 4. ONBOARDING & WILLKOMMENS-SYSTEM (OPTIMIERT)
// =============================================================================

function checkAndWelcomeNewMembers() {
  const modusText = CONFIG.TEST_MODUS_AKTIV ?
    '⚠️ TESTMODUS (AKTIV)' : '🚀 LIVE-BETRIEB';
  Logger.log(`=== STARTE PRÜFUNG AUF NEUE MITGLIEDER [Modus: ${modusText}] ===`);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  const adminEmail = CONFIG.ADMIN_EMAIL;
  if (!sheetId || !adminEmail) {
    Logger.log('❌ KRITISCHER FEHLER: Tabellen-ID oder Admin-E-Mail konnte nicht ermittelt werden.');
    return;
  }

  const welcomedMembersRaw = scriptProperties.getProperty('WELCOMED_MEMBER_IDS');
  const welcomedMemberIds = welcomedMembersRaw ?
    JSON.parse(welcomedMembersRaw) : [];
  const isInitialRun = welcomedMemberIds.length === 0;
  if (isInitialRun) {
    Logger.log('Erster Durchlauf erkannt. Bestehende Mitglieder werden erfasst, ohne E-Mails zu senden.');
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0]; 
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('Keine Mitgliederdaten in der Tabelle gefunden.');
      scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([]));
      return;
    } 
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    // 1. SCHLEIFEN-OPTIMIERUNG: IDs sammeln und Daten validieren in EINEM Durchlauf
    const currentTableIds = new Set();
    const validRows = [];

    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      const memberId = row[0] ? row[0].toString().trim() : '';
      const email = row[3] ? row[3].toString().trim() : '';
      if (memberId) {
        currentTableIds.add(memberId);
        if (email) {
          validRows.push({
            id: memberId,
            vorname: row[1] ? row[1].toString().trim() : '',
            nachname: row[2] ? row[2].toString().trim() : '',
            email: email
          });
        }
      }
    }

    // 2. BEREINIGUNG: Schneller Abgleich dank Set.has()
    const cleanedWelcomedIds = welcomedMemberIds.filter(id => currentTableIds.has(id));
    const removedCount = welcomedMemberIds.length - cleanedWelcomedIds.length;
    if (removedCount > 0) {
      Logger.log(`🧹 BEREINIGUNG: ${removedCount} gelöschte(s) Mitglied(er) aus dem Skript-Gedächtnis entfernt.`);
    }

    // 3. I/O OPTIMIERUNG: PDF einmalig VOR der Schleife holen (spart massiv API-Aufrufe)
    let attachmentBlob = null;
    const fileId = scriptProperties.getProperty('PDF_FILE_ID');
    if (!fileId) {
      throw new Error('Keine gültige Google Drive File ID konfiguriert.');
    }
    try {
      attachmentBlob = DriveApp.getFileById(fileId).getBlob();
    } catch (e) {
      Logger.log(`⚠️ Fehler beim Laden des PDF-Anhangs: ${e.message}. Mails werden ohne Anhang gesendet.`);
    }

    // 4. VERARBEITUNG: Willkommens-Mails senden
    const welcomedSet = new Set(cleanedWelcomedIds);
    let mailsSentCount = 0;

    for (const member of validRows) {
      if (!welcomedSet.has(member.id)) {
        if (!isInitialRun) {
          sendWelcomeMail(member.email, member.vorname, member.nachname, adminEmail, attachmentBlob);
          mailsSentCount++;
        }
        welcomedSet.add(member.id);
      }
    }

    // Zurück in Array konvertieren für Speicherung
    scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([...welcomedSet]));
    Logger.log(`Prüfung abgeschlossen. ${mailsSentCount} neue(s) Mitglied(er) verarbeitet.`);
    
    // 🔥 JETZT MITGLIEDERBERECHTIGUNG AUSFÜHREN:
    if (authorizationChanged && !isInitialRun) {
      Logger.log("⚡ Änderungen an Mitgliedern erkannt. Starte Kalender-Berechtigungen SOFORT...");
      ausfuehrenKalenderSynchronisierung();
    } else {
      Logger.log("ℹ️ Keine Berechtigungsänderungen im Onboarding. Keine Sofort-Kalendersynchronisierung notwendig.");
    }
    
  } catch (e) {
    Logger.log('Fehler im Onboarding-Script: ' + e.message);
  }
}

// Erwartet jetzt den fertigen Blob, um DriveApp-Aufrufe in der Schleife zu verhindern
function sendWelcomeMail(toEmail, vorname, nachname, adminEmail, attachmentBlob) {
  const name = vorname || 'Mitglied';
  
  let finalReceiver = toEmail;
  let finalCc = adminEmail;
  let subject = 'Herzlich willkommen beim Bootsclub 1890! ⛵';
  let testNoticeHtml = '';
  let testNoticePlain = '';

  if (CONFIG.TEST_MODUS_AKTIV) {
    finalReceiver = adminEmail;
    finalCc = '';
    subject = `[TEST-MODUS für: ${toEmail}] Herzlich Willkommen beim Bootsclub 1890! ⛵`;
    testNoticeHtml = `
      <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 12px; margin-bottom: 20px; color: #856404; font-family: sans-serif; border-radius: 4px;">
        ⚠️ <b>SYSTEM-HINWEIS (TEST-MODUS):</b> Diese E-Mail wurde automatisch abgefangen und an den Vorstand umgeleitet.<br>
        <b>Geplanter Empfänger im Live-Betrieb:</b> ${vorname} ${nachname} (&lt;${toEmail}&gt;)
      </div>
    `;
    testNoticePlain = `[⚠️ TEST-MODUS - Geplanter Empfänger im Live-Betrieb: ${vorname} ${nachname} (${toEmail})]\n\n`;
  }
  
  const htmlBody = `
    ${testNoticeHtml}
    Hallo ${name},<br><br>
    Herzlich Willkommen im <b>Bootsclub 1890</b>!<br><br>
    Deine E-Mail-Adresse wurde erfolgreich für unser automatisiertes Reservierungssystem freigeschaltet.<br><br>
    Ab sofort kannst du Bootstermine direkt per E-Mail reservieren.
    <br>
    <b>Im Anhang dieser E-Mail findest du die detaillierte Anleitung als PDF-Datei.</b><br><br>
    Hier sind die wichtigsten Kernpunkte im Überblick:<br>
    • Sende Reservierungen an: <b>${adminEmail}</b>.
    Die E-Mail muss das Wort <b>Reservierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text enthalten.<br><br>
    • Für eine Stornierung sende einfach das Wort <b>Stornierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text (bis max. 24 Stunden vor dem Termin).<br><br>
    Bitte lies dir die angehängte PDF-Anleitung aufmerksam durch, bevor du deine erste Reservierung vornimmst.<br><br>
    Bei Fragen steht dir der Vorstand jederzeit gerne zur Verfügung.<br><br>
    Allzeit gute Fahrt und viel Spass auf dem Wasser!<br><br>
    <b>Dein Vorstand</b><br>
  `;
  const plainBody = `${testNoticePlain}Hallo ${name},\n\nherzlich willkommen beim Bootsclub 1890!\nDeine E-Mail wurde für das Reservierungssystem freigeschaltet.\n\nEine detaillierte Anleitung findest du im Anhang dieser E-Mail als PDF.\n\nBitte sende Reservierungen an ${adminEmail}.\n\nAllzeit gute Fahrt!\nDein Vorstand`;
  try {
    const options = {
      cc: finalCc, 
      replyTo: adminEmail,
      htmlBody: htmlBody
    };
    if (attachmentBlob) {
      options.attachments = [attachmentBlob];
    }

    GmailApp.sendEmail(finalReceiver, subject, plainBody, options);
  } catch (error) {
    Logger.log(`❌ FEHLER beim Senden der Willkommens-Mail: ${error.message}`);
  }
}

/**
 * Gleicht die Google-Kalender-Freigaben mit der aktuellen Mitgliederliste im Sheet ab.
 */
function ausfuehrenKalenderSynchronisierung() {
  Logger.log('🔮 Starte separate Kalender-Synchronisierung...');
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  const kalenderId = CONFIG.CALENDAR_ID || 'primary';

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('🛑 Synchronisierung abgebrochen: Keine Mitglieder im Sheet gefunden.');
      return;
    }

    // --- DYNAMISCHE SPALTENSUCHE ---
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const emailColIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'e-mail' || h.toString().trim().toLowerCase() === 'email');
    const statusColIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'status');
    
    if (emailColIndex === -1) {
      Logger.log('❌ FEHLER: Spalte "E-Mail" oder "Email" konnte in der Kopfzeile nicht gefunden werden.');
      return;
    }
    Logger.log(`ℹ️ E-Mail-Spalte dynamisch auf Index ${emailColIndex} (Spalte ${String.fromCharCode(65 + emailColIndex)}) gefunden.`);

    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const sheetEmailsSet = new Set();

    for (let i = 0; i < dataRange.length; i++) {
      const email = dataRange[i][emailColIndex] ? dataRange[i][emailColIndex].toString().trim().toLowerCase() : '';
      const status = statusColIndex !== -1 && dataRange[i][statusColIndex] ? dataRange[i][statusColIndex].toString().trim().toLowerCase() : '';
      
      // Nur aktive Mitglieder mit gültiger E-Mail-Adresse übernehmen
      if (email && email.includes('@') && status !== 'inaktiv') {
        sheetEmailsSet.add(email);
      }
    }

    Logger.log(`📋 Anzahl gültiger (aktiver) Mitglieder aus der Tabelle: ${sheetEmailsSet.size}`);
    
    const calendar = CalendarApp.getCalendarById(kalenderId);
    if (!calendar) {
      Logger.log(`❌ FEHLER: Kalender mit der ID '${kalenderId}' wurde nicht gefunden.`);
      return;
    }

    // Bestimmt die echte ID (falls 'primary' genutzt wird, holen wir die konkrete Mail-ID des Kalenders)
    const realCalendarId = kalenderId === 'primary' ? calendar.getId() : kalenderId;

    // --- NATIVE API-ABFRAGE DER BERECHTIGUNGEN (ACL) ---
    let currentAclEmails = [];
    try {
      // Holt alle Berechtigungseinträge direkt aus der Google Calendar API
      const aclList = Calendar.Acl.list(realCalendarId);
      
      if (aclList && aclList.items) {
        currentAclEmails = aclList.items
          .filter(item => item.role === 'editor' || item.role === 'owner') // Nur User mit Schreibrechten beachten
          .map(item => item.scope.value.trim().toLowerCase());
      }
    } catch (apiError) {
      Logger.log(`❌ API-FEHLER beim Auslesen der Kalenderrechte: ${apiError.message}`);
      Logger.log("👉 Bitte prüfe, ob du die 'Google Calendar API' links unter 'Dienste' (+ Symbol) im Editor hinzugefügt hast.");
      return;
    }
      
    Logger.log(`📅 Anzahl Personen mit Kalender-Zugriff aktuell: ${currentAclEmails.length}`);

    // --- SYNCHRONISATIONSLAUF ---
    
    // 1. NEUE MITGLIEDER HINZUFÜGEN: Wenn in Tabelle, aber noch nicht im Kalender
    sheetEmailsSet.forEach(email => {
      if (email === CONFIG.ADMIN_EMAIL.toLowerCase()) return; // Admin überspringen
      
      if (!currentAclEmails.includes(email)) {
        try {
          // Native API-Variante: Erstellt eine neue Freigaberegel für den Benutzer
          Calendar.Acl.insert({
            role: 'editor',
            scope: {
              type: 'user',
              value: email
            }
          }, realCalendarId);
          
          Logger.log(`➕ Zugriff ERLAUBT für neues Mitglied: ${email}`);
        } catch (e) {
          // ✅ INTEGRIERTER SCHUTZ: Fängt ungültige Google-Konten ab (Bad Request)
          if (e.message.includes('Bad Request') || e.message.includes('invalid')) {
            Logger.log(`❌ HINWEIS für Admin bei ${email}: Diese Adresse besitzt vermutlich kein Google-Konto oder ist nicht für Google-Dienste registriert. (API-Fehler: Bad Request)`);
          } else {
            Logger.log(`⚠️ Unerwarteter Fehler beim Hinzufügen von ${email}: ${e.message}`);
          }
        }
      }
    });

    // 2. AUSGESCHIEDENE MITGLIEDER ENTFERNEN: Wenn im Kalender, aber nicht mehr in Tabelle
    currentAclEmails.forEach(email => {
      if (email === CONFIG.ADMIN_EMAIL.toLowerCase()) return; // Admin niemals löschen
      
      if (!sheetEmailsSet.has(email)) {
        try {
          // Um ein Recht per API zu löschen, müssen wir zuerst die "ruleId" für diese E-Mail auslesen
          const aclList = Calendar.Acl.list(realCalendarId);
          const userRule = aclList.items.find(item => item.scope.value.trim().toLowerCase() === email);
          
          if (userRule && userRule.id) {
            // Native API-Variante: Löscht die Freigaberegel anhand der eindeutigen Regel-ID
            Calendar.Acl.remove(realCalendarId, userRule.id);
            Logger.log(`➖ Zugriff ENTZOGEN für ausgeschiedenes/inaktives Mitglied: ${email}`);
          } else {
            Logger.log(`ℹ️ Regel-ID für ${email} konnte nicht ermittelt werden. Evtl. bereits entfernt.`);
          }
        } catch (e) {
          Logger.log(`⚠️ Fehler beim Entfernen von ${email}: ${e.message}`);
        }
      }
    });
    
    Logger.log('✅ Kalender-Synchronisierung erfolgreich abgeschlossen!');

  } catch (error) {
    Logger.log(`❌ KRITISCHER FEHLER in ausfuehrenKalenderSynchronisierung: ${error.message}`);
  }
}

// =============================================================================
// 5. HILFSFUNKTIONEN & STORNIERUNGSLOGIK - OPTIMIERT
// =============================================================================

function executeCancellation(data, userId, thread, message) {
  const memberData = getAuthorizedUserData(userId);
  
  if (!memberData) {
    GmailApp.sendEmail(userId, 'Löschen der Buchung abgelehnt', `Deine E-Mail-Adresse (${userId}) ist nicht im System hinterlegt.`, { replyTo: CONFIG.ADMIN_EMAIL });
    return false;
  }
  
  data.name = memberData.name;
  const jetzt = new Date();  
  const slotTime = data.slot === 'vormittag' ? CONFIG.SLOT_VORMITTAG : CONFIG.SLOT_NACHMITTAG;
  const terminStartZeit = new Date(data.parsedDate); 
  const [sh, sm] = slotTime.start.split(':');
  terminStartZeit.setHours(sh, sm, 0, 0);  

  // 24 Stunden Frist berechnen
  const stornierungsFrist = new Date(terminStartZeit.getTime() - (24 * 60 * 60 * 1000));
  if (jetzt > stornierungsFrist) {
    let fehlerGrund = terminStartZeit < jetzt ? 'Der Termin liegt in der Vergangenheit.' : `Die Frist für eine automatische Stornierung (24h vor Beginn) ist abgelaufen.`;
    GmailApp.sendEmail(userId, 'Löschen der Buchung abgelehnt', `Hallo ${data.name},\n\n❌ Grund: ${fehlerGrund}`, { replyTo: CONFIG.ADMIN_EMAIL });
    return false; 
  }

  const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
  const terminEndZeit = new Date(terminStartZeit); 
  const [eh, em] = slotTime.end.split(':');
  terminEndZeit.setHours(eh, em, 0, 0); 

  const events = calendar.getEvents(terminStartZeit, terminEndZeit);
  const userEvent = events.find(e => (e.getDescription() || '').includes(`Mitglieder-ID: ${memberData.id}`));

  if (userEvent) {
    if (userEvent.getTitle().toUpperCase().includes('JOKER')) {
      GmailApp.sendEmail(userId, 'Stornierung fehlgeschlagen', `❌ Joker-Termine können nicht automatisch storniert werden. Bitte wende dich an den Admin.`, { replyTo: CONFIG.ADMIN_EMAIL });
      return false;
    }

    userEvent.deleteEvent();
    // OPTIMIERUNG 1: Tippfehler im Betreff korrigiert ("Bestätigung" statt "BestBTigung")
    GmailApp.sendEmail(userId, 'Bestätigung: Termin freigegeben', `Deine Reservierung für den ${formatDateDDMMYYYY(data.parsedDate)} wurde erfolgreich storniert.`, { replyTo: CONFIG.ADMIN_EMAIL });
    return true;
  } else {
    GmailApp.sendEmail(userId, 'Stornierung fehlgeschlagen', `❌ Es wurde kein passender aktiver Termin für dich an diesem Tag gefunden.`, { replyTo: CONFIG.ADMIN_EMAIL });
    return false;
  }
}

function sendConfirmationEmail(to, event, data, thread) {
  const subject = 'Buchung bestätigt: ' + event.getTitle();
  
  // 1. Der reine Text-Body (Fallback) erhält ebenfalls saubere Emojis
  const plainBody = `Hallo ${data.name},\n\ndein Termin wurde erfolgreich eingetragen:\n\n📅 Datum: ${formatDateDDMMYYYY(data.parsedDate)}\n🕒 Slot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}\n\nDein Vorstand`;
  
  // 2. Der HTML-Body (erzwingt die korrekte Codierung im Mail-Client)
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
      <p>Hallo ${data.name},</p>
      <p>dein Termin wurde erfolgreich eingetragen:</p>
      <p style="line-height: 1.6;">
        📅 <b>Datum:</b> ${formatDateDDMMYYYY(data.parsedDate)}<br>
        🕒 <b>Slot:</b> ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}
      </p>
      <p>Dein Vorstand</p>
    </div>
  `;

  // 3. Erweiterte Optionen für den Entwurf (Draft) vorbereiten
  const advancedOptions = { 
    replyTo: CONFIG.ADMIN_EMAIL, 
    htmlBody: htmlBody 
  };

  try {
    // Direktes Senden der Mail mit den erweiterten Optionen
    GmailApp.sendEmail(to, subject, plainBody, advancedOptions);
  } catch (error) {
    Logger.log(`⚠️ Direktes Senden fehlgeschlagen, erstelle Entwurf... Fehler: ${error.message}`);
    if (thread) {
      try {
        // Falls das direkte Senden fehlschlägt, wird die Antwort im Thread als sauberer Entwurf abgelegt
        thread.createDraftReply(plainBody, advancedOptions);
      } catch (draftError) {
        Logger.log(`❌ Fehler beim Erstellen des Entwurfs: ${draftError.message}`);
      }
    }
  }
}

function sendRejectionEmail(to, reason, thread) {
  const subject = 'Buchung abgelehnt';
  const body = `Hallo,\n\nleider konnte deine Reservierung nicht angenommen werden:\n\n❌ Grund: ${reason}`;
  try {
    GmailApp.sendEmail(to, subject, body, { replyTo: CONFIG.ADMIN_EMAIL });
  } catch (error) {
    if (thread) thread.createDraftReply(`Ablehnungsgrund: ${reason}`);
  }
}

// Globaler Cache zur Vermeidung mehrfacher Tabellen-I/O-Aufrufe während desselben Skript-Laufs
let memberDataCache_ = null;

function getAuthorizedUserData(email) {
  const searchEmail = email.trim().toLowerCase();
  
  // OPTIMIERUNG 2: Cache-Abfrage spart wertvolle Millisekunden bei Schleifendurchläufen
  if (memberDataCache_ && memberDataCache_[searchEmail]) {
    return memberDataCache_[searchEmail];
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  let ss, sheet;
  try { if (sheetId) { ss = SpreadsheetApp.openById(sheetId); sheet = ss.getSheets()[0]; } } catch (e) {}

  // Automatisches Erstellen der Tabelle falls gelöscht oder nicht vorhanden
  if (!ss || !sheet || sheet.getLastRow() === 0) {
    try {
      const folderName = "Google Kalender Reservierungssystem";
      let targetFolder = DriveApp.getFoldersByName(folderName).hasNext() ? DriveApp.getFoldersByName(folderName).next() : DriveApp.createFolder(folderName);

      ss = SpreadsheetApp.create('Mitgliederliste');
      sheet = ss.getSheets()[0];
      sheetId = ss.getId();
      const file = DriveApp.getFileById(sheetId);
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);

      sheet.appendRow(["Mitglieder ID", "Vorname", "Name", "E-Mail", "Mobile"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
      sheet.appendRow(["BJB-000", "Vorstand", "", CONFIG.ADMIN_EMAIL, ""]);
      sheet.autoResizeColumns(1, 5);

      scriptProperties.setProperty('SHEET_CONFIG_ID', sheetId);
    } catch (err) { return null;
    }
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    // Cache initialisieren
    memberDataCache_ = {};
    let foundUserData = null;

    for (let i = 0; i < dataRange.length; i++) {
      const currentEmail = dataRange[i][3] ? dataRange[i][3].toString().trim().toLowerCase() : '';
      if (!currentEmail) continue;

      const userObj = {
        id: dataRange[i][0] ? dataRange[i][0].toString().trim() : 'Keine ID',  
        name: `${dataRange[i][1] || ''} ${dataRange[i][2] || ''}`.trim() || currentEmail,    
        mobile: dataRange[i][4] ? dataRange[i][4].toString().trim() : 'Nicht hinterlegt'
      };
      // Alle Mitglieder in den Cache schreiben für zukünftige Suchen im selben Lauf
      memberDataCache_[currentEmail] = userObj;
      if (currentEmail === searchEmail) {
        foundUserData = userObj;
      }
    }
    return foundUserData;
  } catch (e) { return null;
  }
}

function getCurrentSeasonStart() {
  return new Date(new Date().getFullYear(), 0, 1);
}

function createGmailLabelStructure(fullLabelPath) {
  const parts = fullLabelPath.split('/');
  let currentPath = '';
  let finalLabel = null;
  for (let i = 0; i < parts.length; i++) {
    currentPath = currentPath ? currentPath + '/' + parts[i] : parts[i];
    let label = GmailApp.getUserLabelByName(currentPath) || GmailApp.createLabel(currentPath);
    if (i === parts.length - 1) finalLabel = label;
  }
  return finalLabel;
}

// =============================================================================
// 6. ZEITGESTEUERTE AUTOMATISIERUNGEN & ZENTRALES SETUP
// =============================================================================

function sendDailyReservationReminders() {
  Logger.log("=== STARTE TÄGLICHE ERINNERUNGS-PRÜFUNG ===");
  const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    
  if (!calendar) {
    Logger.log("❌ KRITISCHER FEHLER: Kalender konnte nicht geladen werden.");
    return;  
  }

  // OPTIMIERUNG 3: Exakte "Morgen"-Zeitspanne berechnen
  const tomorrowStart = new Date();
  tomorrowStart.setDate(tomorrowStart.getDate() + 1);
  tomorrowStart.setHours(0, 0, 0, 0);
  
  const tomorrowEnd = new Date(tomorrowStart);
  tomorrowEnd.setHours(23, 59, 59, 999);
  
  const events = calendar.getEvents(tomorrowStart, tomorrowEnd);
  events.forEach(event => {
    const desc = event.getDescription() || "";
    const emailMatch = desc.match(/Kontakt:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch && emailMatch[1]) {
      const slotName = event.getStartTime().getHours() === 8 ? "Vormittag (08:00 - 14:00)" : "Nachmittag (14:00 - 20:00)";
      let body = `Hallo!\n\nAutomatische Erinnerung für deine Reservierung morgen:\n📅 Datum: ${formatDateDDMMYYYY(tomorrowStart)}\n⏰ Slot: ${slotName}\n\nViel Spass mit dem Boot!`;
      GmailApp.sendEmail(emailMatch[1].trim(), `Erinnerung: Deine Boot Buchung für morgen!`, body);
    }
  });
}

function ensureInitialSheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  let sheetExistsAndOpens = false;

  // 1. Prüfen, ob die ID existiert UND die Tabelle geöffnet werden kann
  if (sheetId) {
    try {
      SpreadsheetApp.openById(sheetId);
      sheetExistsAndOpens = true;
      Logger.log(`✅ Bestehendes Google Sheet erfolgreich verifiziert (ID: ${sheetId}).`);
    } catch (e) {
      Logger.log(`⚠️ Warnung: Sheet-ID '${sheetId}' existiert, Datei konnte aber nicht geöffnet werden (evtl. gelöscht). Erstelle neues Sheet...`);
      sheetExistsAndOpens = false;
    }
  }

  // 2. Wenn kein Sheet existiert oder das alte nicht geöffnet werden konnte -> Neu erstellen
  if (!sheetExistsAndOpens) {
    try {
      const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
      const sheetName = `${CONFIG.RESERVATION_FOLDER_NAME || 'Bootsclub'} - Mitgliederdatenbank (Erstellt am ${dateStr})`;
      
      // Neues Google Sheet erstellen (landet zuerst im Root-Verzeichnis)
      const newSpreadsheet = SpreadsheetApp.create(sheetName);
      sheetId = newSpreadsheet.getId();
      
      // Die Kopfzeile direkt initialisieren
      const sheet = newSpreadsheet.getSheets()[0];
      sheet.setName('Mitglieder');
      
      const headers = ['ID', 'Vorname', 'Nachname', 'E-Mail', 'Status'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
      
      // --- NEU: IN SPREZIFISCHEN ORDNER VERSCHIEBEN ---
      const folderName = "Google Kalender Reservierungssystem";
      const folders = DriveApp.getFoldersByName(folderName);
      let targetFolder;
      
      if (folders.hasNext()) {
        // Ordner existiert bereits
        targetFolder = folders.next();
        Logger.log(`📂 Zielordner "${folderName}" gefunden.`);
      } else {
        // Ordner existiert noch nicht und wird neu angelegt
        targetFolder = DriveApp.createFolder(folderName);
        Logger.log(`📁 Zielordner "${folderName}" existierte nicht und wurde neu erstellt.`);
      }
      
      // Datei im Zielordner hinzufügen und aus dem Root-Verzeichnis entfernen
      const file = DriveApp.getFileById(sheetId);
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      
      // 3. ID in den ScriptProperties aktualisieren
      scriptProperties.setProperty('SHEET_CONFIG_ID', sheetId);
      Logger.log(`🆕 Initiales Google Sheet erfolgreich im Ordner "${folderName}" erstellt und ID aktualisiert!`);
      Logger.log(`🔗 Neue Sheet-ID: ${sheetId}`);
      
    } catch (createError) {
      Logger.log(`❌ KRITISCHER FEHLER bei der Neuerstellung des Sheets im Ordner: ${createError.message}`);
    }
  }
  
  return sheetId;
}

function fetchAndSyncAnleitungPDF() {
  Logger.log("🔄 Synchronisiere PDF-Anleitung von GitHub...");
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  if (!sheetId) return;
  
  const sheetFile = DriveApp.getFileById(sheetId);
  const parents = sheetFile.getParents();
  const targetFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  let targetUrl = PDF_SOURCE_URL;
  
  // OPTIMIERUNG 4: Korrekte Übersetzung in die GitHub-RAW-Domain zur Vermeidung von PDF-Korruption
  if (targetUrl.includes('github.com')) {
    targetUrl = targetUrl
      .replace('github.com', 'raw.githubusercontent.com')
      .replace('/blob/', '/');
  }
  
  Logger.log("📥 Rufe URL ab: " + targetUrl);
  const response = UrlFetchApp.fetch(targetUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    Logger.log("❌ Fehler beim Abrufen der PDF von GitHub: " + response.getResponseCode());
    return;
  }
  
  const pdfBlob = response.getBlob().setContentType("application/pdf").setName("Anleitung Bootsreservation.pdf");
  const fileName = "Anleitung Bootsreservation.pdf";
  const files = targetFolder.getFilesByName(fileName);
  let localFile = files.hasNext() ? files.next() : null;
  
  if (localFile) {
    const headers = response.getHeaders();
    const remoteLastModifiedStr = headers["Last-Modified"] || headers["last-modified"];
    let shouldUpdate = true;
    if (remoteLastModifiedStr) {
      const remoteDate = new Date(remoteLastModifiedStr);
      const localDate = new Date(localFile.getLastUpdated());
      if (remoteDate <= localDate) {
        shouldUpdate = false;
        Logger.log("ℹ️ Lokale PDF ist auf dem neuesten Stand.");
      }
    }
    
    if (shouldUpdate) {
      Logger.log("🔄 Lokale PDF veraltet. Ersetze Datei...");
      localFile.setTrashed(true);
      const newFile = targetFolder.createFile(pdfBlob);
      scriptProperties.setProperty('PDF_FILE_ID', newFile.getId());
    } else {
      scriptProperties.setProperty('PDF_FILE_ID', localFile.getId());
    }
    
  } else {
    Logger.log("📥 PDF existiert lokal nicht. Erstelle neue Datei...");
    const newFile = targetFolder.createFile(pdfBlob);
    scriptProperties.setProperty('PDF_FILE_ID', newFile.getId());
  }
}

function setupTriggers() {
  Logger.log('========================================================================');
  Logger.log('🚀 STARTE CENTRAL SYSTEM SETUP...');
  Logger.log('========================================================================');

  ensureInitialSheet();
  try {
    fetchAndSyncAnleitungPDF();
  } catch(pdfError) {
    Logger.log("⚠️ Warnung beim PDF-Sync: " + pdfError.toString());
  }

  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(t => ScriptApp.deleteTrigger(t));

  // Trigger-Definitionen
  ScriptApp.newTrigger('processReservationEmails').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('sendDailyReservationReminders').timeBased().everyDays(1).atHour(4).create();
  ScriptApp.newTrigger('importExcelToSheets').timeBased().everyMinutes(10).create();
  
  ['Reservierung/Neu', 'Reservierung/Erledigt', 'Reservierung/Abgelehnt', CONFIG.EXCEL_TARGET_LABEL].forEach(label => {
    if (!GmailApp.getUserLabelByName(label)) createGmailLabelStructure(label);
  });
  Logger.log('========================================================================');
  Logger.log('🎉 INTEGRIERTES GESAMT-SETUP ERFOLGREICH!');
  Logger.log('========================================================================');
}

// =============================================================================
// 7. ENTWICKLER-WERKZEUGE (MAINTENANCE)
// =============================================================================

function setEarliestBookingDate() {
  const zielDatum = '01.04.' + new Date().getFullYear();
  PropertiesService.getScriptProperties().setProperty('EARLIEST_BOOKING_DATE', zielDatum);
  Logger.log(`Frühestmögliches Startdatum wurde erfolgreich auf den ${zielDatum} gesetzt!`);
}

function resetWelcomeDatabase() {
  PropertiesService.getScriptProperties().deleteProperty('WELCOMED_MEMBER_IDS');
  Logger.log('Onboarding-Datenbank zurückgesetzt.');
}

function resetTrackingSnapshot() {
  PropertiesService.getScriptProperties().deleteProperty('MEMBER_LIST_SNAPSHOT');
  Logger.log('Tracking-Schnappschuss wurde erfolgreich gelöscht.');
}

// Helper zum Parsen europäischer Daten (falls im restlichen Skript benötigt)
function parseEuropeanDate(dateStr) {
  const parts = dateStr.split(/[./-]/);
  if (parts.length === 3) {
    return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  return new Date(dateStr);
}

// Helper zum Formatieren von Daten (DD.MM.YYYY)
function formatDateDDMMYYYY(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yyyy");
}
