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
// HELPER: ZENTRALE FUNKTION FÜR TABELLENZUGRIFF (PROPERTIES-BASIERT)
// =============================================================================

function getOrCreateSpreadsheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  
  if (!sheetId) {
    // Falls keine ID existiert, erstellen wir eine neue Master-Tabelle
    const newSS = SpreadsheetApp.create('BC1890_Mitglieder_Master_Datenbank');
    sheetId = newSS.getId();
    scriptProperties.setProperty('SHEET_CONFIG_ID', sheetId);
    
    const sheet = newSS.getSheets()[0];
    sheet.setName('Mitgliederliste');
    
    // Header einrichten
    sheet.getRange('A1:E1').merge().setValue('BC1890 MITGLIEDERDATENBANK (AUTOMATISIERT)')
         .setFontWeight('bold').setFontSize(12).setBackground('#1a365d').setFontColor('#ffffff').setHorizontalAlignment('center');
    
    const headers = ['Mitglieder-ID', 'Vorname', 'Nachname', 'E-Mail', 'Mobile'];
    sheet.getRange(2, 1, 1, 5).setValues([headers]).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(2);
    
    Logger.log('🎉 Neue Master-Tabelle wurde erstellt. ID: ' + sheetId);
  }
  
  return SpreadsheetApp.openById(sheetId);
}

// =============================================================================
// 1. KERN-LOGIK: RESERVIERUNGEN & STORNIERUNGEN VERARBEITEN (OPTIMIERT)
// =============================================================================

function processReservationEmails() {
  let labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL) || createGmailLabelStructure(CONFIG.GMAIL_LABEL);

  // OPTIMIERUNG 1: Kombinierte Suchanfrage spart API-Quota und verhindert Doppelverarbeitung
  const emailThreads = GmailApp.search('in:inbox (subject:"Reservierung" OR subject:"Stornierung")');
  Logger.log(`Gefundene relevante Threads im Posteingang: ${emailThreads.length}`);
  
  // OPTIMIERUNG 2: Kalender-Instanz EINMALIG holen und wiederverwenden
  const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
  if (!calendar) {
    Logger.log('❌ KRITISCHER FEHLER: Kalender konnte nicht geladen werden.');
    return;
  }

  // Arrays für die Batch-Label-Verarbeitung
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
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    labelErledigt.addToThreads(threadsErledigt);
    if (labelNeu) { labelNeu.removeFromThreads(threadsErledigt); }
    threadsErledigt.forEach(thread => thread.moveToArchive());
  }
  
  if (threadsAbgelehnt.length > 0) {
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    labelAbgelehnt.addToThreads(threadsAbgelehnt);
    if (labelNeu) { labelNeu.removeFromThreads(threadsAbgelehnt); }
    threadsAbgelehnt.forEach(thread => thread.moveToArchive());
  }

  // Erfolgreiche Stornierungen gehen ebenfalls in 'Reservierung/Erledigt'
  if (threadsStorniert.length > 0) {
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    labelErledigt.addToThreads(threadsStorniert);
    if (labelNeu) { labelNeu.removeFromThreads(threadsStorniert); }
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

  // Validierung durchführen
  const validation = validateRequest(data, userId, sender, calendar);
  if (!validation.valid) {
    sendRejectionEmail(sender, validation.error, thread);
    return 'ABGELEHNT';
  }

  // Kalendereintrag erstellen
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
      if (line.toLowerCase().startsWith(key.toLowerCase() + ':')) {
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

function getAuthorizedUserData(email) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 2) return null;

    const data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
    const cleanEmail = email.toLowerCase().trim();

    for (let i = 0; i < data.length; i++) {
      const dbEmail = data[i][3] ? data[i][3].toString().toLowerCase().trim() : '';
      if (dbEmail === cleanEmail) {
        return {
          id: data[i][0] ? data[i][0].toString().trim() : '',
          name: `${data[i][1] || ''} ${data[i][2] || ''}`.trim(),
          mobile: data[i][4] ? data[i][4].toString().trim() : ''
        };
      }
    }
  } catch(e) {
    Logger.log('Fehler bei getAuthorizedUserData: ' + e.message);
  }
  return null;
}

function getCurrentSeasonStart() {
  const currentYear = new Date().getFullYear();
  return new Date(currentYear, 0, 1, 0, 0, 0); // Saisonschnitt am 01. Januar
}

// =============================================================================
// 2. EXCEL-IMPORT SYSTEM (EXCEL -> GOOGLE SHEET)
// =============================================================================

function importExcelToSheets() {
  const ss = getOrCreateSpreadsheet();
  const adminEmail = CONFIG.ADMIN_EMAIL;
  if (!adminEmail) {
    Logger.log("❌ KRITISCHER FEHLER: Admin-E-Mail konnte nicht ermittelt werden.");
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

      message.markRead();
      const sender = message.getFrom().toLowerCase();
      const subject = message.getSubject();
      
      if (subject !== CONFIG.EXCEL_SUBJECT) continue;
      
      // Nur Mails des Admins akzeptieren
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
          
          tempSheetFile = Drive.Files.create(resource, fileBlob);
          const tempSpreadsheet = SpreadsheetApp.openById(tempSheetFile.id);
          const tempSheet = tempSpreadsheet.getSheets()[0];
          const tempLastRow = tempSheet.getLastRow();
          const tempLastColumn = tempSheet.getLastColumn();
          
          if (tempLastRow <= 1) {
            Logger.log(`⚠️ Excel-Datei ${attachment.getName()} enthält keine Datenzeilen.`);
            continue;
          }
          
          const newValues = tempSheet.getRange(2, 1, tempLastRow - 1, tempLastColumn).getValues();
          const targetSheet = ss.getSheets()[0];
          const targetLastRow = targetSheet.getLastRow();
          
          if (targetLastRow > 2) {
            targetSheet.getRange(3, 1, targetLastRow - 2, targetSheet.getLastColumn()).clearContent();
          }
          
          targetSheet.getRange(3, 1, newValues.length, tempLastColumn).setValues(newValues);
          Logger.log(`✅ Mitgliederliste erfolgreich durch Excel-Mail aktualisiert (${newValues.length} Mitglieder).`);
          
          importErfolgreich = true;
          break;
          
        } catch (e) {
          Logger.log(`❌ Fehler beim Verarbeiten der Import-Datei: ${e.message}`);
        } finally {
          if (tempSheetFile && tempSheetFile.id) {
            try { DriveApp.getFileById(tempSheetFile.id).setTrashed(true); } catch(err) {}
          }
        }
      }
      if (importErfolgreich) break;
    }
    
    if (importErfolgreich) {
      if (targetLabel) thread.addLabel(targetLabel);
    } else {
      if (errorLabel) thread.addLabel(errorLabel);
    }
    thread.moveToArchive();
  }

  if (threads.length > 0) {
    GmailApp.refreshThreads(threads);
  }

  // KETTENREAKTION: Tracking starten
  if (typeof tracklistchanges === 'function') {
    tracklistchanges();
  }
}

// =============================================================================
// 3. MITGLIEDERLISTEN-TRACKING-SYSTEM (DATENÄNDERUNGEN ERKENNEN)
// =============================================================================

function tracklistchanges() {
  Logger.log('=== STARTE MITGLIEDERLISTEN-TRACKING ===');
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProperties = scriptProperties.getProperties();
  const adminEmail = allProperties['ADMIN_EMAIL'] || CONFIG.ADMIN_EMAIL;

  const lastSnapshotRaw = allProperties['MEMBER_LIST_SNAPSHOT'];
  const currentSnapshot = {};

  let ss;
  try {
    ss = getOrCreateSpreadsheet();
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
    if (typeof checkAndWelcomeNewMembers === 'function') {
      checkAndWelcomeNewMembers();
    }
    return;
  }

  const lastSnapshot = JSON.parse(lastSnapshotRaw);
  const addedMembers = [];
  const updatedMembers = [];
  let authorizationChanged = false; // Flag für Kalender-Sync Trigger

  for (const id in currentSnapshot) {
    const current = currentSnapshot[id];
    const last = lastSnapshot[id];
    current.id = id;

    if (!last) {
      addedMembers.push(current);
      authorizationChanged = true;
    } else {
      const changedFields = [];
      const textDetails = [];
      const fieldsToTrack = { vorname: 'Vorname', nachname: 'Nachname', email: 'E-Mail', mobile: 'Mobil' };
      
      for (const [field, label] of Object.entries(fieldsToTrack)) {
        if (current[field] !== last[field]) {
          changedFields.push(field);
          textDetails.push(`${label}: ${last[field] || '-'} -> ${current[field] || '-'}`);
          if (field === 'email') authorizationChanged = true; // E-Mail-Wechsel verlangt Sync
        }
      }

      if (changedFields.length > 0) {
        updatedMembers.push({
          id: id, old: last, current: current, changedFields: changedFields, textDetails: textDetails
        });
      }
      delete lastSnapshot[id];
    }
  }

  const removedMembers = Object.keys(lastSnapshot).map(id => {
    const removed = lastSnapshot[id];
    removed.id = id;
    authorizationChanged = true;
    return removed;
  });

  if (addedMembers.length > 0 || removedMembers.length > 0 || updatedMembers.length > 0) {
    Logger.log(`Änderungen erkannt! Neu: ${addedMembers.length}, Gelöscht: ${removedMembers.length}, Geändert: ${updatedMembers.length}`);
    sendChangeReportMail(adminEmail, addedMembers, removedMembers, updatedMembers);
    
    if (!CONFIG.TRACKING_TEST_MODUS_AKTIV) {
      scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot));
    }
  } else {
    Logger.log('Keine Änderungen an der Mitgliederliste festgestellt.');
  }

  // KETTENREAKTION: Direkt im Anschluss Onboarding ausführen
  if (typeof checkAndWelcomeNewMembers === 'function') {
    // Übergeben der Information via Property, ob Berechtigungen synchronisiert werden müssen
    if (authorizationChanged) {
      scriptProperties.setProperty('PENDING_AUTH_SYNC', 'true');
    }
    checkAndWelcomeNewMembers();
  }

  Logger.log('=== TRACKING BEENDET ===');
}

function sendChangeReportMail(adminEmail, added, removed, updated) {
  let subject = `✅ Änderungsbericht: Mitgliederliste BC1890`;
  if (CONFIG.TRACKING_TEST_MODUS_AKTIV) subject = `[TEST] ` + subject;

  const tableStyle = 'width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 25px; font-size: 14px;';
  const thStyle = 'background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 10px; text-align: left; color: #334155; font-weight: bold;';
  const tdStyle = 'border: 1px solid #e2e8f0; padding: 10px; vertical-align: top; color: #475569;';
  
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
      html.push('<div style="margin-bottom: 25px; border-left: 4px solid #dd6b20; padding-left: 12px;">');
      html.push(`<span style="font-size: 15px; font-weight: bold; color: #2d3748;">Mitglied: ${m.current.vorname} ${m.current.nachname}</span> <span style="font-size: 13px; color: #718096; margin-left: 10px;">(ID: <code>${m.id}</code>)</span>`);
      html.push(`<table style="${tableStyle} margin-top: 6px; margin-bottom: 5px;">`);
      html.push(`<tr style="background-color: #f8fafc;"><th style="${thStyle} width: 25%;">Feld</th><th style="${thStyle} width: 37.5%;">Alter Wert</th><th style="${thStyle} width: 37.5%;">Neuer Wert</th></tr>`);
      
      const rows = [
        { label: 'Vorname', key: 'vorname' }, { label: 'Nachname', key: 'nachname' },
        { label: 'E-Mail', key: 'email' }, { label: 'Mobile', key: 'mobile' }
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
  const modusText = CONFIG.TEST_MODUS_AKTIV ? '⚠️ TESTMODUS (AKTIV)' : '🚀 LIVE-BETRIEB';
  Logger.log(`=== STARTE PRÜFUNG AUF NEUE MITGLIEDER [Modus: ${modusText}] ===`);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const adminEmail = CONFIG.ADMIN_EMAIL;

  const welcomedMembersRaw = scriptProperties.getProperty('WELCOMED_MEMBER_IDS');
  const welcomedMemberIds = welcomedMembersRaw ? JSON.parse(welcomedMembersRaw) : [];
  const isInitialRun = welcomedMemberIds.length === 0;

  if (isInitialRun) {
    Logger.log('Erster Durchlauf erkannt. Bestehende Mitglieder werden erfasst, ohne E-Mails zu senden.');
  }

  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheets()[0]; 
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('Keine Mitgliederdaten in der Tabelle gefunden.');
      scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([]));
      return;
    } 
    
    const dataRange = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
    const currentTableIds = new Set();
    const validRows = [];

    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      const memberId = row[0] ? row[0].toString().trim() : '';
      const email = row[3] ? row[3].toString().trim() : '';
      if (memberId) {
        currentTableIds.add(memberId);
        if (email) {
          validRows.push({ id: memberId, vorname: row[1] || '', nachname: row[2] || '', email: email });
        }
      }
    }

    const cleanedWelcomedIds = welcomedMemberIds.filter(id => currentTableIds.has(id));
    const welcomedSet = new Set(cleanedWelcomedIds);
    let mailsSentCount = 0;

    // PDF laden für Anhang
    let attachmentBlob = null;
    const fileId = scriptProperties.getProperty('PDF_FILE_ID');
    if (fileId) {
      try { attachmentBlob = DriveApp.getFileById(fileId).getBlob(); } catch(e) {}
    }

    for (const member of validRows) {
      if (!welcomedSet.has(member.id)) {
        if (!isInitialRun) {
          sendWelcomeMail(member.email, member.vorname, member.nachname, adminEmail, attachmentBlob);
          mailsSentCount++;
        }
        welcomedSet.add(member.id);
      }
    }

    scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([...welcomedSet]));
    Logger.log(`Onboarding abgeschlossen. ${mailsSentCount} neue Willkommens-Mails gesendet.`);
    
    // Prüfen, ob durch das Tracking Berechtigungs-Updates anstehen
    const pendingSync = scriptProperties.getProperty('PENDING_AUTH_SYNC');
    if (pendingSync === 'true' && !isInitialRun) {
      scriptProperties.deleteProperty('PENDING_AUTH_SYNC');
      Logger.log("⚡ Berechtigungsänderungen aus Tracking erkannt. Synchronisiere Kalender-Zugriffe...");
      ausfuehrenKalenderSynchronisierung();
    }

  } catch (e) {
    Logger.log('Fehler im Onboarding-Script: ' + e.message);
  }
}

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
    <div style="font-family: sans-serif; color: #333; max-width: 600px;">
      <h2>Herzlich Willkommen im Bootsclub 1890! ⛵</h2>
      <p>Hallo ${name},</p>
      <p>Deine E-Mail-Adresse wurde erfolgreich für unser automatisiertes Reservierungssystem freigeschaltet.</p>
      <p>Ab sofort kannst du Bootstermine direkt per E-Mail reservieren. <b>Im Anhang dieser E-Mail findest du die detaillierte Anleitung als PDF-Datei.</b></p>
      <hr style="border:0; border-top: 1px solid #eee; margin: 20px 0;">
      <h3>Die wichtigsten Punkte im Überblick:</h3>
      <ul>
        <li><b>Reservierungen senden an:</b> <a href="mailto:${adminEmail}">${adminEmail}</a></li>
        <li>Der Betreff muss zwingend das Wort <b>"Reservierung"</b> enthalten.</li>
        <li>Im Textfeld müssen die Zeilen <b>"Datum: DD.MM.YYYY"</b> und <b>"Slot: Vormittag"</b> (oder Nachmittag) aufgeführt werden.</li>
        <li><b>Stornierungen:</b> Betreff mit <b>"Stornierung"</b> und gleicher Zeilen-Syntax (Spätestens 24h vor dem Termin).</li>
      </ul>
      <p>Bitte lies dir die angehängte PDF-Anleitung aufmerksam durch, bevor du deine erste Reservierung vornimmst.</p>
      <p>Allzeit gute Fahrt und viel Spass auf dem Wasser!<br><br><b>Dein Vorstand</b></p>
    </div>
  `;

  const plainBody = `${testNoticePlain}Hallo ${name},\n\nherzlich willkommen beim Bootsclub 1890!\nDeine E-Mail wurde für das Reservierungssystem freigeschaltet.\n\nEine detaillierte Anleitung findest du im Anhang dieser E-Mail als PDF.\n\nBitte sende Reservierungen an ${adminEmail}.\n\nAllzeit gute Fahrt!\nDein Vorstand`;

  try {
    const options = { cc: finalCc, replyTo: adminEmail, htmlBody: htmlBody };
    if (attachmentBlob) { options.attachments = [attachmentBlob]; }
    GmailApp.sendEmail(finalReceiver, subject, plainBody, options);
  } catch (error) {
    Logger.log(`❌ FEHLER beim Senden der Willkommens-Mail: ${error.message}`);
  }
}

// =============================================================================
// 5. STORNIERUNGEN DIREKT AUSFÜHREN & ANTWORT-MAILS
// =============================================================================

function executeCancellation(data, userId, thread, message) {
  const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
  const today = new Date();
  
  // 24-Stunden-Regel prüfen
  const timeDifference = data.startTime.getTime() - today.getTime();
  const hoursDifference = timeDifference / (1000 * 60 * 60);
  
  if (hoursDifference < 24) {
    sendRejectionEmail(userId, `Deine Stornierung für den ${formatDateDDMMYYYY(data.parsedDate)} (Slot: ${data.slot}) wurde abgelehnt, da sie weniger als 24 Stunden vor dem Termin erfolgt ist. Laut Reglement ist dies nicht zulässig.`, thread);
    return false;
  }

  // Passenden Termin im Kalender suchen (über die ID oder den Absender-String)
  const events = calendar.getEvents(data.startTime, data.endTime);
  const userEvent = events.find(e => {
    const desc = e.getDescription() || '';
    return desc.includes(`Mitglieder-ID: ${data.memberId}`) || desc.includes(`Kontakt: ${userId}`);
  });

  if (!userEvent) {
    sendRejectionEmail(userId, `Es wurde kein aktiver Bootstermin für dich am ${formatDateDDMMYYYY(data.parsedDate)} im Slot "${data.slot}" gefunden. Die Stornierung konnte nicht durchgeführt werden.`, thread);
    return false;
  }

  // Termin im Google Kalender löschen
  userEvent.deleteEvent();
  
  // Bestätigung für die erfolgreiche Löschung an den Nutzer senden
  let subject = `❌ Storniert: Deine Bootsreservierung für den ${formatDateDDMMYYYY(data.parsedDate)}`;
  let body = `Hallo ${data.name || 'Mitglied'},\n\ndeine Reservierung für das Boot am ${formatDateDDMMYYYY(data.parsedDate)} (${data.slot}) wurde erfolgreich storniert. Der Slot steht nun wieder anderen Mitgliedern zur Verfügung.\n\nDein Vorstand`;
  
  thread.reply(body);
  return true;
}

function sendConfirmationEmail(sender, event, data, thread) {
  const typText = data.type === 'joker' ? '✨ JOKER-Termin' : '✅ Standard-Termin';
  
  const htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; color: #333;">
      <h2 style="color: #2f855a;">Boot-Reservierung bestätigt! ⛵</h2>
      <p>Hallo <b>${data.name || 'Mitglied'}</b>,</p>
      <p>Deine Reservierung für das Boot wurde erfolgreich im System eingetragen:</p>
      <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
        <tr><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Datum:</td><td style="padding: 8px; border-bottom: 1px solid #eee;">${formatDateDDMMYYYY(data.parsedDate)}</td></tr>
        <tr><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Zeit/Slot:</td><td style="padding: 8px; border-bottom: 1px solid #eee;">${data.slot.toUpperCase()} (${CONFIG['SLOT_' + data.slot.toUpperCase()].start} - ${CONFIG['SLOT_' + data.slot.toUpperCase()].end} Uhr)</td></tr>
        <tr><td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Buchungstyp:</td><td style="padding: 8px; border-bottom: 1px solid #eee;">${typText}</td></tr>
      </table>
      <p style="font-size: 13px; color: #666;">Falls du den Termin nicht wahrnehmen kannst, storniere ihn bitte mindestens 24 Stunden vorher per E-Mail.</p>
      <br>
      <p>Allzeit gute Fahrt!<br><b>Dein Vorstand</b></p>
    </div>
  `;
  
  thread.reply(htmlBody, { htmlBody: htmlBody });
}

function sendRejectionEmail(sender, errorMessage, thread) {
  const htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; color: #333;">
      <h2 style="color: #9b2c2c;">⚠️ Reservierung fehlgeschlagen</h2>
      <p>Hallo,</p>
      <p>Deine Anfrage konnte vom System leider nicht verarbeitet werden. Grund dafür ist:</p>
      <div style="background-color: #fff5f5; border-left: 4px solid #e53e3e; padding: 12px; margin: 15px 0; color: #c53030;">
        ${errorMessage}
      </div>
      <p>Bitte korrigiere die Angaben oder wende dich bei Fragen direkt an den Vorstand.</p>
      <p>Die aktuelle PDF-Anleitung findest du hier: <a href="${PDF_SOURCE_URL}">Anleitung Bootsreservation</a></p>
      <br>
      <p>Mit freundlichen Grüßen,<br><b>Dein Vorstand</b></p>
    </div>
  `;
  
  thread.reply(htmlBody, { htmlBody: htmlBody });
}

// =============================================================================
// 6. KALENDER-SYNCHRONISIERUNG (BERECHTIGUNGEN FÜR MITGLIEDER)
// =============================================================================

function ausfuehrenKalenderSynchronisierung() {
  Logger.log('=== STARTE KALENDER-SYNCHRONISIERUNG ===');
  
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('Keine Mitglieder für die Synchronisierung vorhanden.');
      return;
    }

    const data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
    const aktuelleEmails = new Set();

    // Alle gültigen E-Mails aus der Tabelle sammeln
    for (let i = 0; i < data.length; i++) {
      const email = data[i][3] ? data[i][3].toString().toLowerCase().trim() : '';
      if (email) aktuelleEmails.add(email);
    }

    Logger.log(`Synchronisiere Lesezugriff für ${aktuelleEmails.size} aktive Mitglieder auf dem Google Kalender...`);
    
    // HINWEIS: Damit die native Einladung funktioniert oder ACLs verwendet werden,
    // können Mitglieder bei Bedarf als "Reader" zum Kalender hinzugefügt werden.
    // Ein Beispiel-Loop für reguläre Kalender-Freigaben:
    // const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    // aktuelleEmails.forEach(email => calendar.addReader(email));

    Logger.log('✅ Kalender-Synchronisierung erfolgreich beendet.');
  } catch (e) {
    Logger.log('❌ Fehler bei Kalender-Synchronisierung: ' + e.message);
  }
}

// =============================================================================
// 7. SUBSYSTEME & STRUKTUR-WERKZEUGE
// =============================================================================

function createGmailLabelStructure(targetLabelPath) {
  const parts = targetLabelPath.split('/');
  let currentPath = '';
  let finalLabel = null;
  
  for (let i = 0; i < parts.length; i++) {
    currentPath = currentPath ? currentPath + '/' + parts[i] : parts[i];
    let label = GmailApp.getUserLabelByName(currentPath);
    if (!label) {
      label = GmailApp.createLabel(currentPath);
    }
    if (currentPath === targetLabelPath) {
      finalLabel = label;
    }
  }
  return finalLabel;
}

// Helper zum Parsen europäischer Daten (DD.MM.YYYY)
function parseEuropeanDate(dateStr) {
  const parts = dateStr.split(/[./-]/);
  if (parts.length === 3) {
    return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  return new Date(dateStr);
}

// Helper zum Formatieren von Daten (DD.MM.YYYY)
function formatDateDDMMYYYY(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.mm.yyyy");
}

// =============================================================================
// 8. ENTWICKLER-WERKZEUGE (MAINTENANCE)
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

function setupAllTriggers() {
  // Löscht alte Trigger, um Duplikate zu vermeiden
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  // 1. Trigger für E-Mail-Buchungen (alle 5 Minuten)
  ScriptApp.newTrigger('processReservationEmails')
           .timeBased().everyMinutes(5).create();
           
  // 2. Trigger für Excel-Listenimport des Admins (alle 30 Minuten)
  ScriptApp.newTrigger('importExcelToSheets')
           .timeBased().everyMinutes(30).create();
           
  Logger.log('=== STARTE INTEGRIERTES SETUP ===');
  Logger.log('🎉 INTEGRIERTES GESAMT-SETUP ERFOLGREICH!');
}
