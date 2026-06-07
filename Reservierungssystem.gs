// =============================================================================
// BC1890 - Integriertes Gesamtskript: Reservierung, Onboarding & Tracking
// Google Apps Script für Google Calendar, Gmail, Drive & Google Spreadsheet
// =============================================================================

// Globale URL-Quelle für die PDF-Anleitung
const PDF_SOURCE_URL = 'https://github.com/marlan99/Bootsverein-BJB/blob/main/Anleitung%20Bootsreservation.pdf';

const CONFIG = {
  // <--- KALENDER & ADMIN EINSTELLUNGEN --->
  CALENDAR_ID: 'Bootsclub1890@gmail.com', // Hier die KALENDER ID eintragen
  ADMIN_EMAIL: 'Bootsclub1890@gmail.com',
  GMAIL_LABEL: 'Reservierung/Neu',               
  SLOT_VORMITTAG: { start: '08:00', end: '14:00' },
  SLOT_NACHMITTAG: { start: '14:00', end: '20:00' },

  // <--- ONBOARDING & EXCEL-IMPORT EINSTELLUNGEN --->
  // true  = Willkommens-Mails werden abgefangen und NUR an den Vorstand gesendet
  // false = Mails gehen direkt an die neuen Mitglieder und der Vorstand im CC
  TEST_MODUS_AKTIV: true, 
  // Einstellungen für den automatischen Excel-Listenimport
  EXCEL_SUBJECT: 'Mitgliederliste',
  EXCEL_TARGET_LABEL: 'Reservierung/Mitgliederliste',

  // <--- MITGLIEDER-TRACKING EINSTELLUNGEN --->
  // true  = Der Änderungsbericht wird gesendet, aber kein neuer Schnappschuss gespeichert.
  // false = Der Bericht wird gesendet und der Schnappschuss aktualisiert (Normalbetrieb).
  TRACKING_TEST_MODUS_AKTIV: false
};

// =============================================================================
// 1. KERN-LOGIK: RESERVIERUNGEN & STORNIERUNGEN VERARBEITEN
// =============================================================================

function processReservationEmails() {
  const label = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
  if (!label) {
    Logger.log('Label nicht gefunden: ' + CONFIG.GMAIL_LABEL);
    return;
  }

  const threads = label.getThreads();
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread()) {
        processSingleEmail(message, thread);
      }
    });
  });
}

function processSingleEmail(message, thread) {
  const sender = message.getFrom().match(/[\w.-]+@[\w.-]+/)?.[0] || 'unbekannt';
  const subject = message.getSubject().toLowerCase();
  const body = message.getPlainBody();

  const data = parseEmailTemplate(body);
  const labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
  
  if (!data.valid) {
    sendRejectionEmail(sender, data.error, thread);
    message.markRead();
    
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    return;
  }

  const userId = sender;

  // LOGIK FÜR STORNIERUNG
  if (subject.includes('stornierung') || subject.includes('absage')) {
    executeCancellation(data, userId, thread, message);
    return; 
  }

  // LOGIK FÜR RESERVIERUNG
  const validation = validateRequest(data, userId, sender);
  
  if (!validation.valid) {
    sendRejectionEmail(sender, validation.error, thread);
    message.markRead();
    
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    return;
  }

  const event = createCalendarEvent(data, userId);
  if (event) {
    sendConfirmationEmail(sender, event, data, thread);
    message.markRead();
    
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    thread.addLabel(labelErledigt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive(); 
  } else {
    sendRejectionEmail(sender, 'Fehler beim Erstellen des Termins.', thread);
    message.markRead();
    
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
  }
}

function parseEmailTemplate(body) {
  const lines = body.split('\n').map(l => l.trim());
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
      if (line.startsWith(key + ':')) {
        data[prop] = line.substring(key.length + 1).trim();
      }
    }
  });

  if (!data.date || !data.slot) {
    data.error = 'Fehlende Pflichtfelder: Datum, Slot';
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
    data.error = 'Slot muss "Vormittag" oder "Nachmittag" sein.';
    return data;
  }

  if (!data.type) {
    data.type = 'standard';
  } else {
    data.type = data.type.toLowerCase();
  }

  if (!['standard', 'joker'].includes(data.type)) {
    data.error = 'Typ kann nur "Standard" oder "Joker" sein.';
    return data;
  }

  data.valid = true;
  return data;
}

function parseEuropeanDate(dateStr) {
  if (!dateStr) return null;
  dateStr = dateStr.trim();
  
  const textMonthRegex = /^(\d{1,2})\.\s*([a-zA-ZäÄöÖüÜß]+)\s*(\d{4})$/;
  const textMatch = dateStr.match(textMonthRegex);
  if (textMatch) {
    const day = parseInt(textMatch[1], 10);
    const monthName = textMatch[2].toLowerCase();
    const year = parseInt(textMatch[3], 10);
    
    const months = {
      'januar': 0, 'jan': 0, 'februar': 1, 'feb': 1, 'märz': 2, 'mrz': 2, 'maerz': 2,
      'april': 3, 'apr': 3, 'mai': 4, 'juni': 5, 'jun': 5, 'juli': 6, 'jul': 6,
      'august': 7, 'aug': 7, 'september': 8, 'sep': 8, 'oktober': 9, 'okt': 9,
      'november': 10, 'nov': 10, 'dezember': 11, 'dez': 11
    };
    
    if (months[monthName] !== undefined) {
      return new Date(year, months[monthName], day, 12, 0, 0);
    }
  }
  
  const numericRegex = /^(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})$/;
  const numMatch = dateStr.match(numericRegex);
  if (numMatch) {
    const day = parseInt(numMatch[1], 10);
    const month = parseInt(numMatch[2], 10) - 1;
    const year = parseInt(numMatch[3], 10);
    
    const parsedDate = new Date(year, month, day, 12, 0, 0);
    
    if (parsedDate.getFullYear() === year && parsedDate.getMonth() === month && parsedDate.getDate() === day) {
      return parsedDate;
    }
  }
  
  const fallbackDate = new Date(dateStr);
  if (!isNaN(fallbackDate.getTime())) {
    return fallbackDate;
  }
  
  return null;
}

function formatDateDDMMYYYY(date) {
  if (!date || isNaN(date.getTime())) return 'Ungültiges Datum';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function validateRequest(data, userId, sender) {
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
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);

  if (!calendar) {
    return { valid: false, error: 'Konfigurationsfehler: Kalender wurde nicht gefunden.' };
  }

  if (data.parsedDate < today) {
    return { valid: false, error: 'Datum liegt in der Vergangenheit.' };
  }

  const seasonStart = getCurrentSeasonStart();

  if (data.type === 'joker') {
    if (data.parsedDate.getFullYear() !== today.getFullYear()) {
      return { 
        valid: false, 
        error: `Joker-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.` 
      };
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
      return { valid: false, error: 'Du hast bereits 2 Joker-Termine in dieser Saison.' };
    }
  }

  if (data.type === 'standard') {
    if (data.parsedDate.getFullYear() !== today.getFullYear()) {
      return {
        valid: false,
        error: `Standard-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.`
      };
    }

    const seasonEnd = new Date(today.getFullYear(), 11, 31, 23, 59, 59);
    const existingEvents = calendar.getEvents(seasonStart, seasonEnd);

    const activeStandardEvents = existingEvents.filter(e => {
      const desc = e.getDescription() || '';
      const title = e.getTitle() || '';
      const eventStart = e.getStartTime();
      
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && 
             !title.includes('JOKER') && 
             eventStart >= today;
    });

    if (activeStandardEvents.length > 0) {
      const bestehenderTermin = activeStandardEvents[0];
      return {
        valid: false,
        error: `Du hast bereits einen Standard-Termin in dieser Saison gebucht (am ${formatDateDDMMYYYY(bestehenderTermin.getStartTime())}). Erst wenn dieser Termin vorbei ist, kannst du einen neuen Standard-Termin vereinbaren.`
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
    return { valid: false, error: 'Dieser Slot ist bereits belegt.' };
  }

  data.startTime = startTime;
  data.endTime = endTime;

  return { valid: true };
}

function createCalendarEvent(data, userId) {
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    let title = data.type === 'joker' ? `JOKER – ${data.name}` : data.name;

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
    Logger.log('Fehler beim Erstellen: ' + e);
    return null;
  }
}

// =============================================================================
// 2. EXCEL-IMPORT SYSTEM (EXCEL -> GOOGLE SHEET)
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
  
  if (threads.length === 0) {
    Logger.log("Keine neuen passenden E-Mails für den Excel-Import gefunden.");
    return;
  }
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    let importErfolgreich = false;
    
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      const sender = message.getFrom().toLowerCase();
      const subject = message.getSubject();
      
      if (subject !== CONFIG.EXCEL_SUBJECT) {
        Logger.log(`Übersprungen: Betreff "${subject}" stimmt nicht exakt mit "${CONFIG.EXCEL_SUBJECT}" überein.`);
        continue;
      }
      
      if (sender.indexOf(adminEmail.toLowerCase()) === -1) {
        Logger.log(`WARNUNG: E-Mail von unbefugtem Absender blockiert: ${sender}`);
        continue; 
      }
      
      const attachments = message.getAttachments();
      
      for (let k = 0; k < attachments.length; k++) {
        const attachment = attachments[k];
        const isExcel = attachment.getContentType() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                        attachment.getName().toLowerCase().endsWith(".xlsx");
        
        if (isExcel) {
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
              Logger.log("⚠️ Warnung: Die importierte Excel-Datei enthält keine Daten ausser der Kopfzeile.");
              Drive.Files.remove(tempSheetFile.id);
              continue;
            }
            
            const newValues = tempSheet.getRange(2, 1, tempLastRow - 1, tempLastColumn).getValues();
            
            const targetSpreadsheet = SpreadsheetApp.openById(sheetId);
            const targetSheet = targetSpreadsheet.getSheets()[0];
            const targetLastRow = targetSheet.getLastRow();
            
            if (targetLastRow > 2) {
              targetSheet.getRange(3, 1, targetLastRow - 2, targetSheet.getLastColumn()).clearContent();
            }
            
            targetSheet.getRange(3, 1, newValues.length, tempLastColumn).setValues(newValues);
            Logger.log(`✅ Mitgliederliste erfolgreich aktualisiert. Daten wurden ab Zeile 3 ersetzt. (ID: ${sheetId})`);
            
            Drive.Files.remove(tempSheetFile.id);
            tempSheetFile = null; 
            importErfolgreich = true;
            
            // AUSFÜHRUNG DER TRACKING-FUNKTION (Kettenreaktion: Import -> Tracking -> Onboarding)
            if (typeof tracklistchanges === 'function') {
              Logger.log("Starte tracklistchanges()...");
              tracklistchanges(); 
            } else {
              Logger.log("Hinweis: Die Funktion tracklistchanges wurde nicht gefunden.");
            }
            
            break; 
            
          } catch (e) {
            Logger.log(`❌ Fehler beim Verarbeiten der Import-Datei: ${e.toString()}`);
            if (tempSheetFile && tempSheetFile.id) {
              try { Drive.Files.remove(tempSheetFile.id); } catch(err) {}
            }
          }
        }
      }
    }
    
    if (importErfolgreich) {
      threads[i].markRead();
      let label = GmailApp.getUserLabelByName(CONFIG.EXCEL_TARGET_LABEL) || createGmailLabelStructure(CONFIG.EXCEL_TARGET_LABEL);
      
      if (label) {
        threads[i].addLabel(label);
        Logger.log(`✉️ E-Mail-Thread wurde als gelesen markiert und nach "${CONFIG.EXCEL_TARGET_LABEL}" verschoben.`);
      }
    }
  }
}

// =============================================================================
// 3. MITGLIEDERLISTEN-TRACKING-SYSTEM (DATENÄNDERUNGEN ERKENNEN)
// =============================================================================

function tracklistchanges() {
  Logger.log('=== STARTE MITGLIEDERLISTEN-TRACKING ===');
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  const adminEmail = scriptProperties.getProperty('ADMIN_EMAIL') || CONFIG.ADMIN_EMAIL;

  if (!sheetId || !adminEmail) {
    Logger.log('❌ FEHLER: Weder SHEET_CONFIG_ID noch ADMIN_EMAIL konnten gefunden werden.');
    return;
  }

  const lastSnapshotRaw = scriptProperties.getProperty('MEMBER_LIST_SNAPSHOT');
  let currentSnapshot = {};

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
    
    // Nach Initialisierung ebenfalls Onboarding aufrufen
    if (typeof checkAndWelcomeNewMembers === 'function') {
      checkAndWelcomeNewMembers();
    }
    return;
  }

  const lastSnapshot = JSON.parse(lastSnapshotRaw);
  let addedMembers = [];
  let removedMembers = [];
  let updatedMembers = [];

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

  for (let id in lastSnapshot) {
    if (!currentSnapshot[id]) {
      const removed = lastSnapshot[id];
      removed.id = id;
      removedMembers.push(removed);
    }
  }

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

  // LOGISCHER ANSCHLUSS: Onboarding starten, um neue Datensätze direkt zu begrüßen
  if (typeof checkAndWelcomeNewMembers === 'function') {
    Logger.log("🚀 Starte automatische Prüfung auf neue Mitglieder (checkAndWelcomeNewMembers)...");
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

  let htmlBody = `
    <div style="font-family: sans-serif; color: #333; max-width: 750px; line-height: 1.5;">
      <h2 style="color: #1a365d; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 20px;">Bericht: Änderungen an der Mitgliederliste</h2>
      <p>Hallo Vorstand,<br>das automatisierte System hat Änderungen in der Mitglieder-Tabelle festgestellt. Nachfolgend findest du alle Details:</p>
  `;

  if (added.length > 0) {
    htmlBody += `<h3 style="color: #2f855a; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #c6f6d5; padding-bottom: 4px;">➕ Neu hinzugefügte Mitglieder (${added.length})</h3>
    <table style="${tableStyle}"><tr><th style="${thStyle} width: 10%;">ID</th><th style="${thStyle} width: 25%;">Name</th><th style="${thStyle} width: 40%;">E-Mail</th><th style="${thStyle} width: 25%;">Mobile</th></tr>`;
    added.forEach(m => { htmlBody += `<tr><td style="${tdStyle}"><code>${m.id || ''}</code></td><td style="${tdStyle}"><b>${m.vorname} ${m.nachname}</b></td><td style="${tdStyle}">${m.email}</td><td style="${tdStyle}">${m.mobile || '-'}</td></tr>`; });
    htmlBody += `</table>`;
  }

  if (removed.length > 0) {
    htmlBody += `<h3 style="color: #9b2c2c; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #fed7d7; padding-bottom: 4px;">➖ Entfernte Mitglieder (${removed.length})</h3>
    <table style="${tableStyle}"><tr><th style="${thStyle} width: 10%;">ID</th><th style="${thStyle} width: 25%;">Name</th><th style="${thStyle} width: 40%;">E-Mail</th><th style="${thStyle} width: 25%;">Mobile</th></tr>`;
    removed.forEach(m => { htmlBody += `<tr style="background-color: #fafafa;"><td style="${tdStyle} color: #94a3b8;"><code>${m.id || ''}</code></td><td style="${tdStyle} color: #94a3b8;">${m.vorname} ${m.nachname}</td><td style="${tdStyle} color: #94a3b8;">${m.email}</td><td style="${tdStyle} color: #94a3b8;">${m.mobile || '-'}</td></tr>`; });
    htmlBody += `</table>`;
  }

  if (updated.length > 0) {
    htmlBody += `<h3 style="color: #dd6b20; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #feebc8; padding-bottom: 4px;">⚠️ Aktualisierte Mitgliedsdaten (${updated.length})</h3>`;
    updated.forEach(m => {
      const vNameStyle = m.changedFields.includes('vorname') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
      const nNameStyle = m.changedFields.includes('nachname') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
      const emailStyle = m.changedFields.includes('email') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';
      const mobilStyle = m.changedFields.includes('mobile') ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : '';

      htmlBody += `
        <div style="margin-bottom: 25px; border-left: 4px solid #dd6b20; padding-left: 12px;">
          <span style="font-size: 15px; font-weight: bold; color: #2d3748;">Mitglied: ${m.current.vorname} ${m.current.nachname}</span> <span style="font-size: 13px; color: #718096; margin-left: 10px;">(ID: <code>${m.id}</code>)</span>
          <table style="${tableStyle} margin-top: 6px; margin-bottom: 5px;">
            <tr style="background-color: #f8fafc;"><th style="${thStyle} width: 25%;">Feld</th><th style="${thStyle} width: 37.5%;">Alter Wert</th><th style="${thStyle} width: 37.5%;">Neuer Wert</th></tr>
            <tr><td style="${tdStyle} ${vNameStyle}">Vorname</td><td style="${tdStyle} ${vNameStyle}">${m.old.vorname || '-'}</td><td style="${tdStyle} ${vNameStyle}">${m.current.vorname || '-'}</td></tr>
            <tr><td style="${tdStyle} ${nNameStyle}">Nachname</td><td style="${tdStyle} ${nNameStyle}">${m.old.nachname || '-'}</td><td style="${tdStyle} ${nNameStyle}">${m.current.nachname || '-'}</td></tr>
            <tr><td style="${tdStyle} ${emailStyle}">E-Mail</td><td style="${tdStyle} ${emailStyle}">${m.old.email || '-'}</td><td style="${tdStyle} ${emailStyle}">${m.current.email || '-'}</td></tr>
            <tr><td style="${tdStyle} ${mobilStyle}">Mobile</td><td style="${tdStyle} ${mobilStyle}">${m.old.mobile || '-'}</td><td style="${tdStyle} ${mobilStyle}">${m.current.mobile || '-'}</td></tr>
          </table>
        </div>`;
    });
  }

  htmlBody += `<hr style="border: 0; border-top: 1px solid #e2e8f0; margin-top: 40px;"><p style="font-size: 12px; color: #a0aec0;">Generiert am: ${new Date().toLocaleString('de-DE')}</p></div>`;

  let plainBody = `Änderungsbericht Mitgliederliste BC1890\n\n`;
  if (added.length > 0) plainBody += `Neu (${added.length}):\n` + added.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}`).join('\n') + `\n\n`;
  if (removed.length > 0) plainBody += `Entfernt (${removed.length}):\n` + removed.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}`).join('\n') + `\n\n`;
  if (updated.length > 0) plainBody += `Geändert (${updated.length}):\n` + updated.map(m => `- ID: ${m.id}, Änderungen: ${m.textDetails.join(', ')}`).join('\n') + `\n`;

  try {
    GmailApp.sendEmail(adminEmail, subject, plainBody, { htmlBody: htmlBody });
  } catch (err) {
    Logger.log('❌ Fehler beim Senden des Änderungsberichts: ' + err.message);
  }
}

// =============================================================================
// 4. ONBOARDING & WILLKOMMENS-SYSTEM
// =============================================================================

function checkAndWelcomeNewMembers() {
  const modusText = CONFIG.TEST_MODUS_AKTIV ? '⚠️ TESTMODUS (AKTIV)' : '🚀 LIVE-BETRIEB';
  Logger.log(`=== STARTE PRÜFUNG AUF NEUE MITGLIEDER [Modus: ${modusText}] ===`);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  const adminEmail = CONFIG.ADMIN_EMAIL;

  if (!sheetId || !adminEmail) {
    Logger.log('❌ KRITISCHER FEHLER: Tabellen-ID oder Admin-E-Mail konnte nicht ermittelt werden.');
    return;
  }

  let welcomedMembersRaw = scriptProperties.getProperty('WELCOMED_MEMBER_IDS');
  let welcomedMemberIds = welcomedMembersRaw ? JSON.parse(welcomedMembersRaw) : [];
  
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
    let currentTableIds = [];
    
    for (let i = 0; i < dataRange.length; i++) {
      const memberId = dataRange[i][0] ? dataRange[i][0].toString().trim() : '';
      if (memberId) currentTableIds.push(memberId);
    }

    let cleanedWelcomedIds = welcomedMemberIds.filter(id => currentTableIds.includes(id));
    let removedCount = welcomedMemberIds.length - cleanedWelcomedIds.length;
    if (removedCount > 0) {
      Logger.log(`🧹 BEREINIGUNG: ${removedCount} gelöschte(s) Mitglied(er) aus dem Skript-Gedächtnis entfernt.`);
    }

    let mailsSentCount = 0;

    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      const memberId = row[0] ? row[0].toString().trim() : '';
      const vorname = row[1] ? row[1].toString().trim() : '';
      const nachname = row[2] ? row[2].toString().trim() : '';
      const email = row[3] ? row[3].toString().trim() : '';
      
      if (!memberId || !email) continue;

      if (!cleanedWelcomedIds.includes(memberId)) {
        if (!isInitialRun) {
          sendWelcomeMail(email, vorname, nachname, adminEmail);
          mailsSentCount++;
        }
        cleanedWelcomedIds.push(memberId);
      }
    }

    scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify(cleanedWelcomedIds));
    Logger.log(`Prüfung abgeschlossen. ${mailsSentCount} neue(s) Mitglied(er) verarbeitet.`);
    
  } catch (e) {
    Logger.log('Fehler im Onboarding-Script: ' + e.message);
  }
}

function sendWelcomeMail(toEmail, vorname, nachname, adminEmail) {
  const name = vorname ? vorname : 'Mitglied';
  const scriptProperties = PropertiesService.getScriptProperties();
  
  let finalReceiver = toEmail;
  let finalCc = adminEmail; // Im Live-Modus erhält der Vorstand standardmässig ein CC
  let subject = 'Herzlich willkommen beim Bootsclub 1890! ⛵';
  let testNoticeHtml = '';
  let testNoticePlain = '';

  if (CONFIG.TEST_MODUS_AKTIV) {
    finalReceiver = adminEmail; 
    finalCc = ''; // Im Test-Modus kein CC nötig, da Receiver bereits der Admin ist
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
    Herzlich Willkommen im <b>Bootsclub 1890</b>! Deine E-Mail-Adresse wurde erfolgreich für unser automatisiertes Reservierungssystem freigeschaltet.<br><br>
    Ab sofort kannst du Bootstermine direkt per E-Mail reservieren. <br>
    <b>Im Anhang dieser E-Mail findest du die detaillierte Anleitung als PDF-Datei.</b><br><br>
    Hier sind die wichtigsten Kernpunkte im Überblick:<br>
    • Sende Reservierungen an: <b>${adminEmail}</b>. Die E-Mail muss das Wort <b>Reservierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text enthalten.<br><br>
    • Für eine Stornierung sende einfach das Wort <b>Stornierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text (bis max. 24 Stunden vor dem Termin).<br><br>
    Bitte lies dir die angehängte PDF-Anleitung aufmerksam durch, bevor du deine erste Reservierung vornimmst.<br><br>
    Bei Fragen steht dir der Vorstand jederzeit gerne zur Verfügung.<br><br>
    Allzeit gute Fahrt und viel Spass auf dem Wasser!<br><br>
    <b>Dein Vorstand</b><br>
  `;

  const plainBody = `${testNoticePlain}Hallo ${name},\n\nherzlich willkommen beim Bootsclub 1890!\nDeine E-Mail wurde für das Reservierungssystem freigeschaltet.\n\nEine detaillierte Anleitung findest du im Anhang dieser E-Mail als PDF.\n\nBitte sende Reservierungen an ${adminEmail}.\n\nAllzeit gute Fahrt!\nDein Vorstand`;

  try {
    // Holt die ID dynamisch aus den permanenten Skript-Eigenschaften
    const fileId = scriptProperties.getProperty('PDF_FILE_ID');
    if (!fileId) throw new Error('Keine gültige Google Drive File ID konfiguriert.');
    
    const pdfFile = DriveApp.getFileById(fileId);
    const attachmentBlob = pdfFile.getBlob();

    GmailApp.sendEmail(finalReceiver, subject, plainBody, {
      cc: finalCc, // Setzt die E-Mail im Live-Betrieb automatisch in Kopie für den Vorstand
      replyTo: adminEmail,
      htmlBody: htmlBody,
      attachments: [attachmentBlob]
    });
  } catch (error) {
    Logger.log(`❌ FEHLER beim Senden der Willkommens-Mail: ${error.message}`);
  }
}

// =============================================================================
// 5. HILFSFUNKTIONEN & STORNIERUNGSLOGIK
// =============================================================================

function executeCancellation(data, userId, thread, message) {
  const labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
  const memberData = getAuthorizedUserData(userId);
  
  if (!memberData) {
    GmailApp.sendEmail(userId, 'Löschen der Buchung abgelehnt', `Deine E-Mail-Adresse (${userId}) ist nicht im System hinterlegt.`, { replyTo: CONFIG.ADMIN_EMAIL });
    message.markRead();
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    return;
  }
  
  data.name = memberData.name;
  const jetzt = new Date(); 
  const slotTime = data.slot === 'vormittag' ? CONFIG.SLOT_VORMITTAG : CONFIG.SLOT_NACHMITTAG;
  const terminStartZeit = new Date(data.parsedDate);
  const [sh, sm] = slotTime.start.split(':');
  terminStartZeit.setHours(sh, sm, 0, 0); 

  const stornierungsFrist = new Date(terminStartZeit.getTime() - (24 * 60 * 60 * 1000));

  if (jetzt > stornierungsFrist) {
    let fehlerGrund = terminStartZeit < jetzt ? 'Der Termin liegt in der Vergangenheit.' : `Die Frist für eine automatische Stornierung (24h vor Beginn) ist abgelaufen.`;
    GmailApp.sendEmail(userId, 'Löschen der Buchung abgelehnt', `Hallo ${data.name},\n\n❌ Grund: ${fehlerGrund}`, { replyTo: CONFIG.ADMIN_EMAIL });
    message.markRead();
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    return; 
  }

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const terminEndZeit = new Date(terminStartZeit);
  const [eh, em] = slotTime.end.split(':');
  terminEndZeit.setHours(eh, em, 0, 0);

  const events = calendar.getEvents(terminStartZeit, terminEndZeit);
  const userEvent = events.find(e => e.getDescription().includes(`Mitglieder-ID: ${memberData.id}`));

  if (userEvent) {
    if (userEvent.getTitle().toUpperCase().includes('JOKER')) {
      GmailApp.sendEmail(userId, 'Stornierung fehlgeschlagen', `❌ Joker-Termine können nicht automatisch storniert werden.`, { replyTo: CONFIG.ADMIN_EMAIL });
      return;
    }

    userEvent.deleteEvent(); 
    GmailApp.sendEmail(userId, 'Bestätigung: Termin freigegeben', `Deine Reservierung für den ${formatDateDDMMYYYY(data.parsedDate)} wurde storniert.`, { replyTo: CONFIG.ADMIN_EMAIL });
    
    message.markRead();
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    thread.addLabel(labelErledigt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
  }
}

function sendConfirmationEmail(to, event, data, thread) {
  const subject = 'Buchung bestätigt: ' + event.getTitle();
  const htmlBody = `Hallo ${data.name},<br><br>dein Termin wurde erfolgreich eingetragen:<br><br>&#128197; <b>Datum:</b> ${formatDateDDMMYYYY(data.parsedDate)}<br>&#9200; <b>Slot:</b> ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}<br><br>Dein Vorstand`;
  const plainBody = `Hallo ${data.name},\n\ndein Termin wurde erfolgreich eingetragen.`;

  try {
    GmailApp.sendEmail(to, subject, plainBody, { replyTo: CONFIG.ADMIN_EMAIL, htmlBody: htmlBody });
  } catch (error) {
    if (thread) thread.createDraftReply(plainBody, { htmlBody: htmlBody });
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

function getAuthorizedUserData(email) {
  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  let ss, sheet;

  try { if (sheetId) { ss = SpreadsheetApp.openById(sheetId); sheet = ss.getSheets()[0]; } } catch (e) {}

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
      sheet.appendRow(["BJB-001", "Vorstand", "Boot", CONFIG.ADMIN_EMAIL, "Nicht hinterlegt"]);
      sheet.autoResizeColumns(1, 5);

      scriptProperties.setProperty('SHEET_CONFIG_ID', sheetId);
    } catch (err) { return null; }
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null; 
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const searchEmail = email.trim().toLowerCase();
    
    for (let i = 0; i < dataRange.length; i++) {
      if (dataRange[i][3] && dataRange[i][3].toString().trim().toLowerCase() === searchEmail) {
        return {
          id: dataRange[i][0] ? dataRange[i][0].toString().trim() : 'Keine ID', 
          name: `${dataRange[i][1] || ''} ${dataRange[i][2] || ''}`.trim() || email,   
          mobile: dataRange[i][4] ? dataRange[i][4].toString().trim() : 'Nicht hinterlegt'
        };
      }
    }
  } catch (e) { return null; }
  return null;
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
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const tomorrowStart = new Date();
  tomorrowStart.setDate(tomorrowStart.getDate() + 1);
  tomorrowStart.setHours(0, 0, 0, 0);
  
  const tomorrowEnd = new Date(tomorrowStart);
  tomorrowEnd.setHours(23, 59, 59, 999);
  
  const events = calendar.getEvents(tomorrowStart, tomorrowEnd);
  
  events.forEach(event => {
    const emailMatch = (event.getDescription() || "").match(/Kontakt:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch && emailMatch[1]) {
      const slotName = event.getStartTime().getHours() === 8 ? "Vormittag (08:00 - 14:00)" : "Nachmittag (14:00 - 20:00)";
      let body = `Hallo!\n\nAutomatische Erinnerung für deine Reservierung morgen:\n📅 Datum: ${formatDateDDMMYYYY(tomorrowStart)}\n⏱️ Slot: ${slotName}\n\nViel Spass!`;
      GmailApp.sendEmail(emailMatch[1].trim(), `Erinnerung: Deine Boot Buchung für morgen!`, body);
    }
  });
}

/**
 * Hilfsfunktion: Garantiert, dass das Google Sheet vor dem Start aller Trigger 
 * mit der korrekten Zeilenstruktur (Zeile 1 & 2) existiert.
 */
function ensureInitialSheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID;
  
  if (!sheetId) {
    Logger.log("📂 Initialisiere Google Sheet und Ordnerstruktur für den Erststart...");
    // Ruft die bestehende Funktion auf, die die Datei, den Ordner sowie Zeile 1 & 2 generiert
    getAuthorizedUserData(CONFIG.ADMIN_EMAIL);
    Logger.log("✅ Google Sheet wurde erfolgreich im Google Drive angelegt (inkl. Strukturzeilen 1 & 2).");
  } else {
    Logger.log("ℹ️ Google Sheet existiert bereits. ID: " + sheetId);
  }
}

/**
 * Holt die PDF von der URL und speichert sie am selben Ort wie die Mitgliederliste.
 * Falls die Datei bereits existiert, wird sie nur aktualisiert, wenn die Online-Datei neuer ist.
 */
function fetchAndSyncAnleitungPDF() {
  Logger.log("🔄 Synchronisiere PDF-Anleitung von GitHub...");
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  
  if (!sheetId) {
    Logger.log("⚠️ Fehler: SHEET_CONFIG_ID noch nicht vorhanden. Kann Pfad nicht ermitteln.");
    return;
  }
  
  // Ordner der Mitgliederliste ermitteln
  const sheetFile = DriveApp.getFileById(sheetId);
  const parents = sheetFile.getParents();
  let targetFolder = DriveApp.getRootFolder();
  if (parents.hasNext()) {
    targetFolder = parents.next();
  }
  
  // Datei von GitHub abrufen
  // Für GitHub-Links konvertieren wir zu raw, falls ein normaler Blob-Abruf scheitert, 
  // aber UrlFetchApp holt standardmäßig die Antwort als Byte-Stream.
  let targetUrl = PDF_SOURCE_URL;
  if (targetUrl.includes('github.com') && !targetUrl.includes('raw.githubusercontent.com') && !targetUrl.includes('?raw=true')) {
    targetUrl = targetUrl.replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/');
  }
  
  const response = UrlFetchApp.fetch(targetUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    Logger.log("❌ Fehler beim Abrufen der PDF von URL: " + response.getResponseCode());
    return;
  }
  
  const pdfBlob = response.getBlob().setName("Anleitung Bootsreservation.pdf");
  
  // Prüfen, ob die Datei im Ordner existiert
  const fileName = "Anleitung Bootsreservation.pdf";
  const files = targetFolder.getFilesByName(fileName);
  let localFile = null;
  
  if (files.hasNext()) {
    localFile = files.next();
  }
  
  if (localFile) {
    // Falls die Datei existiert, prüfen wir die Header ("Last-Modified"), falls vom Server bereitgestellt.
    // GitHub Raw liefert zuverlässige Zeitstempel. Falls nicht lesbar, erzwingen wir Aktualisierung.
    const headers = response.getHeaders();
    const remoteLastModifiedStr = headers["Last-Modified"] || headers["last-modified"];
    let shouldUpdate = true;
    
    if (remoteLastModifiedStr) {
      const remoteDate = new Date(remoteLastModifiedStr);
      const localDate = new Date(localFile.getLastUpdated());
      
      if (remoteDate <= localDate) {
        shouldUpdate = false;
        Logger.log("ℹ️ Lokale PDF ist aktuell oder neuer als die Online-Version. Keine Aktualisierung nötig.");
      }
    }
    
    if (shouldUpdate) {
      Logger.log("🔄 Lokale PDF veraltet. Aktualisiere Dateiinhalt...");
      // In Google Apps Script aktualisiert man Dateien via Drive-API oder Überschreiben des Blobs
      localFile.setContent(pdfBlob.getBytes());
    }
    
    // ID permanent sichern
    scriptProperties.setProperty('PDF_FILE_ID', localFile.getId());
    Logger.log(`📌 PDF File-ID registriert: ${localFile.getId()}`);
  } else {
    // Datei neu erstellen
    Logger.log("📥 PDF existiert lokal nicht. Erstelle neue Datei im Zielverzeichnis...");
    const newFile = targetFolder.createFile(pdfBlob);
    scriptProperties.setProperty('PDF_FILE_ID', newFile.getId());
    Logger.log(`📌 Neue PDF File-ID registriert: ${newFile.getId()}`);
  }
}

/**
 * ZENTRALE SETUP-FUNKTION
 * Richtet die Tabelle, alle Labels und alle VIER benötigten Zeit-Trigger vollautomatisch ein.
 */
function setupTriggers() {
  Logger.log('========================================================================');
  Logger.log('🚀 STARTE CENTRAL SYSTEM SETUP...');
  Logger.log('========================================================================');

  // SCHRITT 1: Datenbank und Tabellen-Struktur garantieren
  ensureInitialSheet();

  // SCHRITT 1b: PDF-Anleitung synchronisieren und ID in den Objekteigenschaften speichern
  try {
    fetchAndSyncAnleitungPDF();
  } catch(pdfError) {
    Logger.log("⚠️ Warnung beim PDF-Sync: " + pdfError.toString());
  }

  // SCHRITT 2: Bestehende Trigger bereinigen (Verhindert Doppelungen)
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(t => ScriptApp.deleteTrigger(t));

  // SCHRITT 3: Alle 4 zeitgesteuerten Trigger neu anlegen
  // 1. Reservierungsverarbeitung (alle 5 Minuten)
  ScriptApp.newTrigger('processReservationEmails').timeBased().everyMinutes(5).create();

  // 2. Tägliche Erinnerung um 04:00 Uhr morgens
  ScriptApp.newTrigger('sendDailyReservationReminders').timeBased().everyDays(1).atHour(4).create();

  // 3. Tägliches Onboarding um 08:00 Uhr morgens
  ScriptApp.newTrigger('checkAndWelcomeNewMembers').timeBased().everyDays(1).atHour(8).create();

  // 4. Stündlicher Excel-Import der Mitgliederliste
  ScriptApp.newTrigger('importExcelToSheets').timeBased().everyHours(1).create();

  // SCHRITT 4: Erforderliche Gmail-Labels anlegen
  ['Reservierung/Neu', 'Reservierung/Erledigt', 'Reservierung/Abgelehnt', CONFIG.EXCEL_TARGET_LABEL].forEach(label => {
    if (!GmailApp.getUserLabelByName(label)) createGmailLabelStructure(label);
  });
  
  Logger.log('========================================================================');
  Logger.log('🎉 INTEGRIERTES GESAMT-SETUP ERFOLGREICH!');
  Logger.log('1. Google Sheet "Mitgliederliste" steht bereit.');
  Logger.log('2. E-Mail-Verarbeitung (alle 5 min) aktiv.');
  Logger.log('3. Tägliche Erinnerungen (04:00 Uhr) aktiv.');
  Logger.log('4. Onboarding-Prüfung (08:00 Uhr) aktiv.');
  Logger.log('5. Stündlicher Excel-Import inkl. Tracking-System aktiv.');
  Logger.log('6. PDF-Anleitung heruntergeladen & Speicherort abgeglichen.');
  Logger.log('Alle Gmail-Labels wurden verifiziert bzw. erstellt.');
  Logger.log('========================================================================');
}

// =============================================================================
// 7. ENTWICKLER-WERKZEUGE (MAINTENANCE)
// =============================================================================

function resetWelcomeDatabase() {
  PropertiesService.getScriptProperties().deleteProperty('WELCOMED_MEMBER_IDS');
  Logger.log('Onboarding-Datenbank zurückgesetzt.');
}

function resetTrackingSnapshot() {
  PropertiesService.getScriptProperties().deleteProperty('MEMBER_LIST_SNAPSHOT');
  Logger.log('Tracking-Schnappschuss wurde erfolgreich gelöscht.');
}
