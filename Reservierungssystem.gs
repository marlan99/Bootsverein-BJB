// ===============================================
// BC1890 - Kalender Reservierungssystem
// Google Apps Script für Google Calendar + Gmail
// ===============================================

const CONFIG = {
  CALENDAR_ID: 'Bootsclub1890@gmail.com', // <--- Hier die KALENDER ID eintragen
  GMAIL_LABEL: 'Reservierung/Neu',               
  SLOT_VORMITTAG: { start: '08:00', end: '14:00' },
  SLOT_NACHMITTAG: { start: '14:00', end: '20:00' },
  ADMIN_EMAIL: 'Bootsclub1890@gmail.com'
};

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
  
  // Falls die E-Mail-Vorlage fehlerhaft ausgefüllt wurde
  if (!data.valid) {
    sendRejectionEmail(sender, data.error, thread);
    message.markRead();
    
    // Archivieren unter "Reservierung/Abgelehnt" und "Neu"-Label entfernen
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    return;
  }

  const userId = sender; // userId ist die E-Mail-Adresse

  // LOGIK FÜR STORNIERUNG
  if (subject.includes('stornierung') || subject.includes('absage')) {
    executeCancellation(data, userId, thread, message);
    return; 
  }

  // LOGIK FÜR RESERVIERUNG
  const validation = validateRequest(data, userId, sender);
  
  // Falls die Validierung fehlschlägt (z.B. Slot belegt, Frist verletzt)
  if (!validation.valid) {
    sendRejectionEmail(sender, validation.error, thread);
    message.markRead();
    
    // Archivieren unter "Reservierung/Abgelehnt" und "Neu"-Label entfernen
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
    
    // Unter "Erledigt" ablegen und aus "Neu" entfernen
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    thread.addLabel(labelErledigt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive(); 
  } else {
    sendRejectionEmail(sender, 'Fehler beim Erstellen des Termins.', thread);
    message.markRead();
    
    // Systemfehler kommen ebenfalls in den Abgelehnt-Ordner
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

  // Pflichtfelder prüfen
  if (!data.date || !data.slot) {
    data.error = 'Fehlende Pflichtfelder: Datum, Slot';
    return data;
  }

  // NUTZT DIE NEUE EUROPÄISCHE PARSING-LOGIK
  data.parsedDate = parseEuropeanDate(data.date);

  // Prüfen, ob ein gültiges Datum erzeugt werden konnte
  if (!data.parsedDate || isNaN(data.parsedDate.getTime())) {
    data.error = 'Ungültiges Datum. Das Datum konnte nicht erkannt werden (Erlaubt z.B.: 05.06.2026, 5.6.2026, 5/6/2026 oder 5. Juni 2026).';
    return data;
  }

  // Uhrzeit des Objekts auf 00:00 Uhr zurücksetzen für exakte Vergleiche
  data.parsedDate.setHours(0, 0, 0, 0);

  data.slot = data.slot.toLowerCase();
  if (!['vormittag', 'nachmittag'].includes(data.slot)) {
    data.error = 'Slot muss "Vormittag" oder "Nachmittag" sein.';
    return data;
  }

  // Wenn kein Typ angegeben wurde, setze automatisch "standard"
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

/**
 * INTELLIGENTER PARSER FÜR EUROPÄISCHE DATUMSFORMATE
 */
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

  if (data.type === 'joker') {
    if (data.parsedDate.getFullYear() !== today.getFullYear()) {
      return { 
        valid: false, 
        error: `Joker-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.` 
      };
    }

    const seasonStart = getCurrentSeasonStart();
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

    const seasonStart = new Date(today.getFullYear(), 0, 1);
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
    let title = '';
    if (data.type === 'joker') {
      title = `JOKER – ${data.name}`;
    } else {
      title = data.name;
    }
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

    const event = calendar.createEvent(
      title,
      data.startTime,
      data.endTime,
      { description: description }
    );

    if (data.type === 'joker') {
      event.setColor(CalendarApp.EventColor.RED);
    } else {
      event.setColor(CalendarApp.EventColor.BLUE);
    }

    return event;
  } catch (e) {
    Logger.log('Fehler beim Erstellen: ' + e);
    return null;
  }
}

function sendConfirmationEmail(to, event, data, thread) {
  const subject = 'Buchung bestätigt: ' + event.getTitle();
  
  const htmlBody = `
    Hallo ${data.name},<br><br>
    dein Termin wurde erfolgreich eingetragen:<br><br>
    &#128197; <b>Datum:</b> ${formatDateDDMMYYYY(data.parsedDate)}<br>
    &#9200; <b>Slot:</b> ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}<br>
    &#127991; <b>Typ:</b> ${data.type === 'joker' ? 'Joker' : 'Standard'}<br><br>
    Du erhältst 1 Tag vorher eine Erinnerung per E-Mail.<br><br>
    Vielen Dank!<br>
    Dein Vorstand
  `;

  const plainBody = `Hallo ${data.name},\n\ndein Termin wurde erfolgreich eingetragen:\n\nDatum: ${formatDateDDMMYYYY(data.parsedDate)}\nSlot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}\nTyp: ${data.type === 'joker' ? 'Joker' : 'Standard'}\n\nDu erhältst 1 Tag vorher eine Erinnerung per E-Mail.\n\nVielen Dank!\nDein Vorstand`;

  try {
    GmailApp.sendEmail(to, subject, plainBody, { 
      replyTo: CONFIG.ADMIN_EMAIL,
      htmlBody: htmlBody 
    });
  } catch (error) {
    Logger.log(`⚠️ WARNUNG (E-Mail-Limit): Bestätigung für ${data.name} konnte nicht gesendet werden.`);
    if (thread) {
      thread.createDraftReply(plainBody, {
        htmlBody: `<b>[SYSTEM-NOTIZ: Mail-Limit erreicht - Entwurf generiert]</b><br><br>${htmlBody}`
      });
    }
  }
}

function sendRejectionEmail(to, reason, thread) {
  const subject = 'Buchung abgelehnt';
  const body = `Hallo,\n\nleider konnte deine Reservierung nicht angenommen werden:\n\n❌ Grund: ${reason}\n\nBitte prüfe die Regeln und sende eine korrigierte Anfrage.`;

  try {
    GmailApp.sendEmail(to, subject, body, { replyTo: CONFIG.ADMIN_EMAIL });
  } catch (error) {
    Logger.log(`⚠️ WARNUNG (E-Mail-Limit): Ablehnung an ${to} konnte nicht gesendet werden.`);
    if (thread) {
      thread.createDraftReply(`[SYSTEM-NOTIZ: Mail-Limit erreicht] Ablehnungsgrund: ${reason}`);
    }
  }
}

function getCurrentSeasonStart() {
  const today = new Date();
  return new Date(today.getFullYear(), 0, 1);
}

function setupTriggers() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('processReservationEmails')
    .timeBased()
    .everyMinutes(1)
    .create();

  ScriptApp.newTrigger('sendDailyReservationReminders')
    .timeBased()
    .everyDays(1)
    .atHour(4) 
    .create();

  ['Reservierung/Neu', 'Reservierung/Erledigt', 'Reservierung/Abgelehnt'].forEach(label => {
    if (!GmailApp.getUserLabelByName(label)) {
      GmailApp.createLabel(label);
    }
  });
  
  Logger.log('Setup erfolgreich abgeschlossen. Beide Trigger wurden eingerichtet.');
}

function executeCancellation(data, userId, thread, message) {
  const labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);

  const memberData = getAuthorizedUserData(userId);
  if (!memberData) {
    GmailApp.sendEmail(
      userId, 
      'Löschen der Buchung abgelehnt - Keine Berechtigung', 
      `Hallo,\n\ndeine E-Mail-Adresse (${userId}) ist nicht im System hinterlegt. Automatische Stornierungen sind nicht möglich.`, 
      { replyTo: CONFIG.ADMIN_EMAIL }
    );
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
    let fehlerGrund = '';
    if (terminStartZeit < jetzt) {
      fehlerGrund = 'Der Termin liegt in der Vergangenheit.';
    } else {
      fehlerGrund = `Die Frist für eine automatische Stornierung (bis spätestens 24 Stunden vor Terminbeginn) ist abgelaufen. Letzte Möglichkeit zur Stornierung wäre am ${formatDateDDMMYYYY(stornierungsFrist)} um ${stornierungsFrist.toLocaleTimeString('de-CH', {hour: '2-digit', minute:'2-digit'})} Uhr gewesen.`;
    }

    GmailApp.sendEmail(
      userId, 
      'Löschen der Buchung abgelehnt - 24h-Frist unterschritten', 
      `Hallo ${data.name},\n\ndeine Stornierung für den ${formatDateDDMMYYYY(data.parsedDate)} wurde abgelehnt.\n\n❌ Grund: ${fehlerGrund}\n\nBitte wende dich bei sehr kurzfristigen Absagen direkt per E-Mail an den Vorstand unter: ${CONFIG.ADMIN_EMAIL}.`, 
      { replyTo: CONFIG.ADMIN_EMAIL }
    );
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
    const terminTitel = userEvent.getTitle();

    if (terminTitel.toUpperCase().includes('JOKER')) {
      GmailApp.sendEmail(
        userId, 
        'Löschen der Buchung fehlgeschlagen - Joker-Termin', 
        `Hallo ${data.name},\n\nder Termin am ${formatDateDDMMYYYY(data.parsedDate)} ist als JOKER-Termin deklariert.\n\n❌ Joker-Termine können nicht automatisch storniert werden. Bitte wende dich hierfür direkt an den Vorstand unter: ${CONFIG.ADMIN_EMAIL}.`, 
        { replyTo: CONFIG.ADMIN_EMAIL }
      );
      message.markRead();
      
      const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
      thread.addLabel(labelAbgelehnt);
      if (labelNeu) thread.removeLabel(labelNeu);
      thread.moveToArchive();
      return;
    }

    userEvent.deleteEvent(); 
    
    const slotFormatted = data.slot.charAt(0).toUpperCase() + data.slot.slice(1);
    const dateFormatted = formatDateDDMMYYYY(data.parsedDate);
    const userBody = `Hallo ${data.name},\n\ndeine Reservierung für den ${dateFormatted} (${slotFormatted}) wurde erfolgreich storniert. Der Slot ist wieder freigegeben.`;
    const userSubject = `Bestätigung: Termin am ${dateFormatted} freigegeben`;

    try {
      GmailApp.sendEmail(userId, userSubject, userBody, {
        replyTo: CONFIG.ADMIN_EMAIL,
        threadId: thread.getId()
      });
    } catch (userMailError) {
      thread.reply(userBody);
    }

    try {
      const adminSubject = `INFO: Buchung entfernt - ${data.name}`;
      const adminBody = `Hallo Admin,\n\nein Termin wurde soeben automatisch storniert und im Kalender freigegeben:\n\n` +
                        `👤 Mitglied: ${data.name} (ID: ${memberData.id})\n` +
                        `📧 E-Mail: ${userId}\n` +
                        `📅 Datum: ${formatDateDDMMYYYY(data.parsedDate)}\n` +
                        `⏱️ Slot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)} (${slotTime.start} - ${slotTime.end} Uhr)\n\n` +
                        `Das System hat den Termin gelöscht und den Slot wieder freigegeben.`;
      
      GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, adminSubject, adminBody);
    } catch (adminError) {
      Logger.log(`⚠️ Fehler beim Senden der Admin-Info: ${adminError.message}`);
    }
    
    message.markRead();
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    thread.addLabel(labelErledigt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
  } else {
    GmailApp.sendEmail(userId, 'Löschen der Buchung fehlgeschlagen', `Hallo ${data.name},\n\nes konnte keine auf dich ausgestellte Buchung für den ${formatDateDDMMYYYY(data.parsedDate)} im Slot ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)} gefunden werden.\n\nBitte prüfe deine Angaben oder wende dich an den Vorstand.`, { replyTo: CONFIG.ADMIN_EMAIL });
    message.markRead();
    
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
  }
}

/**
 * Holt die Mitgliederdaten. Liest die ID aus den Skripteigenschaften aus.
 * Sucht/Erstellt den Ordner "Google Kalender Reservierungssystem" und legt dort bei Erstausführung das Dokument an.
 */
function getAuthorizedUserData(email) {
  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  
  let ss;
  let sheet;

  try {
    if (sheetId) {
      ss = SpreadsheetApp.openById(sheetId);
      sheet = ss.getSheets()[0];
    }
  } catch (e) {
    Logger.log('Tabelle konnte mit gespeicherter ID nicht geöffnet werden. Überprüfe Existenz...');
  }

  // LOGIK FÜR DIE ERSTAUSFÜHRUNG: Wenn kein Sheet existiert oder es komplett leer ist
  if (!ss || !sheet || sheet.getLastRow() === 0) {
    try {
      const folderName = "Google Kalender Reservierungssystem";
      let targetFolder;
      
      // 1. Suche nach dem Ordner über den Namen
      const folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        targetFolder = folders.next(); // Ordner existiert bereits
        Logger.log(`Ordner "${folderName}" gefunden.`);
      } else {
        targetFolder = DriveApp.createFolder(folderName); // Ordner neu anlegen
        Logger.log(`Ordner "${folderName}" existierte nicht und wurde neu erstellt.`);
      }

      // 2. Neues Google Sheet Dokument erstellen
      ss = SpreadsheetApp.create('Mitgliederliste');
      sheet = ss.getSheets()[0];
      sheetId = ss.getId();
      
      // 3. Datei in den Zielordner verschieben und aus der "Ablage" entfernen
      const file = DriveApp.getFileById(sheetId);
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);

      // 4. Spalten generieren
      const headers = ["Mitglieder ID", "Vorname", "Name", "E-Mail", "Mobile"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");

      // 5. Ersten Eintrag generieren (Vorstand Boot)
      const firstEntry = ["BJB-001", "Vorstand", "Boot", CONFIG.ADMIN_EMAIL, "Nicht hinterlegt"];
      sheet.appendRow(firstEntry);
      sheet.autoResizeColumns(1, 5);

      // 6. ID dauerhaft in den Skripteigenschaften speichern
      scriptProperties.setProperty('SHEET_CONFIG_ID', sheetId);
      
      Logger.log('========================================================================');
      Logger.log('🎉 INITIALISIERUNG ERFOLGREICH!');
      Logger.log(`Die Datei "Mitgliederliste" liegt im Ordner: Drive -> ${folderName}`);
      Logger.log('ID wurde automatisch in den Skripteigenschaften hinterlegt: ' + sheetId);
      Logger.log('========================================================================');
      
    } catch (createError) {
      Logger.log('Kritischer Fehler beim Erstellen der Ordner- oder Tabellenstruktur: ' + createError.message);
      return null;
    }
  }

  // Reguläres Auslesen der Mitgliederdaten
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null; 
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const searchEmail = email.trim().toLowerCase();
    
    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      if (!row[3]) continue; 
      
      const currentEmail = row[3].toString().trim().toLowerCase(); 
      
      if (currentEmail === searchEmail) {
        const vorname = row[1] ? row[1].toString().trim() : '';
        const nachname = row[2] ? row[2].toString().trim() : '';
        
        let vollerName = `${vorname} ${nachname}`.trim();
        if (!vollerName) { vollerName = email; }

        const mobileRaw = row[4] ? row[4].toString().trim() : '';
        const mobile = mobileRaw !== '' ? mobileRaw : 'Nicht hinterlegt';

        return {
          id: row[0] ? row[0].toString().trim() : 'Keine ID', 
          vorname: vorname,   
          nachname: nachname, 
          name: vollerName,   
          email: row[3],      
          mobile: mobile      
        };
      }
    }
    return null; 
  } catch (e) {
    Logger.log('Fehler beim Einlesen der Mitgliederdaten: ' + e.message);
    return null;
  }
}

function sendDailyReservationReminders() {
  Logger.log("=== STARTE TÄGLICHE ERINNERUNGS-PRÜFUNG ===");
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  
  const tomorrowStart = new Date();
  tomorrowStart.setDate(tomorrowStart.getDate() + 1);
  tomorrowStart.setHours(0, 0, 0, 0);
  
  const tomorrowEnd = new Date();
  tomorrowEnd.setDate(tomorrowEnd.getDate() + 1);
  tomorrowEnd.setHours(23, 59, 59, 999);
  
  const events = calendar.getEvents(tomorrowStart, tomorrowEnd);
  
  events.forEach(event => {
    const description = event.getDescription() || "";
    const emailMatch = description.match(/Kontakt:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    
    if (emailMatch && emailMatch[1]) {
      const memberEmail = emailMatch[1].trim();
      const slotName = event.getStartTime().getHours() === 8 ? "Vormittag (08:00 - 14:00)" : "Nachmittag (14:00 - 20:00)";
      
      const subject = `Erinnerung: Deine Boot Buchung für morgen!`;
      let body = `Hallo!\n\nDies ist die automatische Erinnerung für deine anstehende Reservierung:\n\n`;
      body += `\u{1F4C5} Datum: ${formatDateDDMMYYYY(tomorrowStart)}\n`;
      body += `\u{23F0} Slot: ${slotName}\n\n`;
      body += `Viel Spass auf dem Wasser!\n\nDein Vorstand`;
      
      MailApp.sendEmail(memberEmail, subject, body);
    }
  });
  Logger.log("=== ERINNERUNGS-PRÜFUNG BEENDET ===");
}
