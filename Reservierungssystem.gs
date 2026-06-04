// ===============================================
// BC1890 - Kalender Reservierungssystem
// Google Apps Script für Google Calendar + Gmail
// ===============================================

const CONFIG = {
  CALENDAR_ID: 'DEINE KALENDER ID', // <--- Hier die KALENDER ID eintragen
  GMAIL_LABEL: 'Reservierung/Neu',               
  SHEET_CONFIG_ID: 'DEINE GOOGLE SHEET ID', // <--- Hier die ID der Google Tabelle eintragen    
  SLOT_VORMITTAG: { start: '06:00', end: '14:00' },
  SLOT_NACHMITTAG: { start: '14:00', end: '20:00' },
  JOKER_MAX_WEEKS: 6, // Maximal 6 Wochen in der Zukunft
  STANDARD_MAX_DAYS: 14,
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
    execute(data, userId, thread, message);
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

  const dateMatch = data.date.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!dateMatch) {
    data.error = 'Ungültiges Datumsformat. Verwende: YYYY-MM-DD';
    return data;
  }

  const [_, y, m, d] = dateMatch;
  data.parsedDate = new Date(y, m - 1, d);
  if (isNaN(data.parsedDate)) {
    data.error = 'Ungültiges Datum.';
    return data;
  }

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

function validateRequest(data, userId, sender) {
  // === 0. PRÜFUNG & DATEN RECHERCHE: Ist der Absender auf der Whitelist? ===
  const memberData = getAuthorizedUserData(userId);
  
  if (!memberData) {
    return { 
      valid: false, 
      error: `Deine E-Mail-Adresse (${userId}) ist nicht für das Reservierungssystem freigeschaltet. Bitte wende dich an den Vorstand.` 
    };
  }
  
  // Die echten, geprüften Daten aus der Excel/Google-Tabelle an "data" anheften
  data.memberId = memberData.id;
  data.memberMobile = memberData.mobile;
  // Optional: Überschreibt den Namen aus der E-Mail mit dem offiziellen Namen aus der Tabelle
  if (memberData.name) data.name = memberData.name;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);

  if (!calendar) {
    return { valid: false, error: 'Konfigurationsfehler: Kalender wurde nicht gefunden.' };
  }

  // 1. Prüfung: Liegt das Datum in der Vergangenheit?
  if (data.parsedDate < today) {
    return { valid: false, error: 'Datum liegt in der Vergangenheit.' };
  }

  // === LOGIK FÜR JOKER-TERMINE ===
  if (data.type === 'joker') {
    // Nur Maximalfrist prüfen (Zukunftssperre)
    const maxJokerDate = new Date(today);
    maxJokerDate.setDate(today.getDate() + CONFIG.JOKER_MAX_WEEKS * 7);
    if (data.parsedDate > maxJokerDate) {
      return { 
        valid: false, 
        error: `Joker-Termine dürfen maximal ${CONFIG.JOKER_MAX_WEEKS} Wochen in der Zukunft liegen. Der späteste mögliche Tag ist der: ${maxJokerDate.toLocaleDateString('de-CH')}` 
      };
    }

    const seasonStart = getCurrentSeasonStart();
    const seasonEnd = new Date(seasonStart);
    seasonEnd.setFullYear(seasonStart.getFullYear() + 1);

    // Alle Termine der Saison holen
    const allEvents = calendar.getEvents(seasonStart, seasonEnd);

    // Präzise Filterung über die Mitglieder-ID
    const jokerEvents = allEvents.filter(e => {
      const desc = e.getDescription() || '';
      const title = e.getTitle() || '';
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && title.includes('JOKER');
    });

    if (jokerEvents.length >= 2) {
      return { valid: false, error: 'Du hast bereits 2 Joker-Termine in dieser Saison.' };
    }
  }

  // === LOGIK FÜR STANDARD-TERMINE ===
  if (data.type === 'standard') {
    const maxFutureDate = new Date(today);
    maxFutureDate.setDate(today.getDate() + CONFIG.STANDARD_MAX_DAYS);

    // Zukunftssperre (Maximal 2 Wochen im Voraus)
    if (data.parsedDate > maxFutureDate) {
      return {
        valid: false,
        error: `Buchungstermine dürfen maximal 2 Wochen in der Zukunft liegen. Der späteste mögliche Tag ist der: ${maxFutureDate.toLocaleDateString('de-CH')}`
      };
    }

    // Alle Termine im relevanten 2-Wochen-Fenster holen
    const existingEvents = calendar.getEvents(today, maxFutureDate);

    // Schliesst JOKER aus (da diese nicht das Limit von 1 Termin pro 2 Wochen belasten).
    // Präzise Filterung über die Mitglieder-ID
    const standardEvents = existingEvents.filter(e => {
      const desc = e.getDescription() || '';
      const title = e.getTitle() || '';
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && !title.includes('JOKER');
    });

    if (standardEvents.length > 0) {
      const nextAvailable = new Date(standardEvents[0].getEndTime());
      nextAvailable.setDate(nextAvailable.getDate() + 1);
      return {
        valid: false,
        error: `Du hast bereits einen Termin in den nächsten 2 Wochen gebucht. Nächster mögliche Buchungszeitpunkt ist: ${nextAvailable.toLocaleDateString('de-CH')}`
      };
    }
  }

  // === PRÜFUNG AUF SLOT-ÜBERSCHNEIDUNG ===
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
      {
        description: description,
      }
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
    &#128197; <b>Datum:</b> ${data.parsedDate.toLocaleDateString('de-CH')}<br>
    &#9200; <b>Slot:</b> ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}<br>
    &#127991; <b>Typ:</b> ${data.type === 'joker' ? 'Joker' : 'Standard'}<br><br>
    Du erhältst 1 Tag vorher eine Erinnerung per E-Mail.<br><br>
    Vielen Dank!<br>
    Dein Vorstand
  `;

  const plainBody = `Hallo ${data.name},\n\ndein Termin wurde erfolgreich eingetragen:\n\nDatum: ${data.parsedDate.toLocaleDateString('de-CH')}\nSlot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}\nTyp: ${data.type === 'joker' ? 'Joker' : 'Standard'}\n\nDu erhältst 1 Tag vorher eine Erinnerung per E-Mail.\n\nVielen Dank!\nDein Vorstand`;

  try {
    // Versuche die Bestätigungs-Mail normal zu versenden
    GmailApp.sendEmail(to, subject, plainBody, { 
      replyTo: CONFIG.ADMIN_EMAIL,
      htmlBody: htmlBody 
    });
  } catch (error) {
    // Falls das Limit erreicht ist, fangen wir den Fehler ab
    Logger.log(`⚠️ WARNUNG (E-Mail-Limit): Bestätigung für ${data.name} (Datum: ${data.parsedDate.toLocaleDateString('de-CH')}) konnte nicht gesendet werden.`);
    Logger.log(`Details zum Fehler: ${error.message}`);
    
    // Erstellt einen Entwurf direkt im Thread als visuellen Nachweis im Test
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
    // Versuche die E-Mail normal zu versenden
    GmailApp.sendEmail(to, subject, body, { replyTo: CONFIG.ADMIN_EMAIL });
  } catch (error) {
    // Falls das Limit erreicht ist, fangen wir den Fehler hier ab
    Logger.log(`⚠️ WARNUNG (E-Mail-Limit): Ablehnung an ${to} konnte nicht gesendet werden.`);
    Logger.log(`Details zum Fehler: ${error.message}`);
    
    // Optional: Füge dem Gmail-Thread trotzdem eine Notiz hinzu, 
    // damit du im Postfach siehst, dass das Skript antworten wollte.
    if (thread) {
      thread.createDraftReply(`[SYSTEM-NOTIZ: Mail-Limit erreicht] Ablehnungsgrund: ${reason}`);
    }
  }
}

function getCurrentSeasonStart() {
  const today = new Date();
  const year = today.getMonth() < 6 ? today.getFullYear() - 1 : today.getFullYear();
  return new Date(year, 6, 1); 
}

function setupTriggers() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(t => ScriptApp.deleteTrigger(t));

  // 1. Trigger für den E-Mail-Import (Jede Minute)
  ScriptApp.newTrigger('processReservationEmails')
    .timeBased()
    .everyMinutes(1)
    .create();

  // 2. NEU: Trigger für die tägliche Erinnerung (Morgens zwischen 4 und 5 Uhr)
  ScriptApp.newTrigger('sendDailyReservationReminders')
    .timeBased()
    .everyDays(1)
    .atHour(4) // Startet das Zeitfenster um 4:00 Uhr morgens
    .create();

  // Ordner/Labels in Gmail sicherstellen
  ['Reservierung/Neu', 'Reservierung/Erledigt', 'Reservierung/Abgelehnt'].forEach(label => {
    if (!GmailApp.getUserLabelByName(label)) {
      GmailApp.createLabel(label);
    }
  });
  
  Logger.log('Setup erfolgreich abgeschlossen. Beide Trigger wurden eingerichtet.');
}

function executeCancellation(data, userId, thread, message) {
  const labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);

  // === 0. PRÜFUNG: Ist der Absender auf der Whitelist? ===
  const memberData = getAuthorizedUserData(userId);
  if (!memberData) {
    GmailApp.sendEmail(
      userId, 
      'Stornierung abgelehnt - Keine Berechtigung', 
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
  
  // Namen aus der Tabelle für die Stornierungs-Mail setzen
  data.name = memberData.name;

  const jetzt = new Date(); 

  const slotTime = data.slot === 'vormittag' ? CONFIG.SLOT_VORMITTAG : CONFIG.SLOT_NACHMITTAG;
  const terminStartZeit = new Date(data.parsedDate);
  const [sh, sm] = slotTime.start.split(':');
  terminStartZeit.setHours(sh, sm, 0, 0); 

  const stornierungsFrist = new Date(terminStartZeit.getTime() - (24 * 60 * 60 * 1000));

  // PRÜFUNG: Ist die 24h-Frist bereits unterschritten?
  if (jetzt > stornierungsFrist) {
    let fehlerGrund = '';
    if (terminStartZeit < jetzt) {
      fehlerGrund = 'Der Termin liegt in der Vergangenheit.';
    } else {
      fehlerGrund = `Die Frist für eine automatische Stornierung (bis spätestens 24 Stunden vor Terminbeginn) ist abgelaufen. Letzte Möglichkeit zur Stornierung wäre am ${stornierungsFrist.toLocaleDateString('de-CH')} um ${stornierungsFrist.toLocaleTimeString('de-CH', {hour: '2-digit', minute:'2-digit'})} Uhr gewesen.`;
    }

    GmailApp.sendEmail(
      userId, 
      'Stornierung abgelehnt - 24h-Frist unterschritten', 
      `Hallo ${data.name},\n\ndeine Stornierung für den ${data.parsedDate.toLocaleDateString('de-CH')} wurde abgelehnt.\n\n❌ Grund: ${fehlerGrund}\n\nBitte wende dich bei sehr kurzfristigen Absagen direkt per E-Mail an den Vorstand unter: ${CONFIG.ADMIN_EMAIL}.`, 
      { replyTo: CONFIG.ADMIN_EMAIL }
    );
    message.markRead();
    
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    return; 
  }

  // LÖSCHUNG PROCESS
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const terminEndZeit = new Date(terminStartZeit);
  const [eh, em] = slotTime.end.split(':');
  terminEndZeit.setHours(eh, em, 0, 0);

  const events = calendar.getEvents(terminStartZeit, terminEndZeit);
  // Sucht das Event anhand der Mitglieder-ID aus der Whitelist
  const userEvent = events.find(e => e.getDescription().includes(`Mitglieder-ID: ${memberData.id}`));

  if (userEvent) {
    const terminTitel = userEvent.getTitle();

    // JOKER-SCHUTZ-PRÜFUNG
    if (terminTitel.toUpperCase().includes('JOKER')) {
      GmailApp.sendEmail(
        userId, 
        'Stornierung fehlgeschlagen - Joker-Termin', 
        `Hallo ${data.name},\n\nder Termin am ${data.parsedDate.toLocaleDateString('de-CH')} ist als JOKER-Termin deklariert.\n\n❌ Joker-Termine können nicht automatisch storniert werden. Bitte wende dich hierfür direkt an den Vorstand unter: ${CONFIG.ADMIN_EMAIL}.`, 
        { replyTo: CONFIG.ADMIN_EMAIL }
      );
      message.markRead();
      
      const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
      thread.addLabel(labelAbgelehnt);
      if (labelNeu) thread.removeLabel(labelNeu);
      thread.moveToArchive();
      Logger.log(`Automatische Stornierung von Joker-Termin blockiert für: ${userId}`);
      return;
    }

    // Event im Kalender löschen
    userEvent.deleteEvent(); 
    
    // 1. Antwort an das Mitglied senden
    thread.reply(`Hallo ${data.name},\n\ndeine Reservierung für den ${data.parsedDate.toLocaleDateString('de-CH')} (${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}) wurde erfolgreich storniert. Der Slot ist wieder freigegeben.`);
    
    // 2. Benachrichtigung an den Admin senden
    try {
      const adminSubject = `INFO: Buchung entfernt - ${data.name}`;
      const adminBody = `Hallo Admin,\n\nein Termin wurde soeben automatisch storniert und im Kalender freigegeben:\n\n` +
                        `👤 Mitglied: ${data.name} (ID: ${memberData.id})\n` +
                        `📧 E-Mail: ${userId}\n` +
                        `📅 Datum: ${data.parsedDate.toLocaleDateString('de-CH')}\n` +
                        `⏱️ Slot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)} (${slotTime.start} - ${slotTime.end} Uhr)\n\n` +
                        `Das System hat den Termin gelöscht und den Slot wieder freigegeben.`;
      
      GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, adminSubject, adminBody);
      Logger.log(`Admin-Benachrichtigung für Stornierung gesendet an: ${CONFIG.ADMIN_EMAIL}`);
    } catch (adminError) {
      Logger.log(`⚠️ Fehler beim Senden der Admin-Info: ${adminError.message}`);
    }
    
    message.markRead();
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    thread.addLabel(labelErledigt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
    Logger.log(`Termin erfolgreich storniert: ${terminTitel} für ${userId}`);
  } else {
    GmailApp.sendEmail(userId, 'Stornierung fehlgeschlagen', `Hallo ${data.name},\n\nes konnte keine auf dich ausgestellte Buchung für den ${data.parsedDate.toLocaleDateString('de-CH')} im Slot ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)} gefunden werden.\n\nBitte prüfe deine Angaben oder wende dich an den Vorstand.`, { replyTo: CONFIG.ADMIN_EMAIL });
    message.markRead();
    
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    thread.addLabel(labelAbgelehnt);
    if (labelNeu) thread.removeLabel(labelNeu);
    thread.moveToArchive();
  }
}

/**
 * Prüft die Whitelist und gibt bei Erfolg alle Benutzerdaten zurück.
 * Unterstützt auch fehlende Mobilnummern oder Nachnamen.
 * @param {string} email - Die zu prüfende E-Mail-Adresse
 * @return {Object|null} - Objekt mit Benutzerdaten oder null, wenn nicht gefunden
 */
function getAuthorizedUserData(email) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_CONFIG_ID);
    
    // ÄNDERUNG: Holt automatisch das ERSTE Tabellenblatt (Index 0)
    const sheet = ss.getSheets()[0]; 
    
    if (!sheet) {
      Logger.log(`Fehler: Kein Tabellenblatt in der Datei gefunden.`);
      return null;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null; // Tabelle ist leer (oder enthält nur die Kopfzeile)
    
    // Holt alle Daten ab Zeile 2 (ohne Kopfzeile) bis zur 5. Spalte (Spalte E / Mobile)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const searchEmail = email.trim().toLowerCase();
    
    // Zeilen durchlaufen und nach der E-Mail in Spalte D (Index 3) suchen
    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      
      // Sicherheitsprüfung, falls eine Zeile komplett leer ist
      if (!row[3]) continue; 
      
      const currentEmail = row[3].toString().trim().toLowerCase(); // Spalte D: E-Mail
      
      if (currentEmail === searchEmail) {
        // Person gefunden! 
        const vorname = row[1] ? row[1].toString().trim() : '';
        const nachname = row[2] ? row[2].toString().trim() : '';
        
        // Verhindert doppelte Leerzeichen, falls der Nachname fehlt
        let vollerName = `${vorname} ${nachname}`.trim();
        if (!vollerName) {
          vollerName = email; // Fallback, falls absolut kein Name eingetragen ist
        }

        // Mobilnummer prüfen. Falls leer, Standardtext setzen
        const mobileRaw = row[4] ? row[4].toString().trim() : '';
        const mobile = mobileRaw !== '' ? mobileRaw : 'Nicht hinterlegt';

        return {
          id: row[0] ? row[0].toString().trim() : 'Keine ID', // Spalte A: Mitglieder-ID
          vorname: vorname,   // Spalte B: Vorname
          nachname: nachname, // Spalte C: Name
          name: vollerName,   // Kombiniert ohne doppelte Leerzeichen
          email: row[3],      // Spalte D: E-Mail
          mobile: mobile      // Spalte E: Mobile
        };
      }
    }
    
    return null; // Keine Übereinstimmung gefunden
  } catch (e) {
    Logger.log('Fehler beim Einlesen der Mitgliederdaten: ' + e.message);
    return null;
  }
}

/**
 * Prüft den Kalender nach Terminen für den Folgetag und sendet E-Mail-Erinnerungen
 * an die buchenden Mitglieder anhand des Eintrags "Kontakt: E-Mail".
 */
function sendDailyReservationReminders() {
  Logger.log("=== STARTE TÄGLICHE ERINNERUNGS-PRÜFUNG ===");
  
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  
  // Berechne den morgigen Tag (Start: 00:00 Uhr, Ende: 23:59 Uhr)
  const tomorrowStart = new Date();
  tomorrowStart.setDate(tomorrowStart.getDate() + 1);
  tomorrowStart.setHours(0, 0, 0, 0);
  
  const tomorrowEnd = new Date();
  tomorrowEnd.setDate(tomorrowEnd.getDate() + 1);
  tomorrowEnd.setHours(23, 59, 59, 999);
  
  // Alle Events für morgen holen
  const events = calendar.getEvents(tomorrowStart, tomorrowEnd);
  Logger.log(`${events.length} Termine für morgen gefunden.`);
  
  events.forEach(event => {
    const description = event.getDescription() || "";
    
    // REGEX-MATCH: Sucht nach "Kontakt: " gefolgt von einer E-Mail-Adresse
    const emailMatch = description.match(/Kontakt:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    
    if (emailMatch && emailMatch[1]) {
      const memberEmail = emailMatch[1].trim();
      const slotName = event.getStartTime().getHours() === 6 ? "Vormittag (06:00 - 14:00)" : "Nachmittag (14:00 - 20:00)";
      
      const subject = `Erinnerung: Deine Boot Buchung für morgen!`;
      let body = `Hallo!\n\nDies ist die automatische Erinnerung für deine anstehende Reservierung:\n\n`;
      body += `\u{1F4C5} Datum: ${tomorrowStart.toLocaleDateString('de-CH')}\n`;
      body += `\u{23F0} Slot: ${slotName}\n\n`;
      body += `Viel Spass auf dem Wasser!\n\nDein Vorstand`;
      
      MailApp.sendEmail(memberEmail, subject, body);
      Logger.log(`   -> Erinnerung erfolgreich an ${memberEmail} gesendet.`);
    } else {
      Logger.log(`   -> Kein gültiger 'Kontakt:'-Eintrag im Event '${event.getTitle()}' gefunden.`);
    }
  });
  
  Logger.log("=== ERINNERUNGS-PRÜFUNG BEENDET ===");
}
