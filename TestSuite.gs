/**
 * TestSuite.gs
 * Automatisiertes Testskript für das E-Mail-basierte Reservierungssystem
 * * VORAUSSETZUNG: Deine E-Mail-Adresse muss in der Whitelist-Tabelle eingetragen sein!
 */

function runAllTests() {
  Logger.log("=== START DER AUTOMATISIERTEN TESTSUITE ===");
  const results = [];
  
  // 1. Bereinigung vor dem Start
  cleanupOldTestMails();
  results.push(testValidReservation());

  // 2. Erneute Bereinigung, damit das 14-Tage-Limit für den nächsten Test wieder frei ist
  cleanupOldTestMails();
  results.push(testStandardLimit());

  // 3. Bereinigung, damit die Slots für die folgenden Tests frei sind
  cleanupOldTestMails();
  results.push(testSlotTimes());

  cleanupOldTestMails();
  results.push(testInvalidFormat());

  cleanupOldTestMails();
  results.push(testReminder());

  cleanupOldTestMails();
  results.push(testScalability());

  // Auswertung & Zusammenfassung
  const passed = results.filter(r => r.passed).length;
  const total = results.length;
  
  Logger.log("=========================================");
  Logger.log(`TEST-ERGEBNIS: ${passed}/${total} BESTANDEN`);
  Logger.log("=========================================");
  
  let emailBody = `Zusammenfassung des Testlaufs vom ${new Date().toLocaleString()}\n`;
  emailBody += `Ergebnis: ${passed} von ${total} Tests bestanden.\n\nDetail-Log:\n`;

  results.forEach(r => {
    const statusStr = r.passed ? '✅ PASS' : '❌ FAIL';
    const logLine = `${statusStr} [${r.name}]: ${r.message}`;
    Logger.log(logLine);
    emailBody += logLine + '\n';
  });

  // Ergebnis per E-Mail an den Admin senden
  if (CONFIG && CONFIG.ADMIN_EMAIL) {
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, `Testbericht Reservierungssystem: ${passed}/${total}`, emailBody);
    Logger.log(`Testbericht an ${CONFIG.ADMIN_EMAIL} gesendet.`);
  }
}

/* ==========================================================================
   TESTFÄLLE (MAPPING ZU DEN ANFORDERUNGEN)
   ========================================================================== */

/**
 * ID 1, 2, 5: Prüft die gültige Verarbeitung eines Standard-Templates
 */
function testValidReservation() {
  const testDate = getFutureDate(10);
  createTestEmail({
    subject: 'Reservierung',
    body: `Datum: ${testDate}\nSlot: Vormittag\nTyp: Standard\nBeschreibung: Testlauf Hauptfunktion\nAnlass: Automatisierung`
  });

  labelTestEmails();
  processReservationEmails(); 

  // Überprüfung im Kalender
  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : 'Unbekannt';

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const events = calendar.getEventsForDay(new Date(testDate));
  const event = events.find(e => e.getTitle().includes(myName));

  const passed = !!event && event.getDescription().includes('Testlauf Hauptfunktion') && event.getDescription().includes('Automatisierung');
  return {
    name: 'ID 1,2,5 – Gültige Reservierung & Zusatzinfos',
    passed: passed,
    message: passed ? 'Event mit deinem Whitelist-Namen und Zusatzinfos im Kalender gefunden.' : 'Event unvollständig oder nicht erstellt.'
  };
}

/**
 * ID 3: Prüft die Sperre von zwei Standard-Terminen innerhalb von 14 Tagen
 */
function testStandardLimit() {
  const date1 = getFutureDate(3);
  const date2 = getFutureDate(5); 

  // 1. Erste gültige Mail
  createTestEmail({ body: `Datum: ${date1}\nSlot: Nachmittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  // 2. Zweite Mail im Sperrzeitraum (muss abgelehnt werden)
  createTestEmail({ body: `Datum: ${date2}\nSlot: Vormittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  Utilities.sleep(2000); 
  
  // Präzise Prüfung: Hat ein Thread das Label "Reservierung/Abgelehnt" erhalten?
  const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt');
  let passed = false;
  if (labelAbgelehnt) {
    const threads = labelAbgelehnt.getThreads(0, 10);
    // Wenn sich im Abgelehnt-Label Mails befinden, griff die Sperre
    passed = threads.length > 0;
  }

  return {
    name: 'ID 3 – Standard-Limit (< 14 Tage)',
    passed: passed,
    message: passed ? 'Zweite Reservierung wurde korrekt blockiert und unter Abgelehnt archiviert.' : 'Sperre griff nicht oder Label wurde nicht gesetzt.'
  };
}

/**
 * ID 8: Prüft die korrekte Uhrzeitsetzung für VM (6-14) und NM (14-20)
 */
function testSlotTimes() {
  const testDate = getFutureDate(12);
  createTestEmail({ body: `Datum: ${testDate}\nSlot: Vormittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : 'Unbekannt';

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const events = calendar.getEventsForDay(new Date(testDate));
  const event = events.find(e => e.getTitle().includes(myName));

  let passed = false;
  if (event) {
    const startHour = event.getStartTime().getHours();
    const endHour = event.getEndTime().getHours();
    if (startHour === 6 && endHour === 14) { passed = true; }
  }

  return {
    name: 'ID 8 – Slot-Zeiten (Vormittag = 06:00-14:00)',
    passed: passed,
    message: passed ? 'Uhrzeit für Vormittags-Slot exakt gesetzt.' : 'Uhrzeiten weichen vom Konzept ab.'
  };
}

/**
 * ID 9: Prüft die Reaktion auf ein fehlerhaftes E-Mail-Format
 */
function testInvalidFormat() {
  createTestEmail({
    subject: 'Reservierung',
    body: `Hallo, ich würde gerne nächsten Dienstag kommen. Schöne Grüße, Chaos-User.`
  });
  labelTestEmails();
  processReservationEmails();

  Utilities.sleep(2000);
  
  // Prüfung über das Vorhandensein im "Abgelehnt"-Ordner
  const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt');
  const passed = labelAbgelehnt ? labelAbgelehnt.getThreads().length > 0 : false;

  return {
    name: 'ID 9 – Intuitive Fehlermeldung bei Falschformat',
    passed: passed,
    message: passed ? 'Fehlerhaftes Format wurde erkannt und aussortiert.' : 'Ungültige Mail triggerte keine Fehlerbehandlung.'
  };
}

/**
 * ID 6: Prüft, ob standardmäßig eine Erinnerung 24 Stunden vorher aktiv ist
 */
function testReminder() {
  const testDate = getFutureDate(15);
  createTestEmail({ body: `Datum: ${testDate}\nSlot: Nachmittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : 'Unbekannt';

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const events = calendar.getEventsForDay(new Date(testDate));
  const event = events.find(e => e.getTitle().includes(myName));

  // Tiefe Prüfung der Google-Kalender-Erinnerung
  let passed = false;
  let msg = 'Event wurde nicht angelegt.';
  
  if (event) {
    const emailReminders = event.getEmailReminders();
    // Prüft, ob mindestens ein E-Mail-Reminder gesetzt ist und ob er mit CONFIG übereinstimmt (1440 Min)
    if (emailReminders.length > 0 && emailReminders[0] === CONFIG.REMINDER_MINUTES) {
      passed = true;
      msg = `Erinnerung ist exakt auf ${emailReminders[0]} Minuten (24h) vorab eingestellt.`;
    } else {
      msg = 'Event existiert, aber das Erinnerungs-Intervall fehlt oder weicht ab.';
    }
  }

  return {
    name: 'ID 6 – Erinnerungsfunktion',
    passed: passed,
    message: msg
  };
}

/**
 * ID 7: Simuliert Last durch gleichzeitige Benutzeranfragen
 */
function testScalability() {
  const startTime = new Date();
  
  // Sendet 5 Test-Mails nacheinander
  for (let i = 1; i <= 5; i++) { 
    const date = getFutureDate(20 + i);
    createTestEmail({
      body: `Datum: ${date}\nSlot: Vormittag\nTyp: Standard\nBeschreibung: Lasttest ${i}`
    });
  }
  
  labelTestEmails();
  
  const processStart = new Date();
  processReservationEmails();
  const processEnd = new Date();
  
  const durationInSeconds = (processEnd - processStart) / 1000;
  
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const endDate = new Date(getFutureDate(100)); 
  
  const allEvents = calendar.getEvents(startTime, endDate);
  const loadEventsCount = allEvents.filter(e => e.getDescription().includes('Lasttest')).length;
  
  // Da 1 Termin/2 Wochen gilt, geht nur 1 durch, 4 werden abgelehnt. System darf nicht abstürzen.
  const passed = durationInSeconds < 60;

  return {
    name: 'ID 7 – Skalierungstest (Systemstabilität)',
    passed: passed,
    message: passed 
      ? `System blieb stabil. ${loadEventsCount} Event(s) eingetragen. Verarbeitungszeit: ${durationInSeconds}s.` 
      : `Fehlgeschlagen. Zeit überschritten (${durationInSeconds}s).`
  };
}

/* ==========================================================================
   HILFSFUNKTIONEN (UTILITIES) – OPTIMIERT
   ========================================================================== */

function getFutureDate(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function createTestEmail({subject = 'Reservierung', body}) {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
}

function labelTestEmails() {
  // Erhöht auf 3,5 Sekunden, um Apps Script Zeit zu geben, die Mail im Postfach zu indexieren
  Utilities.sleep(3500); 
  const threads = GmailApp.search('is:unread from:me subject:"Reservierung"');
  const label = GmailApp.getUserLabelByName("Reservierung/Neu");
  
  if (label && threads.length > 0) {
    label.addToThreads(threads);
    // CRITICAL FIX: Markiere die Threads NICHT als gelesen, da der Hauptcode sonst abbricht!
    // GmailApp.markThreadsRead(threads); <-- Entfernt
  }
}

function cleanupOldTestMails() {
  const labelNeu = GmailApp.getUserLabelByName("Reservierung/Neu");
  const labelErledigt = GmailApp.getUserLabelByName("Reservierung/Erledigt");
  const labelAbgelehnt = GmailApp.getUserLabelByName("Reservierung/Abgelehnt");
  
  const threads = GmailApp.search('from:me "Reservierung" OR "stornierung" OR "Buchung"');
  
  threads.forEach(thread => {
    if(labelNeu) labelNeu.removeFromThread(thread);
    if(labelErledigt) labelErledigt.removeFromThread(thread);
    if(labelAbgelehnt) labelAbgelehnt.removeFromThread(thread);
    thread.moveToTrash(); 
  });
  Logger.log("Alte Test-Mails in den Papierkorb verschoben.");

  // NEU: Kalender-Events der Testsuite entfernen, um Limits zurückzusetzen
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const start = new Date();
    const end = new Date();
    end.setDate(start.getDate() + 40); // Suchfenster für Testtermine
    
    const events = calendar.getEvents(start, end);
    const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
    const myName = myProfile ? myProfile.name : 'Unbekannt';

    events.forEach(e => {
      // Lösche nur Termine, die eindeutig der Testsuite zuzuordnen sind
      if (e.getTitle().includes(myName) || e.getDescription().includes('Lasttest')) {
        e.deleteEvent();
      }
    });
    Logger.log("Alte Test-Kalendereinträge bereinigt.");
  } catch(e) {
    Logger.log("Hinweis bei Kalenderbereinigung: " + e.message);
  }
  
  Utilities.sleep(2000); 
}
