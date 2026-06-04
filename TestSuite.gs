/**
 * TestSuite.gs
 * Automatisiertes Testskript für das E-Mail-basierte Reservierungssystem
 * * VORAUSSETZUNG: Deine E-Mail-Adresse muss in der Whitelist-Tabelle eingetragen sein!
 */

// Konfiguration für die Testsuite
const TEST_CONFIG = {
  RUN_SCALABILITY_TEST: false // <--- Auf 'true' setzen, um den Skalierungstest auszuführen
};

function runAllTests() {
  Logger.log("=================================================================");
  Logger.log("=== START DER AUTOMATISIERTEN TESTSUITE ===");
  Logger.log("=================================================================");
  const results = [];
  
  // -----------------------------------------------------------------
  // TESTCASE 1
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 1,2,5 – Gültige Reservierung & Zusatzinfos");
  cleanupOldTestMails();
  results.push(testValidReservation());

  // -----------------------------------------------------------------
  // TESTCASE 2
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 3 – Standard-Limit (< 14 Tage)");
  cleanupOldTestMails();
  results.push(testStandardLimit());

  // -----------------------------------------------------------------
  // TESTCASE 3
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 8 – Slot-Zeiten (Vormittag = 06:00-14:00)");
  cleanupOldTestMails();
  results.push(testSlotTimes());

  // -----------------------------------------------------------------
  // TESTCASE 4
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 9 – Intuitive Fehlermeldung bei Falschformat");
  cleanupOldTestMails();
  results.push(testInvalidFormat());

  // -----------------------------------------------------------------
  // TESTCASE 5
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 6 – Erinnerungsfunktion");
  cleanupOldTestMails();
  results.push(testReminder());

  // -----------------------------------------------------------------
  // TESTCASE 6: Erfolgreiche Stornierung (NEU)
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 10 – Erfolgreiche Stornierung (Frist eingehalten)");
  cleanupOldTestMails();
  results.push(testSuccessfulCancellation());

  // -----------------------------------------------------------------
  // TESTCASE 7: Abgelehnte Stornierung (NEU)
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 11 – Abgelehnte Stornierung (24h-Frist verletzt)");
  cleanupOldTestMails();
  results.push(testRejectedCancellation());

  // -----------------------------------------------------------------
  // TESTCASE 8: Skalierungstest (OPTIONAL)
  // -----------------------------------------------------------------
  if (TEST_CONFIG.RUN_SCALABILITY_TEST) {
    Logger.log("\n[START] Testcase: ID 7 – Skalierungstest (Systemstabilität)");
    cleanupOldTestMails();
    results.push(testScalability());
  } else {
    Logger.log("\n[INFO] Testcase: ID 7 – Skalierungstest übersprungen (Deaktiviert in TEST_CONFIG)");
  }

  // -----------------------------------------------------------------
  // AUSWERTUNG & ZUSAMMENFASSUNG
  // -----------------------------------------------------------------
  const passed = results.filter(r => r.passed).length;
  const total = results.length;
  
  Logger.log("\n=========================================");
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
    Logger.log(`\nTestbericht erfolgreich an ${CONFIG.ADMIN_EMAIL} gesendet.`);
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

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  // Datumsobjekt für die Suche parsen
  const parts = testDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
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

  createTestEmail({ body: `Datum: ${date1}\nSlot: Nachmittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  createTestEmail({ body: `Datum: ${date2}\nSlot: Vormittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  Utilities.sleep(2000); 
  
  const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt');
  let passed = false;
  if (labelAbgelehnt) {
    const threads = labelAbgelehnt.getThreads(0, 10);
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
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const parts = testDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
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
  
  const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt');
  const passed = labelAbgelehnt ? labelAbgelehnt.getThreads().length > 0 : false;

  return {
    name: 'ID 9 – Intuitive Fehlermeldung bei Falschformat',
    passed: passed,
    message: passed ? 'Fehlerhaftes Format wurde erkannt und aussortiert.' : 'Ungültige Mail triggerte keine Fehlerbehandlung.'
  };
}

/**
 * ID 6: Prüft die automatisierte E-Mail-Erinnerung für den Folgetag
 */
function testReminder() {
  const tomorrowDate = getFutureDate(1);
  
  createTestEmail({ 
    body: `Datum: ${tomorrowDate}\nSlot: Nachmittag\nTyp: Standard\nBeschreibung: Testlauf Erinnerung` 
  });
  
  labelTestEmails();
  processReservationEmails(); 

  sendDailyReservationReminders();
  
  Utilities.sleep(3500);
  
  const threads = GmailApp.search(`subject:"Erinnerung: Deine Boot Buchung für morgen!" to:me`);
  const passed = threads.length > 0;

  return {
    name: 'ID 6 – Erinnerungsfunktion (E-Mail an Buchenden)',
    passed: passed,
    message: passed 
      ? 'Erinnerungs-E-Mail wurde erfolgreich generiert und an den Buchenden zugestellt.' 
      : 'Es wurde keine Erinnerungs-E-Mail im Postfach gefunden.'
  };
}

/**
 * ID 10: Prüft eine erfolgreiche Stornierung weit im Voraus (Frist eingehalten)
 */
function testSuccessfulCancellation() {
  const targetDate = getFutureDate(4); // 4 Tage in der Zukunft (Frist 24h locker eingehalten)
  
  // 1. Erst buchen
  createTestEmail({ body: `Datum: ${targetDate}\nSlot: Nachmittag\nTyp: Standard\nBeschreibung: Wird storniert` });
  labelTestEmails();
  processReservationEmails();
  
  // 2. Stornierungs-Mail senden
  createTestEmail({ 
    subject: 'Stornierung Boot', 
    body: `Datum: ${targetDate}\nSlot: Nachmittag` 
  });
  labelTestEmails();
  processReservationEmails(); // Verarbeitet Stornierung
  
  Utilities.sleep(2000);

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();
  
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const parts = targetDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
  const event = events.find(e => e.getTitle().includes(myName));
  
  // Bestanden, wenn kein Kalendereintrag mehr existiert
  const passed = !event;

  return {
    name: 'ID 10 – Erfolgreiche Stornierung (Frist eingehalten)',
    passed: passed,
    message: passed 
      ? 'Der Termin wurde nach der Stornierungsanfrage erfolgreich aus dem Kalender gelöscht.' 
      : 'Der Termin existiert trotz Stornierung weiterhin im Kalender.'
  };
}

/**
 * ID 11: Prüft die Ablehnung einer Stornierung am selben Tag (Frist verletzt)
 */
function testRejectedCancellation() {
  const todayDate = getFutureDate(0); // Heute buchen & stornieren versuchen -> Verletzt die 24h-Frist
  
  // 1. Provisorisch für heute buchen
  createTestEmail({ body: `Datum: ${todayDate}\nSlot: Nachmittag\nTyp: Standard\nBeschreibung: Kurzfrist-Test` });
  labelTestEmails();
  processReservationEmails();
  
  // 2. Sofort stornieren versuchen
  createTestEmail({ 
    subject: 'Absage Termin', 
    body: `Datum: ${todayDate}\nSlot: Nachmittag` 
  });
  labelTestEmails();
  processReservationEmails();
  
  Utilities.sleep(2000);
  
  const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt');
  const passed = labelAbgelehnt ? labelAbgelehnt.getThreads().length > 0 : false;

  return {
    name: 'ID 11 – Abgelehnte Stornierung (24h-Frist verletzt)',
    passed: passed,
    message: passed 
      ? 'Kurzfristige Stornierung wurde richtigerweise blockiert und die Anfrage zu "Abgelehnt" verschoben.' 
      : 'Die Stornierung wurde trotz verletzter Frist durchgeführt oder nicht korrekt einsortiert.'
  };
}

/**
 * ID 7: Simuliert Last durch gleichzeitige Benutzeranfragen
 */
function testScalability() {
  const startTime = new Date();
  
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
  
  // Datum für das Ende des Suchbereichs parsen
  const maxDateStr = getFutureDate(100);
  const parts = maxDateStr.split('.');
  const endDate = new Date(parts[2], parts[1] - 1, parts[0]);
  
  const allEvents = calendar.getEvents(startTime, endDate);
  const loadEventsCount = allEvents.filter(e => e.getDescription().includes('Lasttest')).length;
  
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
   HILFSFUNKTIONEN (UTILITIES) – MIT TEST-ARCHIV LABELLING
   ========================================================================== */

/**
 * Generiert ein Zukunftsdatum direkt im neuen Format DD.MM.YYYY
 */
function getFutureDate(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${day}.${month}.${year}`; // Geändert auf DD.MM.YYYY
}

function createTestEmail({subject = 'Reservierung', body}) {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
}

function labelTestEmails() {
  Utilities.sleep(3500); 
  const threads = GmailApp.search('is:unread from:me (subject:"Reservierung" OR subject:"Stornierung" OR subject:"Absage")');
  const label = GmailApp.getUserLabelByName("Reservierung/Neu");
  
  if (label && threads.length > 0) {
    label.addToThreads(threads);
  }
}

function cleanupOldTestMails() {
  const labelNeu = GmailApp.getUserLabelByName("Reservierung/Neu");
  const labelErledigt = GmailApp.getUserLabelByName("Reservierung/Erledigt");
  const labelAbgelehnt = GmailApp.getUserLabelByName("Reservierung/Abgelehnt");
  
  const archivLabelName = "Reservierung/Test-Archiv";
  let labelArchiv = GmailApp.getUserLabelByName(archivLabelName);
  if (!labelArchiv) {
    labelArchiv = GmailApp.createLabel(archivLabelName);
  }
  
  const threads = GmailApp.search('from:me "Reservierung" OR "stornierung" OR "Buchung" OR "Absage"');
  
  threads.forEach(thread => {
    if(labelNeu) labelNeu.removeFromThread(thread);
    if(labelErledigt) labelErledigt.removeFromThread(thread);
    if(labelAbgelehnt) labelAbgelehnt.removeFromThread(thread);
    
    labelArchiv.addToThreads([thread]);
    thread.moveToArchive();
    thread.markRead();
  });
  
  Logger.log("   -> Mails nach 'Reservierung/Test-Archiv' verschoben.");

  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const start = new Date();
    const end = new Date();
    end.setDate(start.getDate() + 40); 
    
    const events = calendar.getEvents(start, end);
    const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
    const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();

    let deletedCount = 0;
    events.forEach(e => {
      if (e.getTitle().includes(myName) || e.getDescription().includes('Lasttest') || e.getDescription().includes('Kurzfrist-Test') || e.getDescription().includes('Wird storniert')) {
        e.deleteEvent();
        deletedCount++;
      }
    });
    if (deletedCount > 0) {
      Logger.log(`   -> ${deletedCount} alte(s) Test-Kalenderevent(s) gelöscht.`);
    }
  } catch(e) {
    Logger.log("   -> Hinweis bei Kalenderbereinigung: " + e.message);
  }
  
  Utilities.sleep(2000); 
}
