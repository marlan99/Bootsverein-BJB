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
  // TESTCASE 6: Erfolgreiche Stornierung
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 10 – Erfolgreiche Stornierung (Frist eingehalten)");
  cleanupOldTestMails();
  results.push(testSuccessfulCancellation());

  // -----------------------------------------------------------------
  // TESTCASE 7: Abgelehnte Stornierung
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 11 – Abgelehnte Stornierung (24h-Frist verletzt)");
  cleanupOldTestMails();
  results.push(testRejectedCancellation());

  // -----------------------------------------------------------------
  // TESTCASE 8: Europäische Datumsformate (NEU angepasst)
  // -----------------------------------------------------------------
  Logger.log("\n[START] Testcase: ID 12 – Flexibles europäisches Datums-Parsing");
  cleanupOldTestMails();
  results.push(testEuropeanDateFormats());

  // -----------------------------------------------------------------
  // TESTCASE 9: Skalierungstest (OPTIONAL)
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
  const testDate = getFutureDate(10, 'DOT_LEAD');
  createTestEmail({
    subject: 'Reservierung',
    body: `Datum: ${testDate}\nSlot: Vormittag\nTyp: Standard\nBeschreibung: Testlauf Hauptfunktion\nAnlass: Automatisierung`
  });

  labelTestEmails();
  processReservationEmails(); 

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();

  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
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
  const date1 = getFutureDate(3, 'DOT_LEAD');
  const date2 = getFutureDate(5, 'DOT_LEAD'); 

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
  const testDate = getFutureDate(12, 'DOT_LEAD');
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
  const tomorrowDate = getFutureDate(1, 'DOT_LEAD');
  
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
  const targetDate = getFutureDate(4, 'DOT_LEAD'); 
  
  createTestEmail({ body: `Datum: ${targetDate}\nSlot: Nachmittag\nTyp: Standard\nBeschreibung: Wird storniert` });
  labelTestEmails();
  processReservationEmails();
  
  createTestEmail({ 
    subject: 'Stornierung Boot', 
    body: `Datum: ${targetDate}\nSlot: Nachmittag` 
  });
  labelTestEmails();
  processReservationEmails(); 
  
  Utilities.sleep(2000);

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();
  
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const parts = targetDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
  const event = events.find(e => e.getTitle().includes(myName));
  
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
  const todayDate = getFutureDate(0, 'DOT_LEAD'); 
  
  createTestEmail({ body: `Datum: ${todayDate}\nSlot: Nachmittag\nTyp: Standard\nBeschreibung: Kurzfrist-Test` });
  labelTestEmails();
  processReservationEmails();
  
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
 * ID 12: Prüft europäische Datumsformate auf erfolgreiche Erkennung
 */
function testEuropeanDateFormats() {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();
  
  // Testet verschiedene typisch europäische Schreibweisen an unterschiedlichen Tagen im Kalender
  const formatsToTest = [
    { type: 'DOT_LEAD',   daysOut: 15, desc: '05.06.2026 (Mit führenden Nullen)' },
    { type: 'DOT_NO_LEAD', daysOut: 16, desc: '5.6.2026 (Ohne führende Nullen)' },
    { type: 'EU_SLASH',   daysOut: 17, desc: '5/6/2026 (EU-Schrägstrich Tag/Monat/Jahr)' },
    { type: 'DE_TEXT',    daysOut: 18, desc: '5. Juni 2026 (Textmonat mit Punkt)' }
  ];
  
  let successfulParses = 0;
  
  formatsToTest.forEach(item => {
    const formattedDateString = getFutureDate(item.daysOut, item.type);
    
    // 1. Sende E-Mail mit dem jeweiligen europäisch formatierten Datum
    createTestEmail({
      subject: 'Reservierung',
      body: `Datum: ${formattedDateString}\nSlot: Vormittag\nTyp: Standard\nBeschreibung: Europäisches Format-Test ${item.type}`
    });
    labelTestEmails();
    processReservationEmails();
    Utilities.sleep(1500);
    
    // 2. Prüfe, ob das Event am präzisen Zieltag eingetragen wurde
    const targetDateObj = new Date();
    targetDateObj.setDate(targetDateObj.getDate() + item.daysOut);
    
    const events = calendar.getEventsForDay(targetDateObj);
    const event = events.find(e => e.getTitle().includes(myName) && e.getDescription().includes(item.type));
    
    if (event) {
      successfulParses++;
    } else {
      Logger.log(`   -> [FEHLER] Europäisches Format fehlgeschlagen: "${item.desc}" mit generiertem Text: '${formattedDateString}'`);
    }
    
    // Aufräumen für den nächsten Formattest
    if (event) event.deleteEvent();
  });
  
  const passed = (successfulParses === formatsToTest.length);
  
  return {
    name: 'ID 12 – Europäische Datumsformate',
    passed: passed,
    message: passed 
      ? `Alle ${formatsToTest.length} europäischen Formate (Punkte, Ohne Null, Schrägstrich, Textmonat) wurden erfolgreich verarbeitet.` 
      : `${successfulParses} von ${formatsToTest.length} europäischen Formaten wurden korrekt erkannt.`
  };
}

/**
 * ID 7: Simuliert Last durch gleichzeitige Benutzeranfragen
 */
function testScalability() {
  const startTime = new Date();
  
  for (let i = 1; i <= 5; i++) { 
    const date = getFutureDate(20 + i, 'DOT_LEAD');
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
  
  const maxDateStr = getFutureDate(100, 'DOT_LEAD');
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
 * Generiert ein Zukunftsdatum basierend auf dem gewünschten Formattyp
 * @param {number} days - Tage in der Zukunft
 * @param {string} format - Typ des Datumsformats
 */
function getFutureDate(days, format) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  
  const year = d.getFullYear();
  const monthNum = d.getMonth() + 1;
  const dayNum = d.getDate();
  
  const monthLead = String(monthNum).padStart(2, '0');
  const dayLead = String(dayNum).padStart(2, '0');
  
  const deMonths = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];

  switch(format) {
    case 'DOT_NO_LEAD':
      return `${dayNum}.${monthNum}.${year}`; // z.B. 5.6.2026
    case 'EU_SLASH':
      return `${dayNum}/${monthNum}/${year}`; // z.B. 5/6/2026
    case 'DE_TEXT':
      return `${dayNum}. ${deMonths[d.getMonth()]} ${year}`; // z.B. 5. Juni 2026
    case 'DOT_LEAD':
    default:
      return `${dayLead}.${monthLead}.${year}`; // z.B. 05.06.2026
  }
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
      if (e.getTitle().includes(myName) || e.getDescription().includes('Lasttest') || e.getDescription().includes('Kurzfrist-Test') || e.getDescription().includes('Wird storniert') || e.getDescription().includes('Format-Test')) {
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
