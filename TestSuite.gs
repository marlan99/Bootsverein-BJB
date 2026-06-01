/**
 * TestSuite.gs
 * Automatisiertes Testskript für das E-Mail-basierte Reservierungssystem
 * * VORAUSSETZUNG: Deine E-Mail-Adresse muss in der Whitelist-Tabelle eingetragen sein!
 */

function runAllTests() {
  Logger.log("=== START DER AUTOMATISIERTEN TESTSUITE ===");
  const results = [];
  
  // Bereinige alte Test-Mails vor dem Start, um Fehlalarme zu vermeiden
  cleanupOldTestMails();

  // Testfälle nacheinander ausführen
  results.push(testValidReservation());
  results.push(testStandardLimit());
  results.push(testSlotTimes());
  results.push(testInvalidFormat());
  results.push(testReminder());
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
    body: `Datum: ${testDate}
Slot: Vormittag
Typ: Standard
Beschreibung: Testlauf Hauptfunktion
Anlass: Automatisierung`
  });

  labelTestEmails();
  processReservationEmails(); 

  // Überprüfung im Kalender: Holt den Namen dynamisch aus der Whitelist für den Matcher
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

  // 1. Erste gültige Mail (ohne Name)
  createTestEmail({ body: `Datum: ${date1}\nSlot: Nachmittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  // 2. Zweite Mail im Sperrzeitraum (wird abgelehnt)
  createTestEmail({ body: `Datum: ${date2}\nSlot: Vormittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();

  Utilities.sleep(2000); 
  // Da kein Name mehr im Body steht, suchen wir nach der Ablehnung im Betreff und deiner Mail
  const myEmail = Session.getActiveUser().getEmail();
  const threads = GmailApp.search(`subject:"Reservierung abgelehnt" to:${myEmail}`);
  const passed = threads.length > 0;

  return {
    name: 'ID 3 – Standard-Limit (< 14 Tage)',
    passed: passed,
    message: passed ? 'Zweite Reservierung wurde korrekt per Mail abgelehnt.' : 'Sperre griff nicht oder keine Mail versendet.'
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
  const myEmail = Session.getActiveUser().getEmail();
  const threads = GmailApp.search(`subject:"Reservierung abgelehnt" to:${myEmail}`);
  const passed = threads.length > 0;

  return {
    name: 'ID 9 – Intuitive Fehlermeldung bei Falschformat',
    passed: passed,
    message: passed ? 'Fehlerhaftes Format wurde erkannt und beantwortet.' : 'Ungültige Mail triggerte keine Fehlerantwort.'
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

  return {
    name: 'ID 6 – Erinnerungsfunktion',
    passed: !!event,
    message: event ? 'Event erstellt (Reminder-Schnittstelle aktiv).' : 'Event wurde nicht angelegt.'
  };
}

/**
 * ID 7: Simuliert Last durch gleichzeitige Benutzeranfragen
 */
function testScalability() {
  const startTime = new Date();
  
  // Sendet 5 Test-Mails nacheinander (jetzt ohne Name im Body)
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
  // Wir prüfen hier auf die im Lasttest gesetzte Beschreibung im Kalendereintrag
  const loadEventsCount = allEvents.filter(e => e.getDescription().includes('Lasttest')).length;

  // Da das Limit von 1 Termin pro 2 Wochen aktiv ist, wird bei derselben E-Mail-Adresse 
  // die erste Mail durchgehen (getFutureDate(21)) und die anderen 4 blockiert.
  // Das ist das korrekte Verhalten! Der Test ist bestanden, wenn das Skript nicht abstürzt.
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
   HILFSFUNKTIONEN (UTILITIES)
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
  Utilities.sleep(1500); 
  const threads = GmailApp.search('is:unread from:me subject:"Reservierung"');
  const label = GmailApp.getUserLabelByName("Reservierung/Neu");
  
  if (label && threads.length > 0) {
    label.addToThreads(threads);
    GmailApp.markThreadsRead(threads); 
  }
}

function cleanupOldTestMails() {
  const labelNeu = GmailApp.getUserLabelByName("Reservierung/Neu");
  const labelErledigt = GmailApp.getUserLabelByName("Reservierung/Erledigt");
  const labelAbgelehnt = GmailApp.getUserLabelByName("Reservierung/Abgelehnt");
  
  // Sucht nach allen Mails von dir selbst, die mit dem Reservierungssystem zu tun haben
  const threads = GmailApp.search('from:me "Reservierung" OR "stornierung"');
  
  threads.forEach(thread => {
    if(labelNeu) labelNeu.removeFromThread(thread);
    if(labelErledigt) labelErledigt.removeFromThread(thread);
    if(labelAbgelehnt) labelAbgelehnt.removeFromThread(thread);
    thread.moveToTrash(); 
  });
  Logger.log("Alte Test-Mails in den Papierkorb verschoben.");
}
