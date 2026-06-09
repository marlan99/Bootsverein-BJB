/**
 * TestSuiteAndDebug.gs
 * Automatisiertes Testskript & Whitelist-Check für das E-Mail-basierte Reservierungssystem
 * MIT ERWEITERTEM TIMEOUT-SCHUTZ (Europäische Formate in Einzelschritte zerlegt)
 * * ANLEITUNG: Ersetze den gesamten Inhalt deiner TestSuite.gs mit diesem Code.
 */

// ==========================================================================
// GLOBALE KONFIGURATION FÜR TESTS & DEBUGGING
// ==========================================================================

// Hier die E-Mail-Adresse eintragen, die bei den Tests und beim Whitelist-Check geprüft werden soll:
const DEBUG_EMAIL = "";
// Nutzt jetzt automatisch deinen aktiven Google-Account, wenn das Feld leer bleibt

// Konfiguration für die Testsuite: Hier kannst du jeden Test einzeln steuern
const TEST_CONFIG = {
  RUN_TEST_WHITELIST_CHECK: true,         // ID 0 – Prüft, ob DEBUG_EMAIL in der Whitelist existiert
  RUN_TEST_VALID_RESERVATION: true,       // ID 1,2,5 – Gültige Reservierung & Zusatzinfos
  RUN_TEST_STANDARD_LIMIT: true,          // ID 3 – Saison-Limit (1 aktiver Termin parallel)
  RUN_TEST_SLOT_TIMES: true,              // ID 8 – Slot-Zeiten (Vormittag = 08:00-14:00)
  RUN_TEST_INVALID_FORMAT: true,          // ID 9 – Intuitive Fehlermeldung bei Falschformat
  RUN_TEST_REMINDER: true,                // ID 6 – Erinnerungsfunktion
  RUN_TEST_SUCCESSFUL_CANCELLATION: true, // ID 10 – Erfolgreiche Stornierung
  RUN_TEST_REJECTED_CANCELLATION: true,   // ID 11 – Abgelehnte Stornierung (24h-Frist)
  RUN_TEST_EUROPEAN_DATE_FORMATS: true,   // ID 12 – Flexibles europäisches Datums-Parsing
  RUN_SCALABILITY_TEST: true              // ID 7 – Skalierungstest (Systemstabilität)
};

// ==========================================================================
// HAUPTFUNKTION DER TESTSUITE (TIMEOUT-SICHER)
// ==========================================================================

function runAllTests() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME = 4.0 * 60 * 1000; // Auf 4 Minuten verkürzt für maximale Sicherheit
  
  Logger.log("=================================================================");
  Logger.log("=== START DER AUTOMATISIERTEN TESTSUITE (TIMEOUT-SCHUTZ) ===");
  Logger.log("=================================================================");

  // Überprüfen, ob das Hauptsystem bereits einmal initialisiert wurde
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  if (!sheetId) {
    Logger.log("❌ KRITISCHER FEHLER: Keine 'SHEET_CONFIG_ID' in den Skripteigenschaften gefunden.");
    return;
  }

  // Zustand aus den UserProperties laden
  const userProperties = PropertiesService.getUserProperties();
  let currentTestIndex = parseInt(userProperties.getProperty('TS_CURRENT_INDEX')) || 0;
  let results = JSON.parse(userProperties.getProperty('TS_RESULTS')) || [];

  // Bereinige alte Trigger für diese Funktion
  deleteTrigger('runAllTests');

  // Definition aller Testlauf-Schritte.
  // Die europäischen Formate sind jetzt Einzelschritte!
  const testCases = [
    { 
      name: 'ID 0 – Whitelist-Eintrag für Test-Mail prüfen',
      config: TEST_CONFIG.RUN_TEST_WHITELIST_CHECK,
      exec: () => testWhitelistCheck()
    },
    { 
      name: 'ID 1,2,5 – Gültige Reservierung & Zusatzinfos',
      config: TEST_CONFIG.RUN_TEST_VALID_RESERVATION,
      exec: () => { cleanupOldTestMails();
                    return testValidReservation(); }
    },
    { 
      name: 'ID 3 – Saison-Limit (Max. 1 aktiver Standard-Termin parallel)',
      config: TEST_CONFIG.RUN_TEST_STANDARD_LIMIT,
      exec: () => { cleanupOldTestMails();
                    return testStandardLimit(); }
    },
    { 
      name: 'ID 8 – Slot-Zeiten (Vormittag = 08:00-14:00)',
      config: TEST_CONFIG.RUN_TEST_SLOT_TIMES,
      exec: () => { cleanupOldTestMails();
                    return testSlotTimes(); }
    },
    { 
      name: 'ID 9 – Intuitive Fehlermeldung bei Falschformat',
      config: TEST_CONFIG.RUN_TEST_INVALID_FORMAT,
      exec: () => { cleanupOldTestMails();
                    return testInvalidFormat(); }
    },
    { 
      name: 'ID 6 – Erinnerungsfunktion (E-Mail an Buchenden)',
      config: TEST_CONFIG.RUN_TEST_REMINDER,
      exec: () => { cleanupOldTestMails();
                    return testReminder(); }
    },
    { 
      name: 'ID 10 – Erfolgreiche Stornierung (Frist eingehalten)',
      config: TEST_CONFIG.RUN_TEST_SUCCESSFUL_CANCELLATION,
      exec: () => { cleanupOldTestMails();
                    return testSuccessfulCancellation(); }
    },
    { 
      name: 'ID 11 – Abgelehnte Stornierung (24h-Frist verletzt)',
      config: TEST_CONFIG.RUN_TEST_REJECTED_CANCELLATION,
      exec: () => { cleanupOldTestMails();
                    return testRejectedCancellation(); }
    },
    
    // ID 12: Europäische Datumsformate in 4 Einzelschritte zerlegt, damit der Trigger dazwischenfunken kann
    { 
      name: 'ID 12 – Datumsformat: DOT_LEAD (05.06.2026)',
      config: TEST_CONFIG.RUN_TEST_EUROPEAN_DATE_FORMATS,
      exec: () => executeSingleDateFormatTest({ type: 'DOT_LEAD', daysOut: 6, slot: 'Vormittag', desc: 'Mit führenden Nullen' })
    },
    { 
      name: 'ID 12 – Datumsformat: DOT_NO_LEAD (5.6.2026)',
      config: TEST_CONFIG.RUN_TEST_EUROPEAN_DATE_FORMATS,
      exec: () => executeSingleDateFormatTest({ type: 'DOT_NO_LEAD', daysOut: 7, slot: 'Nachmittag', desc: 'Ohne führende Nullen' })
    },
    { 
      name: 'ID 12 – Datumsformat: EU_SLASH (5/6/2026)',
      config: TEST_CONFIG.RUN_TEST_EUROPEAN_DATE_FORMATS,
      exec: () => executeSingleDateFormatTest({ type: 'EU_SLASH', daysOut: 8, slot: 'Vormittag', desc: 'EU-Schrägstrich Tag/Monat/Jahr' })
    },
    { 
      name: 'ID 12 – Datumsformat: DE_TEXT (5. Juni 2026)',
      config: TEST_CONFIG.RUN_TEST_EUROPEAN_DATE_FORMATS,
      exec: () => executeSingleDateFormatTest({ type: 'DE_TEXT', daysOut: 9, slot: 'Nachmittag', desc: 'Textmonat mit Punkt' })
    },
    
    // Letzter Test
    { 
      name: 'ID 7 – Skalierungstest (Systemstabilität)',
      config: TEST_CONFIG.RUN_SCALABILITY_TEST,
      exec: () => { cleanupOldTestMails();
                    return testScalability(); }
    }
  ];

  // Iteration durch die Testfälle ab dem gespeicherten Index
  for (let i = currentTestIndex; i < testCases.length; i++) {
    const tc = testCases[i];
    // ZEIT-CHECK: Vor jedem Schritt prüfen, ob die Zeit knapp wird
    if (new Date().getTime() - startTime > MAX_RUNTIME) {
      Logger.log(`\n⚠️ ZEITLIMIT REICHT NICHT. Pausiere vor Schritt ${i}: (${tc.name}).`);
      userProperties.setProperty('TS_CURRENT_INDEX', i.toString());
      userProperties.setProperty('TS_RESULTS', JSON.stringify(results));
      
      ScriptApp.newTrigger('runAllTests')
               .timeBased()
               .after(1 * 60 * 1000)
               .create();
      Logger.log("⏰ Automatischen Folge-Trigger erstellt. Skript wird in 60 Sekunden fortgesetzt.");
      return;
    }

    if (tc.config) {
      Logger.log(`\n[START] Testcase: ${tc.name}`);
      try {
        results.push(tc.exec());
      } catch (err) {
        Logger.log(`❌ FEHLER bei Testausführung: ${err.message}`);
        results.push({ name: tc.name, passed: false, message: 'Laufzeitfehler: ' + err.message });
      }
    } else {
      Logger.log(`\n[INFO] Testcase: ${tc.name} – Übersprungen`);
      results.push({ name: tc.name, skipped: true });
    }
  }

  // ==========================================================================
  // AUSWERTUNG & FINISH
  // ==========================================================================
  const activeResults = results.filter(r => !r.skipped);
  const passed = activeResults.filter(r => r.passed).length;
  const totalActive = activeResults.length;
  const skipped = results.filter(r => r.skipped).length;
  
  Logger.log("\n=========================================");
  Logger.log(`TEST-ERGEBNIS: ${passed}/${totalActive} BESTANDEN (${skipped} übersprungen)`);
  Logger.log("=========================================");
  
  let emailBody = `Zusammenfassung des Testlaufs vom ${new Date().toLocaleString()}\n`;
  emailBody += `Ergebnis: ${passed} von ${totalActive} ausgeführten Tests bestanden. (${skipped} übersprungen)\n\nDetail-Log:\n`;
  results.forEach(r => {
    let statusStr = '';
    let msg = r.message || 'Test wurde in der Konfiguration deaktiviert.';
    if (r.skipped) {
      statusStr = '⚪ SKIPPED';
    } else {
      statusStr = r.passed ? '✅ PASS' : '❌ FAIL';
    }
    const logLine = `${statusStr} [${r.name}]: ${msg}`;
    Logger.log(logLine);
    emailBody += logLine + '\n';
  });
  if (typeof CONFIG !== 'undefined' && CONFIG.ADMIN_EMAIL && totalActive > 0) {
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, `Testbericht Reservierungssystem: ${passed}/${totalActive}`, emailBody);
    Logger.log(`\nTestbericht erfolgreich an ${CONFIG.ADMIN_EMAIL} gesendet.`);
  }

  userProperties.deleteProperty('TS_CURRENT_INDEX');
  userProperties.deleteProperty('TS_RESULTS');
}

/**
 * Setzt den Fortschritt der Testsuite komplett zurück,
 * sodass der nächste Lauf wieder ganz von vorne beginnt.
 */
function resetTestSuite() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('TS_CURRENT_INDEX');
  userProperties.deleteProperty('TS_RESULTS');
  
  // Löscht auch eventuell noch wartende automatische Folge-Trigger
  deleteTrigger('runAllTests');
  
  Logger.log("🔄 Testsuite erfolgreich zurückgesetzt! Der nächste Start von 'runAllTests' beginnt von vorn.");
}

/**
 * Hilfsfunktion: Löscht bestehende Trigger für eine bestimmte Funktion
 */
function deleteTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/* ==========================================================================
   NEUE HILFSFUNKTION FÜR DIE ZERLEGTEN DATUMS-TESTS
   ========================================================================== */

function executeSingleDateFormatTest(item) {
  cleanupOldTestMails();

  // WECHSEL ZUM STANDARDKALENDER, FALLS CONFIG.CALENDAR_ID LEER IST
  const calendar = (typeof CONFIG !== 'undefined' && CONFIG.CALENDAR_ID) ? 
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    
  // SICHERHEITS-CHECK: Falls der Kalender immer noch null ist (z.B. wegen Tippfehler in der ID)
  if (!calendar) {
    Logger.log("❌ KRITISCHER FEHLER im Format-Test: Kalender konnte nicht geladen werden. Bitte CALENDAR_ID überprüfen.");
    return {
      name: `ID 12 – Europäisches Format: ${item.type}`,
      passed: false,
      message: "Test abgebrochen: Kalender konnte nicht geladen werden."
    };
  }

  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();
  const targetDateObj = new Date();
  targetDateObj.setDate(targetDateObj.getDate() + item.daysOut);
  
  // Kalender vorab für diesen Tag reinigen
  const existingEvents = calendar.getEventsForDay(targetDateObj);
  existingEvents.forEach(e => {
    if (e.getTitle().includes(myName) || e.getDescription().includes('Format-Test')) e.deleteEvent();
  });
  Utilities.sleep(500);

  const formattedDateString = getFutureDate(item.daysOut, item.type);
  createTestEmail({
    subject: 'Reservierung',
    body: `Datum: ${formattedDateString}\nSlot: ${item.slot}\nTyp: Standard\nBeschreibung: Europäisches Format-Test ${item.type}`
  });
  
  labelTestEmails();
  processReservationEmails();
  Utilities.sleep(2000); 

  const events = calendar.getEventsForDay(targetDateObj);
  const event = events.find(e => e.getTitle().includes(myName) && e.getDescription().includes(item.type));
  
  const passed = !!event;
  if (event) {
    event.deleteEvent();
    Utilities.sleep(1500); 
  }
  
  cleanupOldTestMails();
  return {
    name: `ID 12 – Europäisches Format: ${item.type}`,
    passed: passed,
    message: passed ?
      `Format '${formattedDateString}' (${item.desc}) erfolgreich erkannt.` : `Format fehlgeschlagen: '${formattedDateString}'`
  };
}


/* ==========================================================================
   URSPRÜNGLICHE TESTFÄLLE (UNVERÄNDERT, testEuropeanDateFormats entfernt)
   ========================================================================== */

/**
 * ID 0: Testfall für die Whitelist-Prüfung
 * Nutzt DEBUG_EMAIL oder fällt auf die eigene Account-E-Mail zurück, wenn das Feld leer ist.
 */
function testWhitelistCheck() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  // Dynamische Ermittlung der Test-E-Mail (Fallback auf eigenen Account)
  const targetEmail = (typeof DEBUG_EMAIL !== 'undefined' && DEBUG_EMAIL.trim() !== "") 
    ? DEBUG_EMAIL.trim().toLowerCase() 
    : Session.getActiveUser().getEmail().trim().toLowerCase();

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { name: 'ID 0 – Whitelist-Eintrag für Test-Mail prüfen', passed: false, message: 'Tabelle ist leer oder enthält nur Kopfzeilen.' };
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let gefunden = false;
    
    for (let i = 0; i < dataRange.length; i++) {
      if (!dataRange[i][3]) continue;
      if (dataRange[i][3].toString().trim().toLowerCase() === targetEmail) {
        gefunden = true;
        break;
      }
    }
    
    return {
      name: 'ID 0 – Whitelist-Eintrag für Test-Mail prüfen',
      passed: gefunden,
      message: gefunden 
        ? `Adresse "${targetEmail}" erfolgreich in der Whitelist gefunden.` 
        : `Adresse "${targetEmail}" fehlt in Spalte D der Tabelle.`
    };
  } catch (e) {
    return { name: 'ID 0 – Whitelist-Eintrag für Test-Mail prüfen', passed: false, message: 'Fehler beim Tabellenzugriff: ' + e.message };
  }
}

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

  // Verwende hier ebenfalls den Fallback
  const calendar = (typeof CONFIG !== 'undefined' && CONFIG.CALENDAR_ID) ? 
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();

  if (!calendar) return { name: 'ID 1,2,5 – Gültige Reservierung & Zusatzinfos', passed: false, message: 'Kalender konnte nicht geladen werden.' };

  const parts = testDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
  const event = events.find(e => e.getTitle().includes(myName));

  const passed = !!event && event.getDescription().includes('Testlauf Hauptfunktion') && event.getDescription().includes('Automatisierung');
  return {
    name: 'ID 1,2,5 – Gültige Reservierung & Zusatzinfos',
    passed: passed,
    message: passed ?
      'Event mit deinem Whitelist-Namen und Zusatzinfos im Kalender gefunden.' : 'Event unvollständig oder nicht erstellt.'
  };
}

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
    name: 'ID 3 – Saison-Limit (Max. 1 aktiver Standard-Termin parallel)',
    passed: passed,
    message: passed ?
      'Zweite Reservierung wurde blockiert, da bereits ein aktiver Termin in der Saison existiert.' : 'Sperre für parallele Termine griff nicht.'
  };
}

function testSlotTimes() {
  const testDate = getFutureDate(12, 'DOT_LEAD');
  createTestEmail({ body: `Datum: ${testDate}\nSlot: Vormittag\nTyp: Standard` });
  labelTestEmails();
  processReservationEmails();
  const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
  const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();

  const calendar = (typeof CONFIG !== 'undefined' && CONFIG.CALENDAR_ID) ? 
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();

  if (!calendar) return { name: 'ID 8 – Slot-Zeiten (Vormittag = 08:00-14:00)', passed: false, message: 'Kalender konnte nicht geladen werden.' };

  const parts = testDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
  const event = events.find(e => e.getTitle().includes(myName));
  let passed = false;
  if (event) {
    const startHour = event.getStartTime().getHours();
    const endHour = event.getEndTime().getHours();
    if (startHour === 8 && endHour === 14) { passed = true; }
  }

  return {
    name: 'ID 8 – Slot-Zeiten (Vormittag = 08:00-14:00)',
    passed: passed,
    message: passed ?
      'Uhrzeit für Vormittags-Slot exakt gesetzt.' : 'Uhrzeiten weichen vom Konzept ab.'
  };
}

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
    message: passed ?
      'Fehlerhaftes Format wurde erkannt und aussortiert.' : 'Ungültige Mail triggerte keine Fehlerbehandlung.'
  };
}

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
    message: passed ?
      'Erinnerungs-E-Mail wurde erfolgreich generiert und an den Buchenden zugestellt.' : 'Es wurde keine Erinnerungs-E-Mail im Postfach gefunden.'
  };
}

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
  
  const calendar = (typeof CONFIG !== 'undefined' && CONFIG.CALENDAR_ID) ? 
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();

  if (!calendar) return { name: 'ID 10 – Erfolgreiche Stornierung (Frist eingehalten)', passed: false, message: 'Kalender konnte nicht geladen werden.' };

  const parts = targetDate.split('.');
  const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
  const events = calendar.getEventsForDay(parsedDate);
  const event = events.find(e => e.getTitle().includes(myName));
  
  const passed = !event;
  return {
    name: 'ID 10 – Erfolgreiche Stornierung (Frist eingehalten)',
    passed: passed,
    message: passed ?
      'Der Termin wurde nach der Stornierungsanfrage erfolgreich aus dem Kalender gelöscht.' : 'Der Termin existiert trotz Stornierung weiterhin im Kalender.'
  };
}

function testRejectedCancellation() {
  const tomorrowDate = getFutureDate(1, 'DOT_LEAD');  
  createTestEmail({ body: `Datum: ${tomorrowDate}\nSlot: Nachmittag\nTyp: Standard\nBeschreibung: Kurzfrist-Test` });
  labelTestEmails();
  processReservationEmails();
  createTestEmail({ 
    subject: 'Stornierung Termin', 
    body: `Datum: ${tomorrowDate}\nSlot: Nachmittag` 
  });
  labelTestEmails();
  processReservationEmails();
  Utilities.sleep(2000);
  
  const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt');
  const passed = labelAbgelehnt ? labelAbgelehnt.getThreads().length > 0 : false;
  return {
    name: 'ID 11 – Abgelehnte Stornierung (24h-Frist verletzt)',
    passed: passed,
    message: passed ?
      'Kurzfristige Stornierung wurde richtigerweise blockiert und die Anfrage zu "Abgelehnt" verschoben.'
    : 'Die Stornierung wurde trotz verletzter Frist durchgeführt oder nicht korrekt einsortiert.'
  };
}

function testScalability() {
  const startTime = new Date();
  for (let i = 1; i <= 2; i++) { 
    const date = getFutureDate(20 + i, 'DOT_LEAD');
    createTestEmail({
      body: `Datum: ${date}\nSlot: Vormittag\nTyp: Joker\nBeschreibung: Lasttest ${i}`
    });
  }
  labelTestEmails();
  const processStart = new Date();
  processReservationEmails();
  const processEnd = new Date();
  const durationInSeconds = (processEnd - processStart) / 1000;

  const calendar = (typeof CONFIG !== 'undefined' && CONFIG.CALENDAR_ID) ? 
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();

  if (!calendar) return { name: 'ID 7 – Skalierungstest (Systemstabilität)', passed: false, message: 'Kalender konnte nicht geladen werden.' };

  const maxDateStr = getFutureDate(100, 'DOT_LEAD');
  const parts = maxDateStr.split('.');
  const endDate = new Date(parts[2], parts[1] - 1, parts[0]);
  
  const allEvents = calendar.getEvents(startTime, endDate);
  const loadEventsCount = allEvents.filter(e => e.getDescription().includes('Lasttest')).length;
  const passed = durationInSeconds < 60;
  return {
    name: 'ID 7 – Skalierungstest (Systemstabilität)',
    passed: passed,
    message: passed ?
      `System blieb stabil. ${loadEventsCount} Event(s) eingetragen. Verarbeitungszeit: ${durationInSeconds}s.` : `Fehlgeschlagen. Zeit überschritten (${durationInSeconds}s).`
  };
}


/* ==========================================================================
   URSPRÜNGLICHE HILFSFUNKTIONEN & DEBUG-SKRIPT (UNVERÄNDERT)
   ========================================================================== */

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
    case 'DOT_NO_LEAD': return `${dayNum}.${monthNum}.${year}`;
    case 'EU_SLASH': return `${dayNum}/${monthNum}/${year}`;
    case 'DE_TEXT': return `${dayNum}. ${deMonths[d.getMonth()]} ${year}`;
    case 'DOT_LEAD':
    default: return `${dayLead}.${monthLead}.${year}`;
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
  if (!labelArchiv) { labelArchiv = GmailApp.createLabel(archivLabelName); }
  
  const threads = GmailApp.search('from:me "Reservierung" OR "stornierung" OR "Buchung" OR "Absage"');
  threads.forEach(thread => {
    if(labelNeu) labelNeu.removeFromThread(thread);
    if(labelErledigt) labelErledigt.removeFromThread(thread);
    if(labelAbgelehnt) labelAbgelehnt.removeFromThread(thread);
    
    labelArchiv.addToThreads([thread]);
    thread.moveToArchive();
    thread.markRead();
  });

  try {
    // WECHSEL ZUM STANDARDKALENDER, FALLS CONFIG.CALENDAR_ID LEER IST
    const calendar = (typeof CONFIG !== 'undefined' && CONFIG.CALENDAR_ID) ? 
      CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
      
    // SICHERHEITS-CHECK: Falls der Kalender immer noch null ist
    if (!calendar) {
      Logger.log("❌ KRITISCHER FEHLER bei Kalenderbereinigung: Kalender konnte nicht geladen werden. Bitte CALENDAR_ID überprüfen.");
      return; // Beendet den Block sauber, statt abzustürzen
    }

    const start = new Date(); start.setHours(0,0,0,0);
    const end = new Date();
    end.setDate(start.getDate() + 365);
    
    const events = calendar.getEvents(start, end);
    const myProfile = getAuthorizedUserData(Session.getActiveUser().getEmail());
    const myName = myProfile ? myProfile.name : Session.getActiveUser().getEmail();
    events.forEach(e => {
      if (e.getTitle().includes(myName) || e.getDescription().includes('Lasttest') || e.getDescription().includes('Kurzfrist-Test') || e.getDescription().includes('Wird storniert') || e.getDescription().includes('Format-Test')) {
        e.deleteEvent();
      }
    });
  } catch(e) {
    Logger.log("   -> Hinweis bei Kalenderbereinigung: " + e.message);
  }
  Utilities.sleep(2000);
}

/**
 * Unabhängiges Tool zur händischen Überprüfung der Whitelist.
 * Nutzt DEBUG_EMAIL oder fällt auf die eigene Account-E-Mail zurück, wenn das Feld leer ist.
 */
function debugSpecificWhitelistEmail() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');
  if (!sheetId) {
    Logger.log("❌ FEHLER: Es wurde keine Tabellen-ID ('SHEET_CONFIG_ID') in den Skripteigenschaften gefunden.");
    return;
  }

  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (e) {
    Logger.log(`❌ FEHLER: Die Tabelle mit der ID "${sheetId}" konnte nicht geöffnet werden.`);
    return;
  }
  
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  // Dynamische Ermittlung der Test-E-Mail (Fallback auf eigenen Account)
  const targetEmail = (typeof DEBUG_EMAIL !== 'undefined' && DEBUG_EMAIL.trim() !== "") 
    ? DEBUG_EMAIL.trim().toLowerCase() 
    : Session.getActiveUser().getEmail().trim().toLowerCase();

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  
  Logger.log(`Gesuchte E-Mail-Adresse: "${targetEmail}"`);
  Logger.log(`Untersuchtes Tabellenblatt: "${sheet.getName()}"`); 
  Logger.log("--- Start Tabellen-Scan ---");
  
  let gefunden = false;
  for (let i = 0; i < dataRange.length; i++) {
    const row = dataRange[i];
    if (!row[3]) continue;
    if (row[3].toString().trim().toLowerCase() === targetEmail) {
      Logger.log(`✅ MATCH GEFUNDEN in Zeile ${i + 2}!`);
      const id = row[0] ? row[0].toString().trim() : 'Keine ID';
      const vorname = row[1] ? row[1].toString().trim() : '';
      const nachname = row[2] ? row[2].toString().trim() : '';
      
      let vollerName = `${vorname} ${nachname}`.trim();
      if (!vollerName) vollerName = targetEmail;
      const mobileRaw = row[4] ? row[4].toString().trim() : '';
      const mobile = mobileRaw !== '' ? mobileRaw : 'Nicht hinterlegt';
      Logger.log(`   -> Extrahierte ID:       "${id}"`);
      Logger.log(`   -> Vorname (Rohdaten):   "${vorname}"`);
      Logger.log(`   -> Nachname (Rohdaten):  "${nachname}"`);
      Logger.log(`   -> Generierter Name:     "${vollerName}"`);
      Logger.log(`   -> Mobilnummer:          "${mobile}"`);
      
      gefunden = true;
      break;
    }
  }
  
  if (!gefunden) {
    Logger.log(`❌ FEHLER: Die Adresse "${targetEmail}" wurde in der Whitelist nicht gefunden.`);
  }
}
