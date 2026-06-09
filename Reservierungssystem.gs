// =============================================================================
// BC1890 - Integriertes Gesamtskript: Reservierung, Onboarding & Tracking
// Google Apps Script für Google Calendar, Gmail, Drive & Google Spreadsheet
// =============================================================================

// Globale URL-Quelle für die PDF-Anleitung
const PDF_SOURCE_URL = 'https://github.com/marlan99/Bootsverein-BJB/blob/main/Anleitung%20Bootsreservation.pdf'; [cite: 1]

const CONFIG = {
  CALENDAR_ID: '',  // Hier die KALENDER ID eintragen, falls nicht der Standardkalender verwendet wird [cite: 2]
  ADMIN_EMAIL: Session.getActiveUser().getEmail(), [cite: 2]
  GMAIL_LABEL: 'Reservierung/Neu',               [cite: 2]
  SLOT_VORMITTAG: { start: '08:00', end: '14:00' }, [cite: 2]
  SLOT_NACHMITTAG: { start: '14:00', end: '20:00' }, [cite: 2]
  TEST_MODUS_AKTIV: false, [cite: 2]
  EXCEL_SUBJECT: 'Mitgliederliste', [cite: 2]
  EXCEL_TARGET_LABEL: 'Reservierung/Mitgliederliste', [cite: 2]
  TRACKING_TEST_MODUS_AKTIV: false, [cite: 2]
  SPALTE_EMAIL: 4 // für die Kalender-Synchronisierung (Spalte D)
};

// =============================================================================
// 1. KERN-LOGIK: RESERVIERUNGEN & STORNIERUNGEN VERARBEITEN (OPTIMIERT)
// =============================================================================

function processReservationEmails() {
  let labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL) || createGmailLabelStructure(CONFIG.GMAIL_LABEL); [cite: 3]

  // OPTIMIERUNG 1: Kombinierte Suchanfrage spart API-Quota und verhindert Doppelverarbeitung
  const emailThreads = GmailApp.search('in:inbox (subject:"Reservierung" OR subject:"Stornierung")'); [cite: 4]
  Logger.log(`Gefundene relevante Threads im Posteingang: ${emailThreads.length}`); [cite: 5]
  
  // OPTIMIERUNG 2: Kalender-Instanz EINMALIG holen und wiederverwenden
  const calendar = CONFIG.CALENDAR_ID ? [cite: 5]
    CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar(); [cite: 6]
  if (!calendar) {
    Logger.log('❌ KRITISCHER FEHLER: Kalender konnte nicht geladen werden.'); [cite: 6]
    return;
  }

  // Arrays für die Batch-Label-Verarbeitung (Optimierung 1)
  const threadsErledigt = [];
  const threadsAbgelehnt = [];
  const threadsStorniert = [];

  emailThreads.forEach(thread => {
    const messages = thread.getMessages(); [cite: 7]
    messages.forEach(message => {
      if (message.isUnread()) { [cite: 7]
        // Optimierung 4: Sofort als gelesen markieren, um Endlosschleifen bei Timeouts zu verhindern
        message.markRead();
        thread.addLabel(labelNeu); [cite: 7]
        
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

  // Batch-Label-Zuweisung außerhalb der Schleife (Optimierung 1)
  if (threadsErledigt.length > 0) {
    const labelErledigt = GmailApp.getUserLabelByName('Reservierung/Erledigt') || GmailApp.createLabel('Reservierung/Erledigt');
    GmailApp.addLabelsToThreads([labelErledigt], threadsErledigt);
    GmailApp.removeLabelsFromThreads([labelNeu], threadsErledigt);
    GmailApp.moveThreadsToArchive(threadsErledigt);
  }
  
  if (threadsAbgelehnt.length > 0) {
    const labelAbgelehnt = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt');
    GmailApp.addLabelsToThreads([labelAbgelehnt], threadsAbgelehnt);
    GmailApp.removeLabelsFromThreads([labelNeu], threadsAbgelehnt);
    GmailApp.moveThreadsToArchive(threadsAbgelehnt);
  }

  if (threadsStorniert.length > 0) {
    GmailApp.moveThreadsToArchive(threadsStorniert);
  }
}

function processSingleEmail(message, thread, calendar) {
  const sender = message.getFrom().match(/[\w.-]+@[\w.-]+/)?.[0] || 'unbekannt'; [cite: 8]
  const subject = message.getSubject().toLowerCase(); [cite: 8]
  const body = message.getPlainBody(); [cite: 8]
  const data = parseEmailTemplate(body); [cite: 9]
  
  // Optimierung 3: Fehlerhaften Zugriff CONFIG.CONFIG? korrigiert auf CONFIG.GMAIL_LABEL
  const labelNeu = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL); [cite: 9]

  if (!data.valid) { [cite: 12]
    sendRejectionEmail(sender, data.error, thread); [cite: 12]
    return 'ABGELEHNT';
  }

  const userId = sender; [cite: 12]
  // Erleichterte Erkennung von Stornierungen
  if (subject.includes('stornierung') || subject.includes('absage')) { [cite: 13]
    const cancellationSuccess = executeCancellation(data, userId, thread, message); [cite: 13]
    return cancellationSuccess ? 'STORNIERT' : 'ABGELEHNT';
  }

  // Kalender wird hier direkt übergeben
  const validation = validateRequest(data, userId, sender, calendar); [cite: 14]
  if (!validation.valid) { [cite: 14]
    sendRejectionEmail(sender, validation.error, thread); [cite: 15]
    return 'ABGELEHNT';
  }

  // Kalender wird hier direkt übergeben
  const event = createCalendarEvent(data, userId, calendar); [cite: 16]
  if (event) { [cite: 17]
    sendConfirmationEmail(sender, event, data, thread); [cite: 17]
    return 'ERLEDIGT';
  } else {
    sendRejectionEmail(sender, 'Fehler beim Erstellen des Termins im Google Kalender.', thread); [cite: 18]
    return 'ABGELEHNT';
  }
}

function parseEmailTemplate(body) {
  // OPTIMIERUNG 3: Regex-Split fängt Windows-Zeilenumbrüche (\r\n) sauber ab
  const lines = body.split(/\r?\n/).map(l => l.trim()); [cite: 19]
  const data = { valid: false }; [cite: 20]

  const fields = {
    'Datum': 'date',
    'Slot': 'slot',
    'Typ': 'type',
    'Beschreibung': 'description',
    'Anlass': 'occasion'
  }; [cite: 20]
  lines.forEach(line => { [cite: 21]
    for (const [key, prop] of Object.entries(fields)) {
      if (line.toLowerCase().startsWith(key.toLowerCase() + ':')) { // Tolerant gegenüber Groß-/Kleinschreibung beim Key [cite: 21]
        data[prop] = line.substring(key.length + 1).trim(); [cite: 21]
      }
    }
  }); [cite: 21]
  if (!data.date || !data.slot) { [cite: 22]
    data.error = 'Fehlende Pflichtfelder im Text: "Datum:" oder "Slot:" konnten nicht extrahiert werden.'; [cite: 22]
    return data; [cite: 23]
  }

  data.parsedDate = parseEuropeanDate(data.date); [cite: 23]

  if (!data.parsedDate || isNaN(data.parsedDate.getTime())) { [cite: 23]
    data.error = 'Ungültiges Datum. Das Datum konnte nicht erkannt werden (Erlaubt z.B.: 05.06.2026, 5.6.2026, 5/6/2026 oder 5. Juni 2026).'; [cite: 23, 24]
    return data; [cite: 24]
  }

  data.parsedDate.setHours(0, 0, 0, 0); [cite: 25]

  data.slot = data.slot.toLowerCase(); [cite: 25]
  if (!['vormittag', 'nachmittag'].includes(data.slot)) { [cite: 25]
    data.error = 'Der angegebene Slot ist ungültig. Erlaubt ist: "Vormittag" oder "Nachmittag".'; [cite: 25, 26]
    return data; [cite: 26]
  }

  data.type = data.type ? data.type.toLowerCase() : 'standard'; [cite: 26]
  if (!['standard', 'joker'].includes(data.type)) { [cite: 27]
    data.error = 'Der Typ kann nur "Standard" oder "Joker" sein.'; [cite: 27]
    return data; [cite: 28]
  }

  data.valid = true; [cite: 28]
  return data; [cite: 28]
}

// Akzeptiert jetzt die bestehende Kalenderinstanz
function validateRequest(data, userId, sender, calendar) {
  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 28]
  const memberData = getAuthorizedUserData(userId); [cite: 29]
  
  if (!memberData) { [cite: 29]
    return { 
      valid: false, 
      error: `Deine E-Mail-Adresse (${userId}) ist nicht für das Reservierungssystem freigeschaltet. Bitte wende dich an den Vorstand.` [cite: 29, 30]
    };
  }
  
  data.memberId = memberData.id; [cite: 30]
  data.memberMobile = memberData.mobile; [cite: 31]
  if (memberData.name) data.name = memberData.name; [cite: 31]
  
  const today = new Date(); [cite: 31]
  today.setHours(0, 0, 0, 0); [cite: 31]
  // Prüfen des frühestmöglichen Startdatums
  const startDatumRaw = scriptProperties.getProperty('EARLIEST_BOOKING_DATE'); [cite: 32]
  if (startDatumRaw) { [cite: 32]
    const parts = startDatumRaw.split('.'); [cite: 32]
    if (parts.length >= 3) { [cite: 33]
      const startTag = parseInt(parts[0], 10); [cite: 33]
      const startMonat = parseInt(parts[1], 10) - 1; [cite: 34]
      const startJahr = parseInt(parts[2], 10); [cite: 34]
      if (startJahr === today.getFullYear()) { [cite: 35]
        const earliestAllowedDate = new Date(startJahr, startMonat, startTag, 0, 0, 0, 0); [cite: 35]
        if (today < earliestAllowedDate) { [cite: 36]
          const formatiertesStartDatum = `${String(startTag).padStart(2, '0')}.${String(startMonat + 1).padStart(2, '0')}.${startJahr}`; [cite: 36]
          return { 
            valid: false, 
            error: `Das Reservierungssystem ist für das aktuelle Jahr noch nicht freigeschaltet. Buchungen sind erst ab dem ${formatiertesStartDatum} möglich.` [cite: 37, 38]
          };
        }
      }
    }
  }

  if (data.parsedDate < today) { [cite: 39]
    return { valid: false, error: 'Das gewählte Datum liegt in der Vergangenheit.' }; [cite: 39, 40]
  }

  const seasonStart = getCurrentSeasonStart(); [cite: 40]

  // JOKER-VALIDIERUNG
  if (data.type === 'joker') { [cite: 40]
    if (data.parsedDate.getFullYear() !== today.getFullYear()) { [cite: 40]
      return { valid: false, error: `Joker-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.` }; [cite: 40]
    }

    const seasonEnd = new Date(seasonStart); [cite: 41]
    seasonEnd.setFullYear(seasonStart.getFullYear() + 1); [cite: 41]

    const allEvents = calendar.getEvents(seasonStart, seasonEnd); [cite: 41]
    const jokerEvents = allEvents.filter(e => { [cite: 42]
      const desc = e.getDescription() || ''; [cite: 42]
      const title = e.getTitle() || ''; [cite: 42]
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && title.includes('JOKER'); [cite: 42]
    }); [cite: 42]
    if (jokerEvents.length >= 2) { [cite: 43]
      return { valid: false, error: 'Du hast bereits das Maximum von 2 Joker-Terminen in dieser Saison erreicht.' }; [cite: 43, 44]
    }
  }

  // STANDARD-VALIDIERUNG
  if (data.type === 'standard') { [cite: 44]
    if (data.parsedDate.getFullYear() !== today.getFullYear()) { [cite: 44]
      return { valid: false, error: `Standard-Termine sind nur für das aktuelle Kalenderjahr (${today.getFullYear()}) erlaubt.` }; [cite: 44]
    }

    const seasonEnd = new Date(today.getFullYear(), 11, 31, 23, 59, 59); [cite: 45]
    const existingEvents = calendar.getEvents(seasonStart, seasonEnd); [cite: 45]
    const activeStandardEvents = existingEvents.filter(e => { [cite: 46]
      const desc = e.getDescription() || ''; [cite: 46]
      const title = e.getTitle() || ''; [cite: 46]
      return desc.includes(`Mitglieder-ID: ${memberData.id}`) && !title.includes('JOKER') && e.getStartTime() >= today; [cite: 46]
    }); [cite: 46]
    if (activeStandardEvents.length > 0) { [cite: 47]
      const bestehenderTermin = activeStandardEvents[0]; [cite: 47]
      return {
        valid: false,
        error: `Du hast bereits einen aktiven Standard-Termin gebucht (am ${formatDateDDMMYYYY(bestehenderTermin.getStartTime())}). Erst wenn dieser Termin vorbei ist, kannst du einen neuen Standard-Termin vereinbaren.` [cite: 48, 49]
      };
    }
  }

  const slotTime = data.slot === 'vormittag' ? CONFIG.SLOT_VORMITTAG : CONFIG.SLOT_NACHMITTAG; [cite: 50]
  const startTime = new Date(data.parsedDate); [cite: 50]
  const [sh, sm] = slotTime.start.split(':'); [cite: 51]
  startTime.setHours(sh, sm, 0, 0); [cite: 51]

  const endTime = new Date(startTime); [cite: 51]
  const [eh, em] = slotTime.end.split(':'); [cite: 51]
  endTime.setHours(eh, em, 0, 0); [cite: 52]

  const conflicting = calendar.getEvents(startTime, endTime); [cite: 52]
  if (conflicting.length > 0) { [cite: 52]
    return { valid: false, error: 'Dieser Zeitraum (Slot) ist bereits von einem anderen Mitglied belegt.' }; [cite: 52, 53]
  }

  data.startTime = startTime; [cite: 53]
  data.endTime = endTime; [cite: 53]

  return { valid: true }; [cite: 53]
}

// Akzeptiert jetzt die bestehende Kalenderinstanz
function createCalendarEvent(data, userId, calendar) {
  try {
    const myPrefix = 'Boot:'; [cite: 54]
    const title = data.type === 'joker' ? `JOKER - ${myPrefix} ${data.name}` : `${myPrefix} ${data.name}`; [cite: 55]
    const description = [
      `Name: ${data.name}`,
      `Mitglieder-ID: ${data.memberId || 'Nicht hinterlegt'}`, [cite: 56, 57]
      `Kontakt: ${userId}`, [cite: 57]
      `Mobile: ${data.memberMobile || 'Nicht hinterlegt'}`, [cite: 57, 58]
      `Slot: ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}`, [cite: 58]
      `Typ: ${data.type.charAt(0).toUpperCase() + data.type.slice(1)}`, [cite: 58]
      data.description ? `Beschreibung: ${data.description}` : '', [cite: 58, 59]
      data.occasion ? `Anlass: ${data.occasion}` : '', [cite: 59, 60]
      `Eingereicht per E-Mail`
    ].filter(Boolean).join('\n'); [cite: 60]
    const event = calendar.createEvent(title, data.startTime, data.endTime, { description: description }); [cite: 61]
    event.setColor(data.type === 'joker' ? CalendarApp.EventColor.RED : CalendarApp.EventColor.BLUE); [cite: 61]

    return event; [cite: 61]
  } catch (e) {
    Logger.log('Fehler beim Erstellen des Kalendereintrags: ' + e); [cite: 62]
    return null; [cite: 62]
  }
}

// =============================================================================
// 2. EXCEL-IMPORT SYSTEM (EXCEL -> GOOGLE SHEET) - OPTIMIERT
// =============================================================================

function importExcelToSheets() {
  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 63]
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID; [cite: 64]
  const adminEmail = CONFIG.ADMIN_EMAIL; [cite: 64]
  if (!sheetId || !adminEmail) { [cite: 65]
    Logger.log("❌ KRITISCHER FEHLER: Tabellen-ID ('SHEET_CONFIG_ID') oder Admin-E-Mail konnte nicht ermittelt werden."); [cite: 65]
    return;
  }
  
  const searchQuery = `subject:"${CONFIG.EXCEL_SUBJECT}" is:unread`; [cite: 66]
  const threads = GmailApp.search(searchQuery); [cite: 66]
  
  Logger.log(`Prüfe Posteingang auf neue Excel-Listen... Gefunden: ${threads.length}`); [cite: 66]
  const adminEmailLower = adminEmail.toLowerCase(); [cite: 67]
  const targetLabel = GmailApp.getUserLabelByName(CONFIG.EXCEL_TARGET_LABEL) || createGmailLabelStructure(CONFIG.EXCEL_TARGET_LABEL); [cite: 67]
  const errorLabel = GmailApp.getUserLabelByName('Reservierung/Abgelehnt') || GmailApp.createLabel('Reservierung/Abgelehnt'); [cite: 67]
  
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i]; [cite: 68]
    const messages = thread.getMessages(); [cite: 69]
    let importErfolgreich = false;
    
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j]; [cite: 69]
      if (!message.isUnread()) continue; // Nur ungelesene Nachrichten der Konversation betrachten [cite: 70]

      // Optimierung 4: Sofort auf gelesen setzen, um Timeouts abzufangen
      message.markRead();
      const sender = message.getFrom().toLowerCase(); [cite: 70]
      const subject = message.getSubject(); [cite: 71]
      
      if (subject !== CONFIG.EXCEL_SUBJECT) continue; [cite: 71]
      
      // Berechtigungsprüfung via String-Vergleich
      if (!sender.includes(adminEmailLower)) { [cite: 71]
        Logger.log(`WARNUNG: E-Mail von unbefugtem Absender blockiert: ${sender}`); [cite: 71]
        if (errorLabel) thread.addLabel(errorLabel); [cite: 72]
        continue;
      }
      
      const attachments = message.getAttachments(); [cite: 73]
      for (let k = 0; k < attachments.length; k++) {
        const attachment = attachments[k]; [cite: 74]
        const isExcel = attachment.getContentType() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
                        attachment.getName().toLowerCase().endsWith(".xlsx"); [cite: 75]
        
        if (!isExcel) continue; [cite: 75]

        Logger.log(`Verarbeite Excel-Anhang: ${attachment.getName()}`); [cite: 75]
        const fileBlob = attachment.copyBlob(); [cite: 75]
        let tempSheetFile = null; [cite: 76]
        
        try {
          const resource = {
            title: "temp_mitgliederliste_import_" + new Date().getTime(),
            mimeType: MimeType.GOOGLE_SHEETS
          }; [cite: 76]
          // Temporäres Google Sheet aus Excel-Blob erstellen
          tempSheetFile = Drive.Files.create(resource, fileBlob); [cite: 77]
          const tempSpreadsheet = SpreadsheetApp.openById(tempSheetFile.id); [cite: 78]
          const tempSheet = tempSpreadsheet.getSheets()[0]; [cite: 78]
          const tempLastRow = tempSheet.getLastRow(); [cite: 78]
          const tempLastColumn = tempSheet.getLastColumn(); [cite: 78]
          if (tempLastRow <= 1) { [cite: 79]
            Logger.log(`⚠️ Excel-Datei ${attachment.getName()} enthält keine Datenzeilen.`); [cite: 79]
            continue; [cite: 80]
          }
          
          // Daten im Speicher sichern
          const newValues = tempSheet.getRange(2, 1, tempLastRow - 1, tempLastColumn).getValues(); [cite: 80]
          // Zugriff auf Ziel-Tabelle erst JETZT, wenn Daten validiert sind
          const targetSpreadsheet = SpreadsheetApp.openById(sheetId); [cite: 81]
          const targetSheet = targetSpreadsheet.getSheets()[0]; [cite: 82]
          const targetLastRow = targetSheet.getLastRow(); [cite: 82]
          
          // Erst bestehende Daten löschen (Ab Zeile 3)
          if (targetLastRow > 2) { [cite: 82]
            targetSheet.getRange(3, 1, targetLastRow - 2, targetSheet.getLastColumn()).clearContent(); [cite: 82]
          }
          
          // Neue Daten reinschreiben
          targetSheet.getRange(3, 1, newValues.length, tempLastColumn).setValues(newValues); [cite: 83]
          Logger.log(`✅ Mitgliederliste erfolgreich durch Excel-Mail aktualisiert (${newValues.length} Mitglieder).`); [cite: 84]
          
          importErfolgreich = true; [cite: 84]
          break; // Schleife für Anhänge abbrechen, da Import erfolgreich [cite: 84, 85]
          
        } catch (e) {
          Logger.log(`❌ Fehler beim Verarbeiten der Import-Datei: ${e.message}`); [cite: 85]
        } finally {
          // Optimierung 2: Sicheres Löschen gemäß Drive API v3 (Verwendung von DriveApp für Kompatibilität)
          if (tempSheetFile && tempSheetFile.id) {
            try { 
              DriveApp.getFileById(tempSheetFile.id).setTrashed(true);
            } catch(err) {
              Logger.log(`Hinweis beim Aufräumen: Temp-Datei konnte nicht gelöscht werden: ${err.message}`); [cite: 87]
            }
          }
        }
      }
      if (importErfolgreich) break; [cite: 88]
    }
    
    // E-Mail-Status finalisieren
    if (importErfolgreich) { [cite: 90]
      if (targetLabel) thread.addLabel(targetLabel); [cite: 90]
    } else {
      Logger.log(`⚠️ Thread [${thread.getFirstMessageSubject()}] wurde verarbeitet, konnte aber nicht erfolgreich importiert werden.`); [cite: 91]
      if (errorLabel) thread.addLabel(errorLabel); [cite: 92]
    }
  }

  // ─── KETTENREAKTION: TRACKING WIRD BEI JEDEM DURCHLAUF GESTARTET ──────────
  if (typeof tracklistchanges === 'function') { [cite: 92]
    Logger.log("🔎 Starte routinemässige Prüfung auf manuelle Änderungen (tracklistchanges)..."); [cite: 92]
    tracklistchanges(); [cite: 93]
  } else {
    Logger.log("Hinweis: Die Funktion tracklistchanges wurde nicht gefunden."); [cite: 93]
  }
}

// =============================================================================
// 3. MITGLIEDERLISTEN-TRACKING-SYSTEM (DATENÄNDERUNGEN ERKENNEN) - OPTIMIERT
// =============================================================================

function tracklistchanges() {
  Logger.log('=== STARTE MITGLIEDERLISTEN-TRACKING ==='); [cite: 94]
  
  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 94]
  // Alle Properties in EINEM einzigen Netzwerkaufruf holen
  const allProperties = scriptProperties.getProperties(); [cite: 95]
  const sheetId = allProperties['SHEET_CONFIG_ID'] || CONFIG.SHEET_CONFIG_ID; [cite: 95]
  const adminEmail = allProperties['ADMIN_EMAIL'] || CONFIG.ADMIN_EMAIL; [cite: 96]

  if (!sheetId || !adminEmail) { [cite: 96]
    Logger.log('❌ FEHLER: Weder SHEET_CONFIG_ID noch ADMIN_EMAIL konnten gefunden werden.'); [cite: 96]
    return; [cite: 97]
  }

  const lastSnapshotRaw = allProperties['MEMBER_LIST_SNAPSHOT']; [cite: 97]
  const currentSnapshot = {}; [cite: 97]

  try {
    const ss = SpreadsheetApp.openById(sheetId); [cite: 97]
    const sheet = ss.getSheets()[0]; [cite: 98]
    const lastRow = sheet.getLastRow(); [cite: 98]
    
    if (lastRow > 1) { [cite: 98]
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); [cite: 98]
      for (let i = 0; i < data.length; i++) { [cite: 99]
        const id = data[i][0] ? data[i][0].toString().trim() : ''; [cite: 99, 100]
        if (!id) continue; [cite: 100]
        
        currentSnapshot[id] = {
          vorname: data[i][1] ? data[i][1].toString().trim() : '', [cite: 100, 101]
          nachname: data[i][2] ? data[i][2].toString().trim() : '', [cite: 101, 102]
          email: data[i][3] ? data[i][3].toString().trim() : '', [cite: 102, 103]
          mobile: data[i][4] ? data[i][4].toString().trim() : '' [cite: 103, 104]
        };
      }
    }
  } catch (e) {
    Logger.log('❌ Fehler beim Einlesen der Tabelle für Tracking: ' + e.message); [cite: 105]
    return; [cite: 106]
  }

  if (!lastSnapshotRaw) { [cite: 106]
    Logger.log('Kein alter Schnappschuss vorhanden. Erstelle initialen Datenstand...'); [cite: 106]
    scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot)); [cite: 106]
    Logger.log('=== TRACKING BEENDET (Initialer Lauf) ==='); [cite: 107]
    
    if (typeof checkAndWelcomeNewMembers === 'function') { [cite: 107]
      checkAndWelcomeNewMembers(); [cite: 107]
    }
    return; [cite: 108]
  }

  const lastSnapshot = JSON.parse(lastSnapshotRaw); [cite: 108]
  const addedMembers = []; [cite: 108]
  const updatedMembers = []; [cite: 108]

  // OPTIMIERUNG 1: Abgleich und direktes Erkennen von Updates & Neuzugängen
  for (const id in currentSnapshot) { [cite: 109]
    const current = currentSnapshot[id]; [cite: 109]
    const last = lastSnapshot[id]; [cite: 110]
    current.id = id; [cite: 110]

    if (!last) { [cite: 110]
      addedMembers.push(current); [cite: 110]
    } else {
      const changedFields = []; [cite: 111]
      const textDetails = []; [cite: 111]
      // Felder dynamisch prüfen statt 4x hartem "if"
      const fieldsToTrack = { vorname: 'Vorname', nachname: 'Nachname', email: 'E-Mail', mobile: 'Mobil' }; [cite: 112]
      for (const [field, label] of Object.entries(fieldsToTrack)) { [cite: 113]
        if (current[field] !== last[field]) { [cite: 113]
          changedFields.push(field); [cite: 113]
          textDetails.push(`${label}: ${last[field] || '-'} -> ${current[field] || '-'}`); [cite: 114]
        }
      }

      if (changedFields.length > 0) { [cite: 114]
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
      delete lastSnapshot[id]; [cite: 116]
    }
  }

  // Was jetzt noch im alten Snapshot ist, wurde entfernt
  const removedMembers = Object.keys(lastSnapshot).map(id => { [cite: 117]
    const removed = lastSnapshot[id]; [cite: 117]
    removed.id = id; [cite: 117]
    return removed; [cite: 117]
  });

  if (addedMembers.length > 0 || removedMembers.length > 0 || updatedMembers.length > 0) { [cite: 118]
    Logger.log(`Änderungen erkannt! Neu: ${addedMembers.length}, Gelöscht: ${removedMembers.length}, Geändert: ${updatedMembers.length}`); [cite: 118]
    sendChangeReportMail(adminEmail, addedMembers, removedMembers, updatedMembers); [cite: 119]
    
    if (!CONFIG.TRACKING_TEST_MODUS_AKTIV) { [cite: 119]
      scriptProperties.setProperty('MEMBER_LIST_SNAPSHOT', JSON.stringify(currentSnapshot)); [cite: 119]
      Logger.log('Der neue Schnappschuss wurde erfolgreich gespeichert.'); [cite: 119]
    } else {
      Logger.log('⚠️ HINWEIS: Im Tracking-Testmodus wird der alte Schnappschuss NICHT überschrieben.'); [cite: 120]
    }
  } else {
    Logger.log('Keine Änderungen an der Mitgliederliste festgestellt.'); [cite: 121]
  }

  // KETTENREAKTION: Am Ende des Trackings direkt das Onboarding triggern
  if (typeof checkAndWelcomeNewMembers === 'function') { [cite: 122]
    Logger.log("🚀 Starte automatische Prüfung auf neue Mitglieder (checkAndWelcomeNewMembers)..."); [cite: 122]
    checkAndWelcomeNewMembers(); [cite: 123]
  }

  Logger.log('=== TRACKING BEENDET ==='); [cite: 123]
}

function sendChangeReportMail(adminEmail, added, removed, updated) {
  let subject = `✅ Änderungsbericht: Mitgliederliste BC1890`; [cite: 123]
  if (CONFIG.TRACKING_TEST_MODUS_AKTIV) subject = `[TEST] ` + subject; [cite: 124]

  const tableStyle = 'width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 25px; font-size: 14px;'; [cite: 124, 125]
  const thStyle = 'background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 10px; text-align: left; color: #334155; font-weight: bold;'; [cite: 125]
  const tdStyle = 'border: 1px solid #e2e8f0; padding: 10px; vertical-align: top; color: #475569;'; [cite: 126]
  
  // OPTIMIERUNG 3 / 5: Konsequent HTML-Push-Array statt träger String-Verkettung im inneren Loop
  const html = [
    '<div style="font-family: sans-serif; color: #333; max-width: 750px; line-height: 1.5;">', [cite: 127]
    '<h2 style="color: #1a365d; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 20px;">Bericht: Änderungen an der Mitgliederliste</h2>', [cite: 127]
    '<p>Hallo Vorstand,<br>das automatisierte System hat Änderungen in der Mitglieder-Tabelle festgestellt. Nachfolgend findest du alle Details:</p>' [cite: 127, 128]
  ];

  if (added.length > 0) { [cite: 128]
    html.push(`<h3 style="color: #2f855a; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #c6f6d5; padding-bottom: 4px;">➕ Neu hinzugefügte Mitglieder (${added.length})</h3>`, [cite: 128]
              `<table style="${tableStyle}"><tr><th style="${thStyle} width: 10%;">ID</th><th style="${thStyle} width: 25%;">Name</th><th style="${thStyle} width: 40%;">E-Mail</th><th style="${thStyle} width: 25%;">Mobile</th></tr>`); [cite: 128]
    added.forEach(m => { 
      html.push(`<tr><td style="${tdStyle}"><code>${m.id || ''}</code></td><td style="${tdStyle}"><b>${m.vorname} ${m.nachname}</b></td><td style="${tdStyle}">${m.email}</td><td style="${tdStyle}">${m.mobile || '-'}</td></tr>`); [cite: 129]
    });
    html.push('</table>'); [cite: 130]
  }

  if (removed.length > 0) { [cite: 130]
    html.push(`<h3 style="color: #9b2c2c; margin-top: 30px; margin-bottom: 5px; border-bottom: 1px solid #fed7d7; padding-bottom: 4px;">➖ Entfernte Mitglieder (${removed.length})</h3>`, [cite: 130]
              `<table style="${tableStyle}"><tr><th style="${thStyle} width: 10%;">ID</th><th style="${thStyle} width: 25%;">Name</th><th style="${thStyle} width: 40%;">E-Mail</th><th style="${thStyle} width: 25%;">Mobile</th></tr>`); [cite: 130]
    removed.forEach(m => { 
      html.push(`<tr style="background-color: #fafafa;"><td style="${tdStyle} color: #94a3b8;"><code>${m.id || ''}</code></td><td style="${tdStyle} color: #94a3b8;">${m.vorname} ${m.nachname}</td><td style="${tdStyle} color: #94a3b8;">${m.email}</td><td style="${tdStyle} color: #94a3b8;">${m.mobile || '-'}</td></tr>`); [cite: 131]
    });
    html.push('</table>'); [cite: 132]
  }

  if (updated.length > 0) { [cite: 132]
    html.push(`<h3 style="color: #dd6b20; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #feebc8; padding-bottom: 4px;">⚠️ Aktualisierte Mitgliedsdaten (${updated.length})</h3>`); [cite: 132]
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
        const isChanged = m.changedFields.includes(r.key); [cite: 134]
        const cellStyle = isChanged ? 'background-color: #fffaf0; font-weight: bold; color: #c05621;' : ''; [cite: 135]
        html.push(`<tr><td style="${tdStyle} ${cellStyle}">${r.label}</td><td style="${tdStyle} ${cellStyle}">${m.old[r.key] || '-'}</td><td style="${tdStyle} ${cellStyle}">${m.current[r.key] || '-'}</td></tr>`); [cite: 136]
      });

      html.push('</table></div>'); [cite: 136]
    });
  }

  html.push(`<hr style="border: 0; border-top: 1px solid #e2e8f0; margin-top: 40px;"><p style="font-size: 12px; color: #a0aec0;">Generiert am: ${new Date().toLocaleString('de-DE')}</p></div>`); [cite: 137]
  const plainBody = `Änderungsbericht Mitgliederliste BC1890\n\n` + [cite: 138]
    (added.length > 0 ? `Neu (${added.length}):\n` + added.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}`).join('\n') + `\n\n` : '') + [cite: 138]
    (removed.length > 0 ? `Entfernt (${removed.length}):\n` + removed.map(m => `- ID: ${m.id}, Name: ${m.vorname} ${m.nachname}`).join('\n') + `\n\n` : '') + [cite: 138]
    (updated.length > 0 ? `Geändert (${updated.length}):\n` + updated.map(m => `- ID: ${m.id}, Änderungen: ${m.textDetails.join(', ')}`).join('\n') + `\n` : ''); [cite: 138]
  try {
    GmailApp.sendEmail(adminEmail, subject, plainBody, { htmlBody: html.join('') }); [cite: 139]
  } catch (err) {
    Logger.log('❌ Fehler beim Senden des Änderungsberichts: ' + err.message); [cite: 140]
  }
}

// =============================================================================
// 4. ONBOARDING & WILLKOMMENS-SYSTEM (OPTIMIERT)
// =============================================================================

function checkAndWelcomeNewMembers() {
  const modusText = CONFIG.TEST_MODUS_AKTIV ? '⚠️ TESTMODUS (AKTIV)' : '🚀 LIVE-BETRIEB'; [cite: 141, 142]
  Logger.log(`=== STARTE PRÜFUNG AUF NEUE MITGLIEDER [Modus: ${modusText}] ===`); [cite: 142]
  
  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 142]
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID; [cite: 143]
  const adminEmail = CONFIG.ADMIN_EMAIL; [cite: 143]
  if (!sheetId || !adminEmail) { [cite: 144]
    Logger.log('❌ KRITISCHER FEHLER: Tabellen-ID oder Admin-E-Mail konnte nicht ermittelt werden.'); [cite: 144]
    return;
  }

  const welcomedMembersRaw = scriptProperties.getProperty('WELCOMED_MEMBER_IDS'); [cite: 145]
  const welcomedMemberIds = welcomedMembersRaw ? JSON.parse(welcomedMembersRaw) : []; [cite: 145]
  const isInitialRun = welcomedMemberIds.length === 0; [cite: 145]
  if (isInitialRun) { [cite: 145]
    Logger.log('Erster Durchlauf erkannt. Bestehende Mitglieder werden erfasst, ohne E-Mails zu senden.'); [cite: 146]
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId); [cite: 147]
    const sheet = ss.getSheets()[0];  [cite: 147]
    if (!sheet) return;
    const lastRow = sheet.getLastRow(); [cite: 148]
    if (lastRow <= 1) { [cite: 148]
      Logger.log('Keine Mitgliederdaten in der Tabelle gefunden.'); [cite: 148]
      scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([])); [cite: 149]
      return;
    } 
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); [cite: 149]
    // 1. SCHLEIFEN-OPTIMIERUNG: IDs sammeln und Daten validieren in EINEM Durchlauf
    const currentTableIds = new Set(); [cite: 150]
    const validRows = []; [cite: 151]

    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i]; [cite: 151]
      const memberId = row[0] ? row[0].toString().trim() : ''; [cite: 152]
      const email = row[3] ? row[3].toString().trim() : ''; [cite: 152]
      if (memberId) { [cite: 153]
        currentTableIds.add(memberId); [cite: 153]
        if (email) { [cite: 154]
          validRows.push({
            id: memberId,
            vorname: row[1] ? row[1].toString().trim() : '', [cite: 154]
            nachname: row[2] ? row[2].toString().trim() : '', [cite: 154]
            email: email
          });
        }
      }
    }

    // 2. BEREINIGUNG: Schneller Abgleich dank Set.has()
    const cleanedWelcomedIds = welcomedMemberIds.filter(id => currentTableIds.has(id)); [cite: 155]
    const removedCount = welcomedMemberIds.length - cleanedWelcomedIds.length; [cite: 156]
    if (removedCount > 0) { [cite: 156]
      Logger.log(`🧹 BEREINIGUNG: ${removedCount} gelöschte(s) Mitglied(er) aus dem Skript-Gedächtnis entfernt.`); [cite: 156]
    }

    // 3. I/O OPTIMIERUNG: PDF einmalig VOR der Schleife holen (spart massiv API-Aufrufe)
    let attachmentBlob = null; [cite: 157]
    const fileId = scriptProperties.getProperty('PDF_FILE_ID'); [cite: 158]
    if (!fileId) { [cite: 158]
      throw new Error('Keine gültige Google Drive File ID konfiguriert.'); [cite: 158]
    }
    try {
      attachmentBlob = DriveApp.getFileById(fileId).getBlob(); [cite: 159]
    } catch (e) {
      Logger.log(`⚠️ Fehler beim Laden des PDF-Anhangs: ${e.message}. Mails werden ohne Anhang gesendet.`); [cite: 160]
    }

    // 4. VERARBEITUNG: Willkommens-Mails senden
    const welcomedSet = new Set(cleanedWelcomedIds); [cite: 161]
    let mailsSentCount = 0; [cite: 162]

    for (const member of validRows) {
      if (!welcomedSet.has(member.id)) { [cite: 162]
        if (!isInitialRun) { [cite: 162]
          sendWelcomeMail(member.email, member.vorname, member.nachname, adminEmail, attachmentBlob); [cite: 162]
          mailsSentCount++; [cite: 163]
        }
        welcomedSet.add(member.id); [cite: 163]
      }
    }

    // Zurück in Array konvertieren für Speicherung
    scriptProperties.setProperty('WELCOMED_MEMBER_IDS', JSON.stringify([...welcomedSet])); [cite: 164]
    Logger.log(`Prüfung abgeschlossen. ${mailsSentCount} neue(s) Mitglied(er) verarbeitet.`); [cite: 165]
    
  } catch (e) {
    Logger.log('Fehler im Onboarding-Script: ' + e.message); [cite: 165]
  }
}

// Erwartet jetzt den fertigen Blob, um DriveApp-Aufrufe in der Schleife zu verhindern
function sendWelcomeMail(toEmail, vorname, nachname, adminEmail, attachmentBlob) {
  const name = vorname || 'Mitglied'; [cite: 166, 167]
  
  let finalReceiver = toEmail; [cite: 167]
  let finalCc = adminEmail;  [cite: 167]
  let subject = 'Herzlich willkommen beim Bootsclub 1890! ⛵'; [cite: 167]
  let testNoticeHtml = ''; [cite: 168]
  let testNoticePlain = ''; [cite: 168]

  if (CONFIG.TEST_MODUS_AKTIV) {
    finalReceiver = adminEmail;  [cite: 168]
    finalCc = ''; [cite: 168]
    subject = `[TEST-MODUS für: ${toEmail}] Herzlich Willkommen beim Bootsclub 1890! ⛵`; [cite: 169]
    testNoticeHtml = `
      <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 12px; margin-bottom: 20px; color: #856404; font-family: sans-serif; border-radius: 4px;">
        ⚠️ <b>SYSTEM-HINWEIS (TEST-MODUS):</b> Diese E-Mail wurde automatisch abgefangen und an den Vorstand umgeleitet.<br>
        <b>Geplanter Empfänger im Live-Betrieb:</b> ${vorname} ${nachname} (&lt;${toEmail}&gt;)
      </div>
    `; [cite: 170]
    testNoticePlain = `[⚠️ TEST-MODUS - Geplanter Empfänger im Live-Betrieb: ${vorname} ${nachname} (${toEmail})]\n\n`; [cite: 171]
  }
  
  const htmlBody = `
    ${testNoticeHtml}
    Hallo ${name},<br><br>
    Herzlich Willkommen im <b>Bootsclub 1890</b>!<br><br>
    Deine E-Mail-Adresse wurde erfolgreich für unser automatisiertes Reservierungssystem freigeschaltet.<br><br> [cite: 173]
    Ab sofort kannst du Bootstermine direkt per E-Mail reservieren. [cite: 173]
    <br>
    <b>Im Anhang dieser E-Mail findest du die detaillierte Anleitung als PDF-Datei.</b><br><br> [cite: 174]
    Hier sind die wichtigsten Kernpunkte im Überblick:<br>
    • Sende Reservierungen an: <b>${adminEmail}</b>. [cite: 174]
    Die E-Mail muss das Wort <b>Reservierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text enthalten.<br><br> [cite: 175]
    • Für eine Stornierung sende einfach das Wort <b>Stornierung</b> im Betreff und die Zeilen <b>Datum:</b> und <b>Slot:</b> (Vormittag/Nachmittag) als Text (bis max. 24 Stunden vor dem Termin).<br><br> [cite: 175]
    Bitte lies dir die angehängte PDF-Anleitung aufmerksam durch, bevor du deine erste Reservierung vornimmst.<br><br> [cite: 175]
    Bei Fragen steht dir der Vorstand jederzeit gerne zur Verfügung.<br><br> [cite: 175]
    Allzeit gute Fahrt und viel Spass auf dem Wasser!<br><br> [cite: 175]
    <b>Dein Vorstand</b><br> [cite: 175]
  `;
  const plainBody = `${testNoticePlain}Hallo ${name},\n\nherzlich willkommen beim Bootsclub 1890!\nDeine E-Mail wurde für das Reservierungssystem freigeschaltet.\n\nEine detaillierte Anleitung findest du im Anhang dieser E-Mail als PDF.\n\nBitte sende Reservierungen an ${adminEmail}.\n\nAllzeit gute Fahrt!\nDein Vorstand`; [cite: 176]
  try {
    const options = {
      cc: finalCc, 
      replyTo: adminEmail,
      htmlBody: htmlBody
    }; [cite: 177]
    if (attachmentBlob) { [cite: 178]
      options.attachments = [attachmentBlob]; [cite: 178]
    }

    GmailApp.sendEmail(finalReceiver, subject, plainBody, options); [cite: 178]
  } catch (error) {
    Logger.log(`❌ FEHLER beim Senden der Willkommens-Mail: ${error.message}`); [cite: 179]
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
    // Holt die gesamte Kopfzeile (Zeile 1), um die Spaltenindizes zu ermitteln
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Sucht die Position der E-Mail-Spalte anhand der Beschriftung
    const emailColIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'e-mail' || h.toString().trim().toLowerCase() === 'email');
    // Optional: Sucht auch die Status-Spalte, um inaktive Nutzer auszuschließen
    const statusColIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'status');

    if (emailColIndex === -1) {
      Logger.log('❌ FEHLER: Spalte "E-Mail" oder "Email" konnte in der Kopfzeile (Zeile 1) nicht gefunden werden.');
      return;
    }
    Logger.log(`ℹ️ E-Mail-Spalte dynamisch auf Index ${emailColIndex} (Spalte ${String.fromCharCode(65 + emailColIndex)}) gefunden.`);
    // -------------------------------

    // Holt alle Daten ab Zeile 2
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const sheetEmailsSet = new Set();

    // 1. SCHRITT: Alle gültigen E-Mails aus dem Sheet sammeln
    for (let i = 0; i < dataRange.length; i++) {
      const email = dataRange[i][emailColIndex] ? dataRange[i][emailColIndex].toString().trim().toLowerCase() : '';
      const status = statusColIndex !== -1 && dataRange[i][statusColIndex] ? dataRange[i][statusColIndex].toString().trim().toLowerCase() : '';
      
      // Nur Adressen hinzufügen, die ein '@' enthalten und NICHT den Status 'inaktiv' besitzen
      if (email && email.includes('@') && status !== 'inaktiv') {
        sheetEmailsSet.add(email);
      }
    }

    Logger.log(`📋 Anzahl gültiger (aktiver) Mitglieder aus der Tabelle: ${sheetEmailsSet.size}`);

    // 2. SCHRITT: Aktuelle Kalender-Berechtigungen auslesen
    const calendar = CalendarApp.getCalendarById(kalenderId);
    if (!calendar) {
      Logger.log(`❌ FEHLER: Kalender mit der ID '${kalenderId}' wurde nicht gefunden.`);
      return;
    }

    const aclList = calendar.getUsersWithAccess();
    const currentAclEmails = aclList.map(user => user.toString().trim().toLowerCase());
    
    Logger.log(`📅 Anzahl Personen mit Kalender-Zugriff aktuell: ${currentAclEmails.length}`);

    // 3. SCHRITT: KALENDER-ABGLEICH (Hinzufügen & Entfernen)
    
    // A) Neue Mitglieder hinzufügen, die noch keinen Zugriff haben
    sheetEmailsSet.forEach(email => {
      // Den Admin selbst (Besitzer) nicht nochmals hinzufügen
      if (email === CONFIG.ADMIN_EMAIL.toLowerCase()) return;

      if (!currentAclEmails.includes(email)) {
        try {
          calendar.addEditor(email);
          Logger.log(`➕ Zugriff ERLAUBT für neues Mitglied: ${email}`);
        } catch (e) {
          Logger.log(`⚠️ Fehler beim Hinzufügen von ${email}: ${e.message}`);
        }
      }
    });

    // B) Alte/Inaktive Mitglieder entfernen, die nicht mehr in der erlaubten Liste stehen
    currentAclEmails.forEach(email => {
      // Den Admin/Besitzer niemals aus dem eigenen Kalender entfernen!
      if (email === CONFIG.ADMIN_EMAIL.toLowerCase()) return;

      if (!sheetEmailsSet.has(email)) {
        try {
          calendar.removeUser(email);
          Logger.log(`➖ Zugriff ENTZOGEN für ausgeschiedenes/inaktives Mitglied: ${email}`);
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
  const memberData = getAuthorizedUserData(userId); [cite: 201]
  
  if (!memberData) { [cite: 201]
    GmailApp.sendEmail(userId, 'Löschen der Buchung abgelehnt', `Deine E-Mail-Adresse (${userId}) ist nicht im System hinterlegt.`, { replyTo: CONFIG.ADMIN_EMAIL }); [cite: 201]
    return false;
  }
  
  data.name = memberData.name; [cite: 202]
  const jetzt = new Date();  [cite: 203]
  const slotTime = data.slot === 'vormittag' ? CONFIG.SLOT_VORMITTAG : CONFIG.SLOT_NACHMITTAG; [cite: 203]
  const terminStartZeit = new Date(data.parsedDate); [cite: 203]
  const [sh, sm] = slotTime.start.split(':'); [cite: 204]
  terminStartZeit.setHours(sh, sm, 0, 0);  [cite: 204]

  // 24 Stunden Frist berechnen
  const stornierungsFrist = new Date(terminStartZeit.getTime() - (24 * 60 * 60 * 1000)); [cite: 204]
  if (jetzt > stornierungsFrist) { [cite: 205]
    let fehlerGrund = terminStartZeit < jetzt ? 'Der Termin liegt in der Vergangenheit.' : `Die Frist für eine automatische Stornierung (24h vor Beginn) ist abgelaufen.`; [cite: 205, 206]
    GmailApp.sendEmail(userId, 'Löschen der Buchung abgelehnt', `Hallo ${data.name},\n\n❌ Grund: ${fehlerGrund}`, { replyTo: CONFIG.ADMIN_EMAIL }); [cite: 207]
    return false; 
  }

  const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar(); [cite: 208]
  const terminEndZeit = new Date(terminStartZeit); [cite: 209]
  const [eh, em] = slotTime.end.split(':'); [cite: 209]
  terminEndZeit.setHours(eh, em, 0, 0); [cite: 209]

  const events = calendar.getEvents(terminStartZeit, terminEndZeit); [cite: 209]
  const userEvent = events.find(e => (e.getDescription() || '').includes(`Mitglieder-ID: ${memberData.id}`)); [cite: 210]

  if (userEvent) { [cite: 210]
    if (userEvent.getTitle().toUpperCase().includes('JOKER')) { [cite: 210]
      GmailApp.sendEmail(userId, 'Stornierung fehlgeschlagen', `❌ Joker-Termine können nicht automatisch storniert werden. Bitte wende dich an den Admin.`, { replyTo: CONFIG.ADMIN_EMAIL }); [cite: 210]
      return false;
    }

    userEvent.deleteEvent();  [cite: 211]
    // OPTIMIERUNG 1: Tippfehler im Betreff korrigiert ("Bestätigung" statt "BestBTigung")
    GmailApp.sendEmail(userId, 'Bestätigung: Termin freigegeben', `Deine Reservierung für den ${formatDateDDMMYYYY(data.parsedDate)} wurde erfolgreich storniert.`, { replyTo: CONFIG.ADMIN_EMAIL }); [cite: 211]
    return true;
  } else {
    GmailApp.sendEmail(userId, 'Stornierung fehlgeschlagen', `❌ Es wurde kein passender aktiver Termin für dich an diesem Tag gefunden.`, { replyTo: CONFIG.ADMIN_EMAIL }); [cite: 213]
    return false;
  }
}

function sendConfirmationEmail(to, event, data, thread) {
  const subject = 'Buchung bestätigt: ' + event.getTitle(); [cite: 214]
  const htmlBody = `Hallo ${data.name},<br><br>dein Termin wurde erfolgreich eingetragen:<br><br>📅 <b>Datum:</b> ${formatDateDDMMYYYY(data.parsedDate)}<br>🕒 <b>Slot:</b> ${data.slot.charAt(0).toUpperCase() + data.slot.slice(1)}<br><br>Dein Vorstand`; [cite: 215]
  const plainBody = `Hallo ${data.name},\n\ndein Termin wurde erfolgreich eingetragen.`; [cite: 216]

  try {
    GmailApp.sendEmail(to, subject, plainBody, { replyTo: CONFIG.ADMIN_EMAIL, htmlBody: htmlBody }); [cite: 216]
  } catch (error) {
    if (thread) thread.createDraftReply(plainBody, { htmlBody: htmlBody }); [cite: 217]
  }
}

function sendRejectionEmail(to, reason, thread) {
  const subject = 'Buchung abgelehnt'; [cite: 218]
  const body = `Hallo,\n\nleider konnte deine Reservierung nicht angenommen werden:\n\n❌ Grund: ${reason}`; [cite: 219]
  try {
    GmailApp.sendEmail(to, subject, body, { replyTo: CONFIG.ADMIN_EMAIL }); [cite: 220]
  } catch (error) {
    if (thread) thread.createDraftReply(`Ablehnungsgrund: ${reason}`); [cite: 221]
  }
}

// Globaler Cache zur Vermeidung mehrfacher Tabellen-I/O-Aufrufe während desselben Skript-Laufs
let memberDataCache_ = null; [cite: 222]

function getAuthorizedUserData(email) {
  const searchEmail = email.trim().toLowerCase(); [cite: 223]
  
  // OPTIMIERUNG 2: Cache-Abfrage spart wertvolle Millisekunden bei Schleifendurchläufen
  if (memberDataCache_ && memberDataCache_[searchEmail]) { [cite: 223]
    return memberDataCache_[searchEmail]; [cite: 224]
  }

  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 224]
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID; [cite: 224]
  let ss, sheet; [cite: 224]
  try { if (sheetId) { ss = SpreadsheetApp.openById(sheetId); sheet = ss.getSheets()[0]; } } catch (e) {} [cite: 225, 226]

  // Automatisches Erstellen der Tabelle falls gelöscht oder nicht vorhanden
  if (!ss || !sheet || sheet.getLastRow() === 0) { [cite: 226]
    try {
      const folderName = "Google Kalender Reservierungssystem"; [cite: 226]
      let targetFolder = DriveApp.getFoldersByName(folderName).hasNext() ? DriveApp.getFoldersByName(folderName).next() : DriveApp.createFolder(folderName); [cite: 227]

      ss = SpreadsheetApp.create('Mitgliederliste'); [cite: 227]
      sheet = ss.getSheets()[0]; [cite: 227]
      sheetId = ss.getId(); [cite: 227]
      const file = DriveApp.getFileById(sheetId); [cite: 228]
      targetFolder.addFile(file); [cite: 228]
      DriveApp.getRootFolder().removeFile(file); [cite: 228]

      sheet.appendRow(["Mitglieder ID", "Vorname", "Name", "E-Mail", "Mobile"]); [cite: 228]
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold"); [cite: 228]
      sheet.appendRow(["BJB-001", "Vorstand", "Boot", CONFIG.ADMIN_EMAIL, "Nicht hinterlegt"]); [cite: 229]
      sheet.autoResizeColumns(1, 5); [cite: 229]

      scriptProperties.setProperty('SHEET_CONFIG_ID', sheetId); [cite: 229]
    } catch (err) { return null; [cite: 229]
    }
  }

  try {
    const lastRow = sheet.getLastRow(); [cite: 230]
    if (lastRow <= 1) return null; [cite: 230]
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); [cite: 231]
    
    // Cache initialisieren
    memberDataCache_ = {}; [cite: 231]
    let foundUserData = null; [cite: 232]

    for (let i = 0; i < dataRange.length; i++) {
      const currentEmail = dataRange[i][3] ? dataRange[i][3].toString().trim().toLowerCase() : ''; [cite: 232, 233]
      if (!currentEmail) continue; [cite: 233]

      const userObj = {
        id: dataRange[i][0] ? dataRange[i][0].toString().trim() : 'Keine ID',  [cite: 233, 234]
        name: `${dataRange[i][1] || ''} ${dataRange[i][2] || ''}`.trim() || currentEmail,    [cite: 234, 235]
        mobile: dataRange[i][4] ? dataRange[i][4].toString().trim() : 'Nicht hinterlegt' [cite: 235, 236]
      };
      // Alle Mitglieder in den Cache schreiben für zukünftige Suchen im selben Lauf
      memberDataCache_[currentEmail] = userObj; [cite: 237]
      if (currentEmail === searchEmail) { [cite: 238]
        foundUserData = userObj; [cite: 239]
      }
    }
    return foundUserData; [cite: 239]
  } catch (e) { return null; [cite: 240]
  }
}

function getCurrentSeasonStart() {
  return new Date(new Date().getFullYear(), 0, 1); [cite: 240]
}

function createGmailLabelStructure(fullLabelPath) {
  const parts = fullLabelPath.split('/'); [cite: 240]
  let currentPath = ''; [cite: 241]
  let finalLabel = null; [cite: 241]
  for (let i = 0; i < parts.length; i++) {
    currentPath = currentPath ? currentPath + '/' + parts[i] : parts[i]; [cite: 241, 242]
    let label = GmailApp.getUserLabelByName(currentPath) || GmailApp.createLabel(currentPath); [cite: 242]
    if (i === parts.length - 1) finalLabel = label; [cite: 243]
  }
  return finalLabel; [cite: 244]
}

// =============================================================================
// 6. ZEITGESTEUERTE AUTOMATISIERUNGEN & ZENTRALES SETUP
// =============================================================================

function sendDailyReservationReminders() {
  Logger.log("=== STARTE TÄGLICHE ERINNERUNGS-PRÜFUNG ==="); [cite: 244]
  const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar(); [cite: 245]
    
  if (!calendar) { [cite: 245]
    Logger.log("❌ KRITISCHER FEHLER: Kalender konnte nicht geladen werden."); [cite: 245]
    return;  [cite: 246]
  }

  // OPTIMIERUNG 3: Exakte "Morgen"-Zeitspanne berechnen
  const tomorrowStart = new Date(); [cite: 246]
  tomorrowStart.setDate(tomorrowStart.getDate() + 1); [cite: 246]
  tomorrowStart.setHours(0, 0, 0, 0); [cite: 247]
  
  const tomorrowEnd = new Date(tomorrowStart); [cite: 247]
  tomorrowEnd.setHours(23, 59, 59, 999); [cite: 247]
  
  const events = calendar.getEvents(tomorrowStart, tomorrowEnd); [cite: 247]
  events.forEach(event => {
    const desc = event.getDescription() || ""; [cite: 248]
    const emailMatch = desc.match(/Kontakt:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/); [cite: 248]
    if (emailMatch && emailMatch[1]) { [cite: 248]
      const slotName = event.getStartTime().getHours() === 8 ? "Vormittag (08:00 - 14:00)" : "Nachmittag (14:00 - 20:00)"; [cite: 248]
      let body = `Hallo!\n\nAutomatische Erinnerung für deine Reservierung morgen:\n📅 Datum: ${formatDateDDMMYYYY(tomorrowStart)}\n⏱️ Slot: ${slotName}\n\nViel Spass mit dem Boot!`; [cite: 248]
      GmailApp.sendEmail(emailMatch[1].trim(), `Erinnerung: Deine Boot Buchung für morgen!`, body); [cite: 248]
    }
  });
}

function ensureInitialSheet() {
  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 249]
  let sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID; [cite: 249]
  if (!sheetId) { [cite: 250]
    Logger.log("📂 Initialisiere Google Sheet und Ordnerstruktur für den Erststart..."); [cite: 250]
    getAuthorizedUserData(CONFIG.ADMIN_EMAIL); [cite: 250]
    Logger.log("✅ Google Sheet wurde erfolgreich im Google Drive angelegt."); [cite: 251]
  } else {
    Logger.log("ℹ️ Google Sheet existiert bereits. ID: " + sheetId); [cite: 251]
  }
}

function fetchAndSyncAnleitungPDF() {
  Logger.log("🔄 Synchronisiere PDF-Anleitung von GitHub..."); [cite: 252]
  const scriptProperties = PropertiesService.getScriptProperties(); [cite: 252]
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID') || CONFIG.SHEET_CONFIG_ID; [cite: 252]
  if (!sheetId) return; [cite: 253]
  
  const sheetFile = DriveApp.getFileById(sheetId); [cite: 253]
  const parents = sheetFile.getParents(); [cite: 253]
  const targetFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder(); [cite: 253]
  let targetUrl = PDF_SOURCE_URL; [cite: 254]
  
  // OPTIMIERUNG 4: Korrekte Übersetzung in die GitHub-RAW-Domain zur Vermeidung von PDF-Korruption
  if (targetUrl.includes('github.com')) { [cite: 254]
    targetUrl = targetUrl [cite: 254]
      .replace('github.com', 'raw.githubusercontent.com') [cite: 255]
      .replace('/blob/', '/'); [cite: 255]
  }
  
  Logger.log("📥 Rufe URL ab: " + targetUrl); [cite: 255]
  const response = UrlFetchApp.fetch(targetUrl, { muteHttpExceptions: true }); [cite: 255]
  if (response.getResponseCode() !== 200) { [cite: 256]
    Logger.log("❌ Fehler beim Abrufen der PDF von GitHub: " + response.getResponseCode()); [cite: 256]
    return; [cite: 257]
  }
  
  const pdfBlob = response.getBlob().setContentType("application/pdf").setName("Anleitung Bootsreservation.pdf"); [cite: 257]
  const fileName = "Anleitung Bootsreservation.pdf"; [cite: 257]
  const files = targetFolder.getFilesByName(fileName); [cite: 257]
  let localFile = files.hasNext() ? files.next() : null; [cite: 258]
  
  if (localFile) { [cite: 258]
    const headers = response.getHeaders(); [cite: 258]
    const remoteLastModifiedStr = headers["Last-Modified"] || headers["last-modified"]; [cite: 259]
    let shouldUpdate = true; [cite: 259]
    if (remoteLastModifiedStr) { [cite: 260]
      const remoteDate = new Date(remoteLastModifiedStr); [cite: 260]
      const localDate = new Date(localFile.getLastUpdated()); [cite: 260]
      if (remoteDate <= localDate) { [cite: 261]
        shouldUpdate = false; [cite: 261]
        Logger.log("ℹ️ Lokale PDF ist auf dem neuesten Stand."); [cite: 262]
      }
    }
    
    if (shouldUpdate) { [cite: 262]
      Logger.log("🔄 Lokale PDF veraltet. Ersetze Datei..."); [cite: 262]
      localFile.setTrashed(true); [cite: 263]
      const newFile = targetFolder.createFile(pdfBlob); [cite: 263]
      scriptProperties.setProperty('PDF_FILE_ID', newFile.getId()); [cite: 263]
    } else {
      scriptProperties.setProperty('PDF_FILE_ID', localFile.getId()); [cite: 263]
    }
    
  } else {
    Logger.log("📥 PDF existiert lokal nicht. Erstelle neue Datei..."); [cite: 264]
    const newFile = targetFolder.createFile(pdfBlob); [cite: 265]
    scriptProperties.setProperty('PDF_FILE_ID', newFile.getId()); [cite: 265]
  }
}

function setupTriggers() {
  Logger.log('========================================================================'); [cite: 265]
  Logger.log('🚀 STARTE CENTRAL SYSTEM SETUP...'); [cite: 265]
  Logger.log('========================================================================'); [cite: 265]

  ensureInitialSheet(); [cite: 265]
  try {
    fetchAndSyncAnleitungPDF(); [cite: 266]
  } catch(pdfError) {
    Logger.log("⚠️ Warnung beim PDF-Sync: " + pdfError.toString()); [cite: 266]
  }

  const existingTriggers = ScriptApp.getProjectTriggers(); [cite: 267]
  existingTriggers.forEach(t => ScriptApp.deleteTrigger(t)); [cite: 267]

  // Trigger-Definitionen
  ScriptApp.newTrigger('processReservationEmails').timeBased().everyMinutes(1).create(); [cite: 267]
  ScriptApp.newTrigger('sendDailyReservationReminders').timeBased().everyDays(1).atHour(4).create(); [cite: 267]
  ScriptApp.newTrigger('importExcelToSheets').timeBased().everyMinutes(10).create(); [cite: 267]
  
  ['Reservierung/Neu', 'Reservierung/Erledigt', 'Reservierung/Abgelehnt', CONFIG.EXCEL_TARGET_LABEL].forEach(label => { [cite: 268]
    if (!GmailApp.getUserLabelByName(label)) createGmailLabelStructure(label); [cite: 268]
  });
  
  Logger.log('========================================================================'); [cite: 268]
  Logger.log('🎉 INTEGRIERTES GESAMT-SETUP ERFOLGREICH!'); [cite: 268]
  Logger.log('========================================================================'); [cite: 268]
}

// =============================================================================
// 7. ENTWICKLER-WERKZEUGE (MAINTENANCE)
// =============================================================================

function setEarliestBookingDate() {
  const zielDatum = '01.04.' + new Date().getFullYear();  [cite: 269]
  PropertiesService.getScriptProperties().setProperty('EARLIEST_BOOKING_DATE', zielDatum); [cite: 269]
  Logger.log(`Frühestmögliches Startdatum wurde erfolgreich auf den ${zielDatum} gesetzt!`); [cite: 270]
}

function resetWelcomeDatabase() {
  PropertiesService.getScriptProperties().deleteProperty('WELCOMED_MEMBER_IDS'); [cite: 270]
  Logger.log('Onboarding-Datenbank zurückgesetzt.'); [cite: 270]
}

function resetTrackingSnapshot() {
  PropertiesService.getScriptProperties().deleteProperty('MEMBER_LIST_SNAPSHOT'); [cite: 270]
  Logger.log('Tracking-Schnappschuss wurde erfolgreich gelöscht.'); [cite: 271]
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
