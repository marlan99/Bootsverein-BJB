/**
 * SEPARATES DEBUG-SKRIPT
 * zur Überprüfung einer bestimmten E-Mail-Adresse in der Whitelist.
 * * Anleitung: Füge diesen Code in eine neue, separate .gs-Datei in deinem Projekt ein.
 */

function debugSpecificWhitelistEmail() {
  // ==========================================
  // HIER DIE ZU PRÜFENDE E-MAIL EINTRAGEN:
  const testEmail = "Bootsclub1890@gmail.com"; 
  // ==========================================

  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('SHEET_CONFIG_ID');

  if (!sheetId) {
    Logger.log("❌ FEHLER: Es wurde keine Tabellen-ID ('SHEET_CONFIG_ID') in den Skripteigenschaften gefunden. Wurde das Hauptsystem bereits initialisiert?");
    return;
  }

  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (e) {
    Logger.log(`❌ FEHLER: Die Tabelle mit der ID "${sheetId}" konnte nicht geöffnet werden.`);
    return;
  }
  
  // Holt automatisch das ERSTE Tabellenblatt (Index 0)
  const sheet = ss.getSheets()[0];
  
  if (!sheet) {
    Logger.log(`❌ FEHLER: Kein Tabellenblatt in der Datei gefunden.`);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("❌ FEHLER: Das Tabellenblatt ist leer oder enthält nur die Kopfzeile.");
    return;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const searchEmail = testEmail.trim().toLowerCase();
  
  Logger.log(`Gesuchte E-Mail-Adresse: "${searchEmail}"`);
  Logger.log(`Untersuchtes Tabellenblatt: "${sheet.getName()}"`); // Zeigt im Log den echten Namen des ersten Blatts an
  Logger.log("--- Start Tabellen-Scan ---");
  
  let gefunden = false;
  
  for (let i = 0; i < dataRange.length; i++) {
    const row = dataRange[i];
    
    // Sicherheitsprüfung für komplett leere Zeilen in der Whitelist
    if (!row[3]) continue; 
    
    const emailInTable = row[3].toString();
    const emailInTableCompare = emailInTable.trim().toLowerCase();
    
    if (emailInTableCompare === searchEmail) {
      Logger.log(`✅ MATCH GEFUNDEN in Zeile ${i + 2}!`);
      
      // Dieselbe Logik wie im Hauptskript anwenden
      const id = row[0] ? row[0].toString().trim() : 'Keine ID';
      const vorname = row[1] ? row[1].toString().trim() : '';
      const nachname = row[2] ? row[2].toString().trim() : '';
      
      let vollerName = `${vorname} ${nachname}`.trim();
      if (!vollerName) vollerName = emailInTableCompare;

      const mobileRaw = row[4] ? row[4].toString().trim() : '';
      const mobile = mobileRaw !== '' ? mobileRaw : 'Nicht hinterlegt';

      // Detailreiches Logging für die Fehlersuche
      Logger.log(`   -> Extrahierte ID:       "${id}"`);
      Logger.log(`   -> Vorname (Rohdaten):   "${vorname}" ${vorname === '' ? '(LEER)' : ''}`);
      Logger.log(`   -> Nachname (Rohdaten):  "${nachname}" ${nachname === '' ? '(LEER)' : ''}`);
      Logger.log(`   -> Generierter Name:     "${vollerName}"`);
      Logger.log(`   -> E-Mail in Zelle:      "${emailInTable}" (Länge: ${emailInTable.length} Zeichen)`);
      Logger.log(`   -> Mobilnummer:          "${mobile}" ${mobileRaw === '' ? '(Zelle war leer -> Fallback aktiv)' : ''}`);
      
      gefunden = true;
      break;
    }
  }
  
  if (!gefunden) {
    Logger.log(`❌ FEHLER: Die Adresse "${testEmail}" wurde in Spalte D nicht gefunden.`);
    Logger.log("Bitte kontrolliere nochmals, ob die Adressen wirklich exakt in der 4. Spalte (Spalte D) stehen.");
  }
}
