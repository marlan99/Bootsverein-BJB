function debugSpecificWhitelistEmail() {
  // ==========================================
  // HIER DIE ZU PRÜFENDE E-MAIL EINTRAGEN:
  const testEmail = "rene.knoblauch@juliusbaer.com"; 
  // ==========================================

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_CONFIG_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_WHITELIST_NAME);
  
  if (!sheet) {
    Logger.log(`❌ FEHLER: Tabellenblatt "${CONFIG.SHEET_WHITELIST_NAME}" wurde nicht gefunden.`);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const searchEmail = testEmail.trim().toLowerCase();
  
  Logger.log(`Gesuchte E-Mail-Adresse: "${searchEmail}"`);
  Logger.log("--- Start Tabellen-Scan ---");
  
  let gefunden = false;
  
  for (let i = 0; i < dataRange.length; i++) {
    const row = dataRange[i];
    const emailInTable = row[3] ? row[3].toString() : '';
    const emailInTableCompare = emailInTable.trim().toLowerCase();
    
    if (emailInTableCompare === searchEmail) {
      Logger.log(`✅ MATCH GEFUNDEN in Zeile ${i + 2}!`);
      Logger.log(`   -> ID: "${row[0]}"`);
      Logger.log(`   -> Vorname: "${row[1]}"`);
      Logger.log(`   -> Nachname: "${row[2]}"`);
      Logger.log(`   -> E-Mail in Zelle: "${emailInTable}" (Länge: ${emailInTable.length} Zeichen)`);
      Logger.log(`   -> Mobile: "${row[4]}"`);
      gefunden = true;
      break;
    }
  }
  
  if (!gefunden) {
    Logger.log(`❌ FEHLER: Die Adresse "${testEmail}" wurde in Spalte D nicht gefunden.`);
    Logger.log("Bitte kontrolliere nochmals, ob die Adressen wirklich exakt in der 4. Spalte (Spalte D) stehen.");
  }
}
