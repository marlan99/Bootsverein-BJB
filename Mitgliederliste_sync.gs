function syncWeblingWithGoogleSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_CONFIG_ID'); // Deine Google Sheet ID
  const subdomain = "bootsclub1890";    // Deine Webling-Subdomain

  const scriptProperties = PropertiesService.getScriptProperties();
  let apiKey = scriptProperties.getProperty('WEBLING_API_KEY');

  // Falls die Property noch nicht existiert, mit Standardwert anlegen
  if (apiKey === null) {
    scriptProperties.setProperty('WEBLING_API_KEY', 'undefiniert');
    apiKey = 'undefiniert';
    Logger.log("WEBLING_API_KEY war nicht vorhanden und wurde mit dem Standardwert 'undefiniert' angelegt.");
  }

  // Wenn der Wert (noch) der Platzhalter ist, macht ein API-Aufruf keinen Sinn
  if (apiKey === "undefiniert") {
    Logger.log("Fehler: WEBLING_API_KEY ist noch nicht konfiguriert (Wert = 'undefiniert'). Bitte in den Skripteigenschaften einen echten API-Key hinterlegen.");
    return;
  }

  const options = {
    "method": "get",
    "headers": {
      "apikey": apiKey,
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  };

  // 1. Webling-IDs abrufen (die JSON-Liste mit "objects")
  const listUrl = `https://${subdomain}.webling.ch/api/1/member`;
  const listResponse = UrlFetchApp.fetch(listUrl, options);
  
  if (listResponse.getResponseCode() !== 200) {
    Logger.log("Fehler beim Abrufen der Webling-Mitgliederliste: " + listResponse.getContentText());
    return;
  }
  
  const weblingIds = JSON.parse(listResponse.getContentText()).objects.map(Number);

  // 2. Google Sheet öffnen und das ERSTE Tabellenblatt nehmen
  const sheets = SpreadsheetApp.openById(sheetId).getSheets();
  if (sheets.length === 0) {
    Logger.log("Fehler: Das Google Sheet enthält keine Tabellenblätter!");
    return;
  }
  const sheet = sheets[0]; // [0] greift auf das erste Tabellenblatt von links zu

  // 2b. Sheet-Daten einlesen (inkl. Header in Zeile 1)
  const sheetData = sheet.getDataRange().getValues();

  // sheetIds: Spalte A ab Zeile 2 (Header überspringen).
  // Nicht-numerische Werte (z.B. Platzhalter wie "BJB-000") werden zu null,
  // damit sie beim Abgleich ignoriert werden.
  const sheetIds = sheetData.slice(1).map(row => {
    const num = Number(row[0]);
    return (row[0] === "" || isNaN(num)) ? null : num;
  });

  // 3. IDs abgleichen
  // Neue Mitglieder: ID ist in Webling, aber nicht im Sheet (und wir ignorieren dort "null"-Werte von "BJB-000")
  const newMembers = weblingIds.filter(id => !sheetIds.includes(id));
  
  // Gelöschte/Inaktive Mitglieder: Numerische ID ist im Sheet, aber nicht mehr in Webling
  const removedMembers = sheetIds.filter(id => id !== null && !weblingIds.includes(id));

  Logger.log(`Gefundene neue Mitglieder in Webling: ${newMembers.length}`);
  Logger.log(`In Webling gelöschte Mitglieder: ${removedMembers.length}`);

  // 4. NEUE MITGLIEDER AUS WEBLING INS SHEET EINTRAGEN
  newMembers.forEach(id => {
    const memberUrl = `https://${subdomain}.webling.ch/api/1/member/${id}`;
    const memberResponse = UrlFetchApp.fetch(memberUrl, options);
    
    if (memberResponse.getResponseCode() === 200) {
      const member = JSON.parse(memberResponse.getContentText());
      const props = member.properties || {};
      
      const newRow = [
        id,                                // Spalte A: Mitglieder ID
        props["Vorname"] || "",            // Spalte B: Vorname
        props["Name"] || "",               // Spalte C: Name
        props["E-Mail"] || "",             // Spalte D: E-Mail
        props["Mobile"] || props["Telefon"] || "", // Spalte E: Mobile
        ""                                 // Spalte F: Webformular (bleibt vorerst leer für manuelle Zuordnung)
      ];
      
      sheet.appendRow(newRow);
      Logger.log(`Hinzugefügt: ID ${id} - ${newRow[1]} ${newRow[2]}`);
    }
    Utilities.sleep(100); // Kurze Pause wegen Rate Limit
  });

  // 5. GELÖSCHTE MITGLIEDER IM SHEET MARKIEREN (OPTIONAL)
  if (removedMembers.length > 0) {
    sheetData.forEach((row, index) => {
      const currentId = Number(row[0]);
      if (removedMembers.includes(currentId)) {
        const rowNum = index + 2; // +2 weil Zeile 1 der Header ist und das Array bei 0 startet
        
        sheet.getRange(rowNum, 1, 1, 6).setBackground("#fce8e6");
        Logger.log(`Als inaktiv markiert (Rot hinterlegt): ID ${currentId} - ${row[1]} ${row[2]}`);
      }
    });
  }

  Logger.log("Synchronisation abgeschlossen.");
}
