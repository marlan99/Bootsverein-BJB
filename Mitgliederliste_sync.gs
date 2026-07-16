function syncWeblingWithGoogleSheet() {
  const sheetId = "SHEET_CONFIG_ID"; // Deine Google Sheet ID
  const tabName = "Mitglieder";      // Name deines Tabellenblatts
  const subdomain = "deinverein";    // Deine Webling-Subdomain
  const apiKey = PropertiesService.getScriptProperties().getProperty('WEBLING_API_KEY');
  
  if (!apiKey) {
    Logger.log("Fehler: Kein API-Key in den Skripteigenschaften gefunden!");
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

  // 2. Google Sheet öffnen und bestehende IDs auslesen
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(tabName);
  if (!sheet) {
    Logger.log(`Fehler: Tabellenblatt "${tabName}" wurde nicht gefunden!`);
    return;
  }

  const lastRow = sheet.getLastRow();
  let sheetData = [];
  let sheetIds = [];
  
  if (lastRow > 1) {
    // Holt alle Daten ab Zeile 2 (ID ist in Spalte A)
    sheetData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    sheetIds = sheetData.map(row => {
      const val = row[0];
      // Nur rein numerische Werte als Webling-ID interpretieren
      return (!isNaN(val) && val !== "" && val !== null) ? Number(val) : null;
    });
  }

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
      
      // Werte gemäss deinen Spalten zuordnen (Passe ggf. die Feldnamen in den eckigen Klammern an!)
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
  // Anstatt sie zu löschen, markieren wir sie farblich oder schreiben "Inaktiv" dazu, 
  // um den Verlauf für dein Reservierungssystem nicht zu stören.
  if (removedMembers.length > 0) {
    sheetData.forEach((row, index) => {
      const currentId = Number(row[0]);
      if (removedMembers.includes(currentId)) {
        const rowNum = index + 2; // +2 weil Zeile 1 der Header ist und das Array bei 0 startet
        
        // Markiert die Zeile hellrot, um anzuzeigen, dass das Mitglied ausgetreten ist
        sheet.getRange(rowNum, 1, 1, 6).setBackground("#fce8e6");
        Logger.log(`Als inaktiv markiert (Rot hinterlegt): ID ${currentId} - ${row[1]} ${row[2]}`);
      }
    });
  }

  Logger.log("Synchronisation abgeschlossen.");
}
