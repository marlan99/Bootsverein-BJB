function fetchWeblingData() {
  // 1. Subdomain und API-Schlüssel laden
  const subdomain = "bootsclub1890"; // Ersetze dies mit deiner Webling-Subdomain
  const apiKey = PropertiesService.getScriptProperties().getProperty('WEBLING_API_KEY');
  
  if (!apiKey) {
    Logger.log("Fehler: Kein API-Key in den Skripteigenschaften gefunden!");
    return;
  }

  // 2. Ziel-URL definieren (z.B. Endpunkt für Mitglieder)
  const url = `https://${subdomain}.webling.ch/api/1/member`;

  // 3. HTTP-Optionen festlegen (API-Key im Header übergeben)
  const options = {
    "method": "get",
    "headers": {
      "apikey": apiKey,
      "Accept": "application/json"
    },
    "muteHttpExceptions": true // Verhindert, dass das Skript bei API-Fehlern sofort abstürzt
  };

  try {
    // 4. API-Abfrage ausführen
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const data = JSON.parse(responseBody);
      Logger.log("Erfolgreich verbunden! Anzahl Datensätze: " + (data.objects ? data.objects.length : 0));
      
      // Hier kannst du die Daten weiterverarbeiten (z.B. in ein Google Sheet schreiben)
      return data;
    } else {
      Logger.log(`Fehler von Webling API (Status ${responseCode}): ${responseBody}`);
    }
  } catch (e) {
    Logger.log("Verbindungsfehler: " + e.toString());
  }
}
