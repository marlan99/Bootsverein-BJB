function fetchWeblingMemberDetails() {
  const subdomain = "bootsclub1890"; // Ersetze dies mit deiner Webling-Subdomain
  const apiKey = PropertiesService.getScriptProperties().getProperty('WEBLING_API_KEY');
  
  if (!apiKey) {
    Logger.log("Fehler: Kein API-Key vorhanden!");
    return;
  }

  const options = {
    "method": "get",
    "headers": { "apikey": apiKey, "Accept": "application/json" },
    "muteHttpExceptions": true
  };

  // 1. Alle Mitglieder-IDs holen
  const listUrl = `https://${subdomain}.webling.ch/api/1/member`;
  const response = UrlFetchApp.fetch(listUrl, options);
  
  if (response.getResponseCode() !== 200) {
    Logger.log("Fehler beim Laden der Mitgliederliste: " + response.getContentText());
    return;
  }

  const listData = JSON.parse(response.getContentText());
  const memberIds = listData.objects || [];
  
  Logger.log(`${memberIds.length} Mitglieder gefunden. Rufe Details ab...`);

  // 2. Details für die ersten 10 Mitglieder abrufen (als Beispiel, um Rate Limits zu schonen)
  const limit = Math.min(memberIds.length, 10); 
  const membersDetails = [];

  for (let i = 0; i < limit; i++) {
    const memberId = memberIds[i];
    const detailUrl = `https://${subdomain}.webling.ch/api/1/member/${memberId}`;
    const detailResponse = UrlFetchApp.fetch(detailUrl, options);

    if (detailResponse.getResponseCode() === 200) {
      const memberData = JSON.parse(detailResponse.getContentText());
      
      // Felder extrahieren (die exakten Feldnamen hängen von deinen Webling-Einstellungen ab)
      const details = {
        id: memberId,
        vorname: memberData.properties["Vorname"] || "-",
        nachname: memberData.properties["Name"] || "-",
        email: memberData.properties["E-Mail"] || "-",
        mobile: memberData.properties["Mobile"] || "-",
        email2: memberData.properties["Webformular"] || "-"
      };
      
      membersDetails.push(details);
      Logger.log(`Geladen: ${details.id}: ${details.vorname} ${details.nachname} (${details.email} / ${details.email2} / ${details.mobile})`);
    }
    
    // Kurze Pause, um das Webling-API-Rate-Limit nicht zu überschreiten
    Utilities.sleep(100); 
  }

  return membersDetails;
}
