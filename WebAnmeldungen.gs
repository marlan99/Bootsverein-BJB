// ==========================================
// CONFIGURATION / EINSTELLUNGEN
// ==========================================
var FORM_ID = "1g2Ij65-zo0jL8T0hi0yufe8J77iNVVZLawOyivDlFuE"; // <-- Deine Formular-ID hier rein
var ADMIN_EMAIL = Session.getActiveUser().getEmail();

// ==========================================
// 1. DIE HAUPTFUNKTION (WIRD VOM TRIGGER GESTARTET)
// ==========================================
function sendeFormularAntwortenPerMail(e) {
  // Das Formular über die ID ansteuern
  var form = FormApp.openById(FORM_ID);
  
  // Die allerletzte Antwort abgreifen, die gerade abgeschickt wurde
  var antworten = form.getResponses();
  var letzteAntwort = antworten[antworten.length - 1];
  var einzelAntworten = letzteAntwort.getItemResponses();
  
  // Variablen für die Felder initialisieren
  var datumRaw = "-";
  var datum = "-";
  var slot = "-";
  var typ = ""; // Standardmäßig leer, falls kein "Joker" gewählt wurde
  var beschreibung = "-";
  
  // Antworten sicher nach Index zuweisen
  if (einzelAntworten.length > 0) datumRaw = einzelAntworten[0].getResponse();
  if (einzelAntworten.length > 1) slot = einzelAntworten[1].getResponse();
  
  // Logik für Index 2 (Typ / Joker-Prüfung) und Index 3 (Beschreibung)
  if (einzelAntworten.length > 2) {
    var antwortIndex2 = einzelAntworten[2].getResponse();
    
    // Prüfen, ob der spezifische String in der Antwort enthalten ist
    if (antwortIndex2.includes("Joker Buchung (max. 2 pro Saison möglich)")) {
      typ = "Joker";
      
      // Falls es auch noch eine reguläre Beschreibung an Index 3 gibt, holen wir diese
      if (einzelAntworten.length > 3) {
        beschreibung = einzelAntworten[3].getResponse();
      }
    } else {
      // Wenn es kein Joker ist, bleibt typ "" (leer) und die Antwort wird zur Beschreibung
      typ = ""; 
      beschreibung = antwortIndex2;
    }
  }
  
  // Das erste Feld (Datum) in dasselbe Format wie den Zeitstempel bringen
  if (datumRaw !== "-") {
    try {
      var datumObjekt = new Date(datumRaw);
      // Prüfen, ob die Umwandlung ein gültiges Datum erzeugt hat
      if (!isNaN(datumObjekt.getTime())) {
        datum = Utilities.formatDate(datumObjekt, Session.getScriptTimeZone(), "dd.MM.yyyy");
      } else {
        datum = datumRaw; // Fallback, falls es ein normales Textfeld war
      }
    } catch(err) {
      datum = datumRaw; // Fallback bei Fehlern
    }
  }
  
  // Falls die E-Mail-Erfassung im Formular aktiv ist, holen wir den Absender
  var absenderEmail = letzteAntwort.getRespondentEmail() || "Nicht erfasst / Anonym";
  var zeitstempel = Utilities.formatDate(letzteAntwort.getTimestamp(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
  
  // Den Link zu den bisherigen Antworten (Antwortübersicht) abrufen
  var antwortenLink = form.getSummaryUrl();
  
  // Den Inhalt der E-Mail vorbereiten
  var subject = "⛵ Neue Buchungsanfrage/Reservierung (" + form.getTitle() + ")";

  // Der HTML-Body
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.5;">
        <h3 style="color: #0056b3; margin-top: 0;">Details der Buchungsanfrage:</h3>
        <b>Eingegangen am:</b> ${zeitstempel}<br>
        <b>Absender:</b> ${absenderEmail}<br>
        <hr style="border: 0; border-top: 1px solid #ccc; margin: 15px 0;">
        <b>Datum:</b> ${datum}<br>
        <b>Slot:</b> ${slot}<br>
        <b>Typ:</b> ${typ}<br>
        <b>Beschreibung:</b> ${beschreibung}<br>
        <hr style="border: 0; border-top: 1px solid #ccc; margin: 15px 0;">
        <p style="margin-bottom: 0;">
          <a href="${antwortenLink}" style="color: #0056b3; text-decoration: none; font-weight: bold;">Bisherige Antworten im Formular ansehen</a>
        </p>
    </div>
  `;

  // Plain-Text-Body: Jedes Feld steht auf einer eigenen Zeile und beginnt mit
  // dem Feldnamen, damit Reservierungssystem.gs (parseEmailTemplate) die Werte
  // korrekt per startsWith() erkennen kann.
  const plainBody =
    `Details der Buchungsanfrage:\n` +
    `Eingegangen am: ${zeitstempel}\n` +
    `Absender: ${absenderEmail}\n` +
    `\n` +
    `Datum: ${datum}\n` +
    `Slot: ${slot}\n` +
    `Typ: ${typ}\n` +
    `Beschreibung: ${beschreibung}\n` +
    `\n` +
    `Bisherige Antworten im Formular ansehen: ${antwortenLink}`;

  // Erweiterte Optionen für die Mail vorbereiten
  const advancedOptions = { 
    replyTo: absenderEmail, 
    htmlBody: htmlBody 
  };

  try {
    // Direktes Senden der Mail mit den erweiterten Optionen
    GmailApp.sendEmail(ADMIN_EMAIL, subject, plainBody, advancedOptions);
    console.info(`✅ Neue Buchungsanfrage als Mail weitergeleitet.`);
  } catch (error) {
    console.warn(`⚠️ Direktes Senden fehlgeschlagen. Fehler: ${error.message}`);
    
    try {
      // Fallback: Wenn das direkte Senden scheitert, erstellen wir einen Entwurf
      GmailApp.createDraft(ADMIN_EMAIL, subject, plainBody, advancedOptions);
      console.info(`📝 Entwurf im Postfach erstellt, da der direkte Versand fehlschlug.`);
    } catch (draftError) {
      console.error(`❌ Fehler beim Erstellen des Entwurfs: ${draftError.message}`);
    }
  }
}

// ==========================================
// 2. SETUP-FUNKTION (NUR 1x HIERNACH AUSFÜHREN)
// ==========================================
function setupTrigger() {
  var form = FormApp.openById(FORM_ID);
  
  var alleTrigger = ScriptApp.getProjectTriggers();
  for (var i = 0; i < alleTrigger.length; i++) {
    if (alleTrigger[i].getHandlerFunction() === "sendeFormularAntwortenPerMail") {
      console.info("Der Trigger existiert bereits und muss nicht neu erstellt werden.");
      return;
    }
  }
  
  ScriptApp.newTrigger("sendeFormularAntwortenPerMail")
           .forForm(form)
           .onFormSubmit()
           .create();
           
  console.info("✅ Trigger erfolgreich eingerichtet!");
}
