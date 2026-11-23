// Apps-Script-Backend für das Excel-Add-in "Schreibe Zwei".
// Bereitstellung als Web-App: "Ausführen als: Ich", "Zugriff: Jeder".
// Die Zugriffskontrolle passiert hier im Code (Google-ID-Token + Freigabeliste).

// Gleiche Client-ID wie in config.js (GOOGLE_CLIENT_ID)
const CLIENT_ID = "554034917596-hm7k5n64v30d8dm8sqaqh327sqqk5b4l.apps.googleusercontent.com";

// Optional: zusätzliche E-Mail-Adressen, die immer zugelassen sind
// (z. B. Personen, die per Gruppe/Domain Zugriff haben, aber nicht einzeln freigegeben sind).
const EXTRA_ALLOWED = [];

function doGet() {
  return json_({ ok: true, info: "API läuft. Bitte per POST aufrufen." });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const user = verifyIdToken_(req.idToken);

    if (!isAllowed_(user.email)) {
      return json_({ ok: false, code: "FORBIDDEN", error: "Kein Zugriff für " + user.email });
    }

    switch (req.action) {
      case "getTwo":
        return json_({ ok: true, result: 2, user: user.email });
      default:
        return json_({ ok: false, error: "Unbekannte Aktion: " + req.action });
    }
  } catch (err) {
    return json_({ ok: false, code: err.code || "ERROR", error: String(err.message || err) });
  }
}

// Prüft das Google-ID-Token über den tokeninfo-Endpunkt von Google.
function verifyIdToken_(idToken) {
  if (!idToken) throw authError_("Kein Token übermittelt");

  const resp = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) throw authError_("Token ungültig oder abgelaufen");

  const claims = JSON.parse(resp.getContentText());
  if (claims.aud !== CLIENT_ID) throw authError_("Token gehört zu einer anderen App");
  if (String(claims.email_verified) !== "true") throw authError_("E-Mail nicht verifiziert");
  if (!claims.email) throw authError_("Keine E-Mail im Token");
  return claims;
}

// Zugelassen sind: Besitzer + alle, die das Skript-Projekt (Editor/Betrachter) geteilt bekommen haben.
function isAllowed_(email) {
  const wanted = email.toLowerCase();
  const cache = CacheService.getScriptCache();
  let list = cache.get("allowed");

  if (list) {
    list = JSON.parse(list);
  } else {
    const file = DriveApp.getFileById(ScriptApp.getScriptId());
    const owner = file.getOwner();
    list = [owner ? owner.getEmail() : ""]
      .concat(file.getEditors().map((u) => u.getEmail()))
      .concat(file.getViewers().map((u) => u.getEmail()))
      .concat(EXTRA_ALLOWED)
      .filter(Boolean)
      .map((s) => s.toLowerCase());
    cache.put("allowed", JSON.stringify(list), 300); // 5 Minuten
  }
  return list.indexOf(wanted) !== -1;
}

function authError_(message) {
  const e = new Error(message);
  e.code = "AUTH";
  return e;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Einmalig im Editor ausführen, um die Berechtigungen (Drive, externe Requests) zu erteilen.
function authorize() {
  console.log(isAllowed_(Session.getEffectiveUser().getEmail()));
  UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo", { muteHttpExceptions: true });
}
