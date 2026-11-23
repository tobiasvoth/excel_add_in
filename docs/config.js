// Hier die eigenen Werte eintragen. Nichts davon ist geheim.
window.CONFIG = {
  // OAuth-Client-ID (Typ "Webanwendung") aus der Google Cloud Console
  GOOGLE_CLIENT_ID: "554034917596-hm7k5n64v30d8dm8sqaqh327sqqk5b4l.apps.googleusercontent.com",
  // Deployment-ID der "API-Ausführbaren Datei" des NEUEN Backend-Projekts
  // (Apps Script: Bereitstellen > Bereitstellungen verwalten)
  DEPLOYMENT_ID: "DEINE_DEPLOYMENT_ID",
  // Muss alle oauthScopes aus appsscript.json enthalten (userinfo.email nur für die Anzeige der Adresse)
  SCOPE: "https://www.googleapis.com/auth/cloud-platform https://www.googleapis.com/auth/script.external_request https://www.googleapis.com/auth/userinfo.email"
};
