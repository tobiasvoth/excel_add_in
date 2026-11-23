// Läuft in der Seitenleiste (Task Pane).
// Ablauf: Google-Login im Dialog -> ID-Token -> Apps Script prüft Token -> Wert in aktive Zelle schreiben.
let idToken = null;
let userEmail = null;

const el = (id) => document.getElementById(id);
const setStatus = (text) => { el("status").textContent = text; };

// Fehler sichtbar machen, statt dass etwas still scheitert
window.addEventListener("error", (e) => setStatus(`Skriptfehler: ${e.message}`));

let readyTimer;

Office.onReady((info) => {
  clearTimeout(readyTimer);
  if (info.host !== Office.HostType.Excel) {
    setStatus(`Die Seite läuft nicht in Excel (Host: ${info.host}). Bitte über das Menüband in Excel öffnen.`);
    return;
  }
  el("login").disabled = false;
  el("login").addEventListener("click", login);
  el("writeTwo").addEventListener("click", writeTwo);
  el("logout").addEventListener("click", logout);
  render();
});

// Falls Office.js nicht innerhalb von 5 Sekunden bereit ist
readyTimer = setTimeout(() => {
  setStatus("Office.js ist nicht bereit. Ist die Seite in Excel geöffnet und ist appsforoffice.microsoft.com erreichbar?");
}, 5000);

// ---------- Anzeige ----------
function render() {
  const loggedIn = !!idToken;
  el("loggedOut").hidden = loggedIn;
  el("loggedIn").hidden = !loggedIn;
  el("user").textContent = loggedIn ? `Angemeldet als ${userEmail}` : "";
}

// ---------- Token-Hilfen ----------
function decodeJwt(token) {
  let b64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
  while (b64.length % 4) b64 += "=";
  const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  return JSON.parse(new TextDecoder().decode(bytes));
}

function tokenIsValid() {
  if (!idToken) return false;
  const { exp } = decodeJwt(idToken);
  return exp * 1000 > Date.now() + 30000; // 30 s Sicherheitsabstand
}

function randomNonce() {
  return Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, "0")).join("");
}

// ---------- Anmeldung über den Office-Dialog ----------
function login() {
  setStatus("");
  const nonce = randomNonce();
  const url = new URL("login.html", location.href);
  url.searchParams.set("nonce", nonce);

  Office.context.ui.displayDialogAsync(url.href, { height: 60, width: 30 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      setStatus(`Anmeldefenster konnte nicht geöffnet werden: ${result.error.message}`);
      return;
    }
    const dialog = result.value;

    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      dialog.close();
      try {
        const msg = JSON.parse(arg.message);
        if (msg.error) throw new Error(msg.error);
        const payload = decodeJwt(msg.idToken);
        if (payload.nonce !== nonce) throw new Error("Ungültige Antwort (Nonce stimmt nicht)");
        idToken = msg.idToken;
        userEmail = payload.email;
        setStatus("");
      } catch (e) {
        idToken = null;
        userEmail = null;
        setStatus(`Anmeldung fehlgeschlagen: ${e.message}`);
      }
      render();
    });

    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      // 12006 = Nutzer hat das Fenster geschlossen
      setStatus(arg.error === 12006 ? "Anmeldung abgebrochen." : `Anmeldefenster: Fehler ${arg.error}`);
    });
  });
}

function logout() {
  idToken = null;
  userEmail = null;
  setStatus("");
  render();
}

// ---------- Aufruf von Apps Script ----------
async function callBackend(action) {
  let res;
  try {
    // text/plain vermeidet die CORS-Preflight-Anfrage, die Apps Script nicht beantwortet
    res = await fetch(CONFIG.API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ idToken, action })
    });
  } catch (e) {
    throw new Error(
      "Das Backend ist nicht erreichbar oder wurde von Google blockiert. " +
      "Steht die Web-App auf „Zugriff: Jeder“ und stimmt API_URL in config.js?"
    );
  }

  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error(
      `Das Backend hat keine gültige Antwort geliefert (Status ${res.status}). ` +
      "Meist ist das die Google-Anmeldeseite: Web-App auf „Zugriff: Jeder“ stellen und eine neue Version bereitstellen."
    );
  }
}

// ---------- Wert vom Backend holen und in die aktive Zelle schreiben ----------
async function writeTwo() {
  const button = el("writeTwo");
  setStatus("");

  if (!tokenIsValid()) {
    logout();
    setStatus("Die Anmeldung ist abgelaufen. Bitte neu anmelden.");
    return;
  }

  button.disabled = true;
  setStatus("Rufe Backend auf …");

  try {
    const data = await callBackend("value");

    if (!data.ok) {
      if (data.error === "not_signed_in") {
        logout();
        setStatus("Nicht angemeldet. Bitte neu anmelden.");
      } else if (data.error === "not_allowed") {
        setStatus("Du bist für diese Funktion nicht freigegeben.");
      } else {
        setStatus(`Fehler im Backend (${data.error}). Im Apps-Script-Editor unter „Ausführungen“ steht die Ursache.`);
      }
      return;
    }

    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load("address");
      cell.values = [[data.value]];
      await context.sync();
      setStatus(`Der Wert ${data.value} wurde in ${cell.address} geschrieben.`);
    });
  } catch (error) {
    setStatus(`Fehler: ${error.message}`);
    console.error(error);
  } finally {
    button.disabled = false;
  }
}
