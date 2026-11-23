// Seitenleiste: Login mit Google -> Apps Script API (scripts.run) aufrufen -> Ergebnis in aktive Zelle schreiben.
// Benötigt config.js (window.CONFIG) sowie login.html und auth-callback.html im selben Ordner.

const SESSION_KEY = "schreibeZwei.session";
const $ = (id) => document.getElementById(id);

// ---------- Sitzung (Token im localStorage der Seitenleiste) ----------

function clearSession() {
  try { localStorage.removeItem(SESSION_KEY); } catch (e) { /* ignorieren */ }
}

// Liefert { token, email } oder null, wenn nicht angemeldet / Token abgelaufen.
function getSession() {
  try {
    const s = JSON.parse(localStorage.getItem(SESSION_KEY) || "null");
    if (!s || !s.token || s.exp < Date.now() + 30000) { clearSession(); return null; }
    return s;
  } catch (e) {
    clearSession();
    return null;
  }
}

function setStatus(text, isError) {
  const el = $("status");
  el.textContent = text || "";
  el.style.color = isError ? "#c5221f" : "";
}

function render() {
  const s = getSession();
  $("loggedOut").hidden = !!s;
  $("loggedIn").hidden = !s;
  if (s) $("user").textContent = "Angemeldet" + (s.email ? " als " + s.email : "");
}

// ---------- Login (Dialog -> Google -> auth-callback.html) ----------

function login() {
  return new Promise((resolve, reject) => {
    const state = crypto.randomUUID();
    const url = new URL("login.html", location.href);
    url.searchParams.set("state", state);

    Office.context.ui.displayDialogAsync(url.href, { height: 60, width: 30 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("Dialog konnte nicht geöffnet werden: " + res.error.message));
      }
      const dialog = res.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        dialog.close();
        try {
          const msg = JSON.parse(arg.message);
          if (msg.error || !msg.accessToken) throw new Error(msg.error || "Kein Token erhalten");
          if (msg.state !== state) throw new Error("State stimmt nicht überein");

          // E-Mail nur für die Anzeige abfragen
          let email = "";
          try {
            const r = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
              headers: { Authorization: "Bearer " + msg.accessToken }
            });
            email = (await r.json()).email || "";
          } catch (e) { /* nur Anzeige, kein Abbruch */ }

          localStorage.setItem(SESSION_KEY, JSON.stringify({
            token: msg.accessToken,
            email: email,
            exp: Date.now() + msg.expiresIn * 1000
          }));
          resolve(getSession());
        } catch (e) {
          reject(e);
        }
      });

      // Nutzer hat das Fenster selbst geschlossen
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        reject(new Error("Anmeldung abgebrochen"));
      });
    });
  });
}

// ---------- Apps Script API ----------

function explainHttpError(status, data) {
  const apiMsg = data && data.error && data.error.message ? " (" + data.error.message + ")" : "";
  if (status === 404) {
    return "Skript nicht gefunden" + apiMsg + ". Prüfe: Deployment-ID in config.js, Typ „API-Ausführbare Datei“, " +
      "und ob das Skript mit demselben Google-Cloud-Projekt wie die OAuth-Client-ID verknüpft ist.";
  }
  if (status === 403) {
    return "Zugriff verweigert" + apiMsg + ". Prüfe: Ist die Apps Script API im Cloud-Projekt aktiviert und " +
      "darf dein Konto die Bereitstellung nutzen (Wer hat Zugriff)?";
  }
  return "HTTP " + status + apiMsg;
}

async function runScript(functionName, retry = true) {
  const session = getSession() || (await login());

  const res = await fetch(
    "https://script.googleapis.com/v1/scripts/" + encodeURIComponent(CONFIG.DEPLOYMENT_ID) + ":run",
    {
      method: "POST",
      headers: { Authorization: "Bearer " + session.token, "Content-Type": "application/json" },
      body: JSON.stringify({ function: functionName, parameters: [] })
    }
  );

  let data = {};
  try { data = await res.json(); } catch (e) { /* keine JSON-Antwort */ }

  if (res.status === 401 && retry) {
    clearSession();
    await login();
    return runScript(functionName, false);
  }
  if (!res.ok) throw new Error(explainHttpError(res.status, data));

  // HTTP 200, aber das Skript hat einen Fehler geworfen (z. B. "Kein Zugriff für ...")
  if (data.error) {
    const detail = data.error.details && data.error.details[0];
    throw new Error((detail && detail.errorMessage) || data.error.message || "Skriptfehler");
  }
  return data.response ? data.response.result : undefined;
}

// ---------- Hauptfunktion ----------

async function writeTwoToActiveCell() {
  const btn = $("writeTwo");
  btn.disabled = true;
  setStatus("Rufe Apps Script auf …");
  try {
    const result = await runScript("getTwo");
    await Excel.run(async (ctx) => {
      ctx.workbook.getActiveCell().values = [[result]];
      await ctx.sync();
    });
    setStatus("Fertig: " + result + " eingetragen.");
  } catch (e) {
    console.error(e);
    setStatus("Fehler: " + e.message, true);
  } finally {
    btn.disabled = false;
    render();
  }
}

// ---------- Start ----------

Office.onReady(() => {
  $("login").disabled = false;
  render();

  $("login").onclick = async () => {
    setStatus("");
    try {
      await login();
      setStatus("Angemeldet.");
    } catch (e) {
      setStatus("Fehler: " + e.message, true);
    }
    render();
  };

  $("writeTwo").onclick = writeTwoToActiveCell;

  $("logout").onclick = () => {
    clearSession();
    setStatus("Abgemeldet.");
    render();
  };
});
