// Seitenleiste: Login mit Google -> Apps-Script-Web-App aufrufen -> Ergebnis in aktive Zelle schreiben.
// Benötigt config.js (window.CONFIG) und login.html + auth-callback.html im selben Ordner.

const TOKEN_KEY = "schreibeZwei.idToken";
const $ = (id) => document.getElementById(id);

// ---------- Hilfsfunktionen ----------

function decodeJwt(token) {
  const payload = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
  const json = decodeURIComponent(
    atob(payload)
      .split("")
      .map((c) => "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2))
      .join("")
  );
  return JSON.parse(json);
}

function clearToken() {
  try { localStorage.removeItem(TOKEN_KEY); } catch (e) { /* ignorieren */ }
}

// Liefert { token, email } oder null, wenn nicht angemeldet / Token abgelaufen.
function getSession() {
  try {
    const token = localStorage.getItem(TOKEN_KEY);
    if (!token) return null;
    const claims = decodeJwt(token);
    if (claims.exp * 1000 < Date.now() + 30000) { clearToken(); return null; }
    return { token, email: claims.email };
  } catch (e) {
    clearToken();
    return null;
  }
}

function setStatus(text, isError) {
  const el = $("status");
  el.textContent = text || "";
  el.style.color = isError ? "#c5221f" : "";
}

function render() {
  const session = getSession();
  $("loggedOut").hidden = !!session;
  $("loggedIn").hidden = !session;
  if (session) $("user").textContent = "Angemeldet als " + session.email;
}

// ---------- Login ----------

function login() {
  return new Promise((resolve, reject) => {
    const nonce = crypto.randomUUID();
    const url = new URL("login.html", location.href);
    url.searchParams.set("nonce", nonce);

    Office.context.ui.displayDialogAsync(url.href, { height: 60, width: 30 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("Dialog konnte nicht geöffnet werden: " + res.error.message));
      }
      const dialog = res.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dialog.close();
        try {
          const msg = JSON.parse(arg.message);
          if (msg.error || !msg.idToken) throw new Error(msg.error || "Kein Token erhalten");
          const claims = decodeJwt(msg.idToken);
          if (claims.nonce !== nonce) throw new Error("Nonce stimmt nicht überein");
          localStorage.setItem(TOKEN_KEY, msg.idToken);
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

// ---------- API-Aufruf ----------

async function callApi(action, retry = true) {
  const session = getSession() || (await login());

  // "text/plain" und keine eigenen Header => einfacher Request ohne CORS-Preflight.
  // (Apps Script beantwortet keine Preflight-Requests.)
  const res = await fetch(CONFIG.API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, idToken: session.token })
  });

  const text = await res.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error(
      "Antwort ist kein JSON (HTTP " + res.status + "). Ist die Web-App für „Jeder“ freigegeben " +
      "und als neue Version bereitgestellt?"
    );
  }

  if (data.ok) return data.result;

  if (data.code === "AUTH" && retry) {
    clearToken();
    await login();
    return callApi(action, false);
  }
  throw new Error(data.error || "Unbekannter Fehler");
}

// ---------- Hauptfunktion ----------

async function writeTwoToActiveCell() {
  const btn = $("writeTwo");
  btn.disabled = true;
  setStatus("Rufe Apps Script auf …");
  try {
    const result = await callApi("getTwo");
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
    clearToken();
    setStatus("Abgemeldet.");
    render();
  };
});
