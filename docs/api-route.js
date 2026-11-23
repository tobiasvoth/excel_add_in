// Beispiel für Node.js + Express. npm install google-auth-library
// In dein bestehendes Backend übernehmen (app = deine Express-Instanz).
const { OAuth2Client } = require("google-auth-library");

const CLIENT_ID = "DEINE-CLIENT-ID.apps.googleusercontent.com"; // dieselbe wie in docs/config.js
const ALLOWED = new Set(["anna@firma.de", "ben@firma.de"]);      // Freigabeliste (kleingeschrieben)
const ORIGINS = ["https://tobiasvoth.github.io"];                // Herkunft des Add-ins

const client = new OAuth2Client(CLIENT_ID);

// CORS inkl. Preflight. Muss vor der Token-Prüfung stehen.
function cors(req, res, next) {
  const origin = req.headers.origin;
  if (origin && ORIGINS.includes(origin)) {
    res.set({
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Vary": "Origin"
    });
  }
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
}

async function requireUser(req, res, next) {
  try {
    const token = (req.headers.authorization || "").replace(/^Bearer /i, "");
    const ticket = await client.verifyIdToken({ idToken: token, audience: CLIENT_ID });
    const p = ticket.getPayload();
    if (!p.email_verified || !ALLOWED.has(p.email.toLowerCase())) {
      return res.status(403).json({ error: "Nicht freigegeben" });
    }
    req.user = p;
    next();
  } catch {
    res.status(401).json({ error: "Nicht angemeldet" });
  }
}

app.use("/api", cors);
app.get("/api/value", requireUser, (req, res) => res.json({ value: 2 }));
