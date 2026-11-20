// Kleiner HTTPS-Server für die Entwicklung (https://localhost:3000).
// Die Entwicklungszertifikate kommen von office-addin-dev-certs.
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = path.join(__dirname, "public");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
};

(async () => {
  const options = await devCerts.getHttpsServerOptions();

  https
    .createServer(options, (req, res) => {
      const urlPath = decodeURIComponent(req.url.split("?")[0]);
      const file = path.normalize(path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath));

      // Zugriff außerhalb von /public verhindern
      if (!file.startsWith(ROOT)) {
        res.writeHead(403).end("Forbidden");
        return;
      }

      fs.readFile(file, (err, data) => {
        if (err) {
          res.writeHead(404).end("Not found");
          return;
        }
        res.writeHead(200, {
          "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
          "Cache-Control": "no-cache",
        });
        res.end(data);
      });
    })
    .listen(PORT, () => console.log(`Add-in-Server läuft auf https://localhost:${PORT}`));
})();
