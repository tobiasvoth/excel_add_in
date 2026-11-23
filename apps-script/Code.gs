// Backend für das Excel-Add-in "Cube".
// Eigenes, eigenständiges Apps-Script-Projekt (NICHT das Google-Sheets-Add-on).
// Bereitstellung: Bereitstellen > Neue Bereitstellung > Typ "API-Ausführbare Datei".
//
// Das Skript läuft als der ANGEMELDETE NUTZER (scripts.run). Dessen Token erzeugt ein ID-Token
// für das Dienstkonto (Impersonation) – genau wie im Sheets-Add-on. Wer Cube nutzen darf,
// bestimmt weiterhin IAM: Rolle "Service Account OpenID Connect Identity Token Creator"
// auf dem Dienstkonto. Ohne diese Rolle liefert generateIdToken einen 403.

// ---------- Konfiguration (identisch zum Sheets-Add-on) ----------
const SERVICE_URL = 'https://cubeproxy-554034917596.europe-west1.run.app';
const CUBE_BASE = SERVICE_URL + '/cubejs-api/v1';
const SA_EMAIL = 'cloud-run@nozmhn-co-play-1.iam.gserviceaccount.com';

// ---------- Anmeldung am Cloud-Run-Dienst (Impersonation) ----------
function getIdToken_() {
  const cache = CacheService.getUserCache(); // pro Nutzer, nicht Skript-weit
  const cached = cache.get('id_token');
  if (cached) return cached;

  const res = UrlFetchApp.fetch(
    'https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts/' + SA_EMAIL + ':generateIdToken',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify({ audience: SERVICE_URL, includeEmail: true }),
      muteHttpExceptions: true,
    }
  );
  let body = {};
  try {
    body = JSON.parse(res.getContentText());
  } catch (e) {
    throw new Error('Token-Fehler: HTTP ' + res.getResponseCode());
  }
  if (!body.token) {
    if (res.getResponseCode() === 403) {
      throw new Error(
        'Kein Zugriff auf Cube (403). Bitte beim Admin die Token-Rolle für das Dienstkonto anfragen.\n' +
          res.getContentText().slice(0, 200)
      );
    }
    throw new Error('Token-Fehler: ' + res.getContentText().slice(0, 300));
  }
  cache.put('id_token', body.token, 3000); // Token gilt 3600 s
  return body.token;
}

// Ruft den Cube-Proxy auf. Das Cloud-Run-Token geht in X-Serverless-Authorization.
function cubeFetch_(path, options) {
  const headers = { 'X-Serverless-Authorization': 'Bearer ' + getIdToken_() };
  return UrlFetchApp.fetch(
    CUBE_BASE + path,
    Object.assign({ headers: headers, muteHttpExceptions: true }, options || {})
  );
}

// ---------- Von Excel aufrufbar (öffentliche Funktionen) ----------

// Liefert alle Cubes mit Measures und Dimensions (inkl. Typ) für die Seitenleiste
function getMeta() {
  const res = cubeFetch_('/meta');
  if (res.getResponseCode() !== 200) {
    throw new Error('HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  return JSON.parse(res.getContentText()).cubes.map(function (c) {
    return {
      name: c.name,
      title: c.title || c.name,
      type: c.type || 'cube',
      group: groupOf_(c),
      measures: c.measures.map(function (m) {
        return { name: m.name, title: m.shortTitle || m.title || m.name };
      }),
      dimensions: c.dimensions.map(function (d) {
        return { name: d.name, title: d.shortTitle || d.title || d.name, type: d.type };
      }),
    };
  });
}

// Führt die Abfrage aus und liefert { values: [Kopfzeile, Zeile, ...] } (leer = keine Daten).
function runCubeQuery(query) {
  return { values: fetchTable_(query) || [] };
}

// ---------- Interne Hilfsfunktionen (unverändert aus dem Sheets-Add-on) ----------

// Kombinierbarkeit: Cubes mit gleichem connectedComponent haben mindestens einen
// Join-Pfad zueinander. Views sind vorgefertigt: nur Felder derselben View mischen.
function groupOf_(c) {
  if (c.type === 'view') return 'view:' + c.name;
  if (c.connectedComponent !== undefined && c.connectedComponent !== null) {
    return 'cc:' + c.connectedComponent;
  }
  return '';
}

function fetchTable_(query) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ query: query }),
  };
  let body = {};
  for (let i = 0; i < 15; i++) {
    const res = cubeFetch_('/load', options);
    try {
      body = JSON.parse(res.getContentText());
    } catch (e) {
      throw new Error('HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
    }
    if (body.error === 'Continue wait') { Utilities.sleep(1000); continue; }
    if (res.getResponseCode() !== 200) throw new Error(body.error || 'HTTP ' + res.getResponseCode());
    break;
  }
  if (body.error) throw new Error('Cube antwortet nicht rechtzeitig: ' + body.error);

  const data = body.data || [];
  if (!data.length) return null;

  const annotation = body.annotation || {};
  const info = function (key) {
    return (
      (annotation.measures || {})[key] ||
      (annotation.dimensions || {})[key] ||
      (annotation.timeDimensions || {})[key] ||
      {}
    );
  };

  // Bei Zeitdimension mit Granularität liefert Cube zwei Spalten
  // (z. B. orders.created_at und orders.created_at.month): die doppelte entfernen
  let keys = Object.keys(data[0]);
  (query.timeDimensions || []).forEach(function (td) {
    if (td.granularity && (query.dimensions || []).indexOf(td.dimension) === -1) {
      keys = keys.filter(function (k) { return k !== td.dimension; });
    }
  });

  const isMeasure = {};
  (query.measures || []).forEach(function (m) { isMeasure[m] = true; });
  const isTime = function (k) {
    return info(k).type === 'time' || !!(annotation.timeDimensions || {})[k];
  };

  const headers = keys.map(function (k) { return info(k).title || k; });
  const rows = data.map(function (r) {
    return keys.map(function (k) {
      const v = r[k];
      if (v === null || v === undefined || v === '') return '';
      if (isMeasure[k] && !isNaN(v)) return Number(v);
      if (isTime(k) && typeof v === 'string') return cleanTime_(v);
      return v;
    });
  });
  return [headers].concat(rows);
}

// '2024-03-01T00:00:00.000' -> '2024-03-01'
function cleanTime_(v) {
  return v.replace('T', ' ').replace(/\.\d+$/, '').replace(/ 00:00:00$/, '');
}
