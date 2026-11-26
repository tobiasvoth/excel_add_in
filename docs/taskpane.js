// Excel-Add-in "Cube": gleiche Funktionen wie das Google-Sheets-Add-on.
// Ablauf: Google-Login (Dialog) -> Apps Script API (scripts.run) -> Backend holt Daten von Cube
// -> Ergebnis wird mit Office.js in die Tabelle geschrieben.
// Benötigt config.js (window.CONFIG) sowie login.html und auth-callback.html im selben Ordner.
// Backend: Backend.gs (getMeta, runCubeQuery(query, extras)).

var BUILD = "cube4";
var SESSION_KEY = "cube.session"; // eigener Schlüssel: alte Tokens mit anderen Berechtigungen werden nicht genutzt
var SAVED_KEY = "cube.saved";     // Name der Einstellung in der Arbeitsmappe
var MAX_CELLS_PER_WRITE = 20000;  // große Ergebnisse in Portionen schreiben

var MAX_LIMIT = 50000;     // Höchstwert der Instanz
var PRESET_LIMIT = 50000;  // Vorlagen werden vollständig geladen

var META = [];             // alle Cubes und Views aus /meta
var VIEWS = [];            // nur Views mit meta.case (Case-Modus)
var CUR = null;            // im Case-Modus gewählte View
var MODE = 'case';         // 'case' oder 'dev'
var DIMS = {};             // Dimension-Name -> { title, short, type }
var FMT = {};              // Measure-Name -> d3-Format (meta.format)
var PRESET_FILTERS = [];   // Filter der geladenen Vorlage (unverändert weitergereicht)
var started = false;

// Pivot-Editor
var OUTPUT = 'flat';       // 'flat' (Tabelle) oder 'pivot' (Pivot-Tabelle)
var PIVOT = { rows: [], cols: [], vals: [] }; // Belegung der Bereiche (vals: { name, agg })
var FIELDS = {};           // Feld-Name -> { kind: 'measure' | 'dim', title }
var DEFAULT_AGG = {};      // Measure-Name -> Standard-Aggregation (meta.pivot_aggregation, sonst sum)
var AGG = { sum: 'Sum', average: 'Average', count: 'Count', min: 'Min', max: 'Max' }; // Excel.AggregationFunction
var AGG_LABEL = { sum: 'Summe', average: 'Mittelwert', count: 'Anzahl', min: 'Minimum', max: 'Maximum' };

// Operatoren je Datentyp (Cube-Filteroperatoren)
var OPS = {
  string: [
    ['equals', 'ist gleich'], ['notEquals', 'ist nicht gleich'],
    ['contains', 'enthält'], ['notContains', 'enthält nicht'],
    ['startsWith', 'beginnt mit'], ['endsWith', 'endet mit'],
    ['set', 'ist gesetzt'], ['notSet', 'ist leer']
  ],
  number: [
    ['equals', '='], ['notEquals', '≠'],
    ['gt', '>'], ['gte', '≥'], ['lt', '<'], ['lte', '≤'],
    ['set', 'ist gesetzt'], ['notSet', 'ist leer']
  ],
  boolean: [
    ['equals', 'ist'], ['notEquals', 'ist nicht'],
    ['set', 'ist gesetzt'], ['notSet', 'ist leer']
  ]
};

function $(id) { return document.getElementById(id); }

function setStatus(msg, isErr) {
  var s = $('status');
  msg = msg || '';
  s.textContent = msg;
  s.className = isErr ? 'err' : (msg.indexOf('⚠') !== -1 ? 'warn' : '');
}

function setBusy(busy) {
  Array.prototype.forEach.call(document.querySelectorAll('button'), function (b) { b.disabled = busy; });
  if (!busy) fillPresetsState();
}

// =====================================================================
// Anmeldung (Token im localStorage der Seitenleiste)
// =====================================================================

function clearSession() {
  try { localStorage.removeItem(SESSION_KEY); } catch (e) { /* ignorieren */ }
}

// Liefert { token, email, exp } oder null, wenn nicht angemeldet / Token abgelaufen.
function getSession() {
  try {
    var s = JSON.parse(localStorage.getItem(SESSION_KEY) || 'null');
    if (!s || !s.token || s.exp < Date.now() + 30000) { clearSession(); return null; }
    return s;
  } catch (e) {
    clearSession();
    return null;
  }
}

function renderAuth() {
  var s = getSession();
  $('loggedOut').hidden = !!s;
  $('app').hidden = !s;
  if (s) $('user').textContent = 'Angemeldet' + (s.email ? ' als ' + s.email : '');
}

function login() {
  return new Promise(function (resolve, reject) {
    var state = crypto.randomUUID();
    var url = new URL('login.html', location.href);
    url.searchParams.set('state', state);

    Office.context.ui.displayDialogAsync(url.href, { height: 60, width: 30 }, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error('Dialog konnte nicht geöffnet werden: ' + res.error.message));
      }
      var dialog = res.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async function (arg) {
        dialog.close();
        try {
          var msg = JSON.parse(arg.message);
          if (msg.error || !msg.accessToken) throw new Error(msg.error || 'Kein Token erhalten');
          if (msg.state !== state) throw new Error('State stimmt nicht überein');

          // E-Mail nur für die Anzeige abfragen
          var email = '';
          try {
            var r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
              headers: { Authorization: 'Bearer ' + msg.accessToken }
            });
            email = (await r.json()).email || '';
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
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        reject(new Error('Anmeldung abgebrochen'));
      });
    });
  });
}

// =====================================================================
// Backend aufrufen (Apps Script API, scripts.run)
// =====================================================================

function explainHttpError(status, data) {
  var apiMsg = data && data.error && data.error.message ? ' (' + data.error.message + ')' : '';
  if (status === 404) {
    return 'Skript nicht gefunden' + apiMsg + '. Prüfe: Deployment-ID in config.js, Typ „API-Ausführbare Datei“, ' +
      'und ob das Backend-Projekt mit demselben Google-Cloud-Projekt wie die OAuth-Client-ID verknüpft ist.';
  }
  if (status === 403) {
    return 'Zugriff verweigert' + apiMsg + '. Prüfe: Apps Script API im Cloud-Projekt aktiviert? ' +
      'Darf dein Konto die Bereitstellung nutzen (Wer hat Zugriff)? Wurde beim Login alles erlaubt?';
  }
  return 'HTTP ' + status + apiMsg;
}

async function runScript(functionName, parameters, retry) {
  if (retry === undefined) retry = true;
  var session = getSession() || (await login());
  renderAuth();

  var res = await fetch(
    'https://script.googleapis.com/v1/scripts/' + encodeURIComponent(CONFIG.DEPLOYMENT_ID) + ':run',
    {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + session.token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ function: functionName, parameters: parameters || [] })
    }
  );

  var data = {};
  try { data = await res.json(); } catch (e) { /* keine JSON-Antwort */ }

  if (res.status === 401 && retry) {
    clearSession();
    await login();
    renderAuth();
    return runScript(functionName, parameters, false);
  }
  if (!res.ok) throw new Error(explainHttpError(res.status, data));

  // HTTP 200, aber das Skript hat einen Fehler geworfen (z. B. "Kein Zugriff auf Cube (403)")
  if (data.error) {
    var detail = data.error.details && data.error.details[0];
    throw new Error((detail && detail.errorMessage) || data.error.message || 'Skriptfehler');
  }
  return data.response ? data.response.result : undefined;
}

// Abfrage ausführen: liefert { values, formats, warn } oder null (keine Daten).
// formats: { 'view.measure': d3-Format } -> das Backend liefert daraus das Excel-Zahlenformat je Spalte.
// Ein älteres Backend (nur { values }) funktioniert weiter, dann ohne Formate und ohne Limit-Hinweis.
async function fetchTable(query, formats) {
  var res = await runScript('runCubeQuery', [query, { formats: formats || {} }]);
  if (!res || !res.values || !res.values.length) return null;
  return { values: res.values, formats: res.formats || [], warn: res.warn || '', keys: res.keys || null };
}

// =====================================================================
// In die Tabelle schreiben (Office.js)
// =====================================================================

// Texte, die mit = + - @ beginnen, würde Excel als Formel auslesen: als Text kennzeichnen
function escapeCell(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

async function sheetExists(sheetId) {
  return Excel.run(async function (ctx) {
    var sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetId);
    await ctx.sync();
    return !sheet.isNullObject;
  });
}

// Zahlenformat je Spalte setzen (nur Datenzeilen, nicht die Kopfzeile)
async function applyFormats(ctx, sheet, row, col, rows, formats, skipRows) {
  if (skipRows === undefined || skipRows === null) skipRows = 1;
  if (!formats || rows <= skipRows) return;
  for (var c = 0; c < formats.length; c++) {
    if (!formats[c]) continue;
    for (var i = skipRows; i < rows; i += MAX_CELLS_PER_WRITE) {
      var n = Math.min(MAX_CELLS_PER_WRITE, rows - i);
      var fmt = [];
      for (var k = 0; k < n; k++) fmt.push([formats[c]]);
      sheet.getRangeByIndexes(row + i, col + c, n, 1).numberFormat = fmt;
      await ctx.sync();
    }
  }
}

// Ohne target: ab der aktiven Zelle. Mit target ({ sheetId, row, col }, 0-basiert): an dieser Position.
// formats (optional): Excel-Zahlenformat je Spalte oder null. headerRows: Zeilen oben ohne Zahlenformat (Standard 1).
// Liefert die geschriebene Position und Größe zurück.
async function writeTable(values, target, formats, headerRows) {
  return Excel.run(async function (ctx) {
    var sheet, row, col, sheetId;
    if (target) {
      sheet = ctx.workbook.worksheets.getItemOrNullObject(target.sheetId);
      await ctx.sync();
      if (sheet.isNullObject) throw new Error('Das Zielblatt existiert nicht mehr.');
      row = target.row; col = target.col; sheetId = target.sheetId;
    } else {
      var cell = ctx.workbook.getActiveCell();
      cell.load(['rowIndex', 'columnIndex']);
      sheet = cell.worksheet;
      sheet.load('id');
      await ctx.sync();
      row = cell.rowIndex; col = cell.columnIndex; sheetId = sheet.id;
    }

    var rows = values.length, cols = values[0].length;
    if (row + rows > 1048576 || col + cols > 16384) {
      throw new Error('Das Ergebnis passt ab dieser Zelle nicht mehr ins Tabellenblatt.');
    }

    var chunk = Math.max(1, Math.floor(MAX_CELLS_PER_WRITE / cols));
    for (var i = 0; i < rows; i += chunk) {
      var part = values.slice(i, i + chunk).map(function (r) { return r.map(escapeCell); });
      sheet.getRangeByIndexes(row + i, col, part.length, cols).values = part;
      await ctx.sync();
    }
    await applyFormats(ctx, sheet, row, col, rows, formats, headerRows);
    return { sheetId: sheetId, row: row, col: col, rows: rows, cols: cols };
  });
}

// Alten Ergebnisbereich leeren (nur Inhalte, Formatierung bleibt)
async function clearArea(q, all) {
  if (!q.rows || !q.cols) return;
  await Excel.run(async function (ctx) {
    var sheet = ctx.workbook.worksheets.getItemOrNullObject(q.sheetId);
    await ctx.sync();
    if (sheet.isNullObject) throw new Error('Das Zielblatt existiert nicht mehr.');
    sheet.getRangeByIndexes(q.row, q.col, q.rows, q.cols).clear(all ? 'All' : 'Contents');
    await ctx.sync();
  });
}

// =====================================================================
// Gespeicherte Abfragen (liegen in den Einstellungen der Arbeitsmappe)
// =====================================================================

async function loadSavedList() {
  var list = [];
  await Excel.run(async function (ctx) {
    var item = ctx.workbook.settings.getItemOrNullObject(SAVED_KEY);
    item.load('value');
    await ctx.sync();
    if (!item.isNullObject) {
      try { list = JSON.parse(item.value) || []; } catch (e) { list = []; }
    }
  });
  list.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return list;
}

async function storeSavedList(list) {
  await Excel.run(async function (ctx) {
    ctx.workbook.settings.add(SAVED_KEY, JSON.stringify(list));
    await ctx.sync();
  });
}

async function saveQuery(name, query, target, formats, pivot) {
  var list = await loadSavedList();
  var existing = list.filter(function (q) { return q.name === name; })[0];
  var entry = {
    id: existing ? existing.id : crypto.randomUUID(),
    name: name,
    query: query,
    formats: formats || {},
    pivot: pivot || null,
    sheetId: target.sheetId,
    row: target.row,
    col: target.col,
    rows: target.rows,
    cols: target.cols,
    lastRun: new Date().toISOString()
  };
  list = list.filter(function (q) { return q.name !== name; }).concat([entry]);
  await storeSavedList(list);
}

function summary(query) {
  var n = function (a) { return (a || []).length; };
  var parts = [n(query.measures) + ' Measures', n(query.dimensions) + ' Dimensions'];
  if (n(query.filters)) parts.push(n(query.filters) + ' Filter');
  if (n(query.timeDimensions)) parts.push('Zeitraum');
  return parts.join(' · ');
}

// Liste für die Anzeige (mit Position wie "Tabelle1!B2")
async function getSavedQueries() {
  var list = await loadSavedList();
  var out = list.map(function (q) {
    return { id: q.id, name: q.name, where: 'Blatt nicht gefunden', lastRun: q.lastRun, summary: summary(q.query) + (q.pivot ? ' · Pivot' : '') };
  });
  await Excel.run(async function (ctx) {
    var sheets = list.map(function (q) { return ctx.workbook.worksheets.getItemOrNullObject(q.sheetId); });
    await ctx.sync();
    var ranges = list.map(function (q, i) {
      if (sheets[i].isNullObject) return null;
      var r = sheets[i].getRangeByIndexes(q.row, q.col, 1, 1);
      r.load('address');
      return r;
    });
    await ctx.sync();
    ranges.forEach(function (r, i) { if (r) out[i].where = r.address; });
  });
  return out;
}

async function deleteSavedQuery(id) {
  var list = await loadSavedList();
  await storeSavedList(list.filter(function (q) { return q.id !== id; }));
}

// Holt frische Daten und schreibt sie mit derselben Abfrage an die gespeicherte Position.
// Erst abfragen, dann den alten Bereich leeren: bei einem Fehler bleiben die alten Daten stehen.
async function refreshSaved(q) {
  if (!(await sheetExists(q.sheetId))) throw new Error('Das Zielblatt existiert nicht mehr.');
  if (q.pivot) {
    var pmsg = await refreshPivot(q);
    q.lastRun = new Date().toISOString();
    return pmsg;
  }
  var table = await fetchTable(q.query, q.formats);
  await clearArea(q);

  var msg;
  if (!table) {
    q.rows = 0; q.cols = 0;
    msg = 'Keine Daten';
  } else {
    var t = await writeTable(table.values, q, table.formats);
    q.rows = t.rows; q.cols = t.cols;
    msg = (table.values.length - 1) + ' Zeilen aktualisiert';
    if (table.warn) msg += '\n' + table.warn;
  }
  q.lastRun = new Date().toISOString();
  return msg;
}

// =====================================================================
// Gespeicherte Abfragen: Anzeige und Aktionen
// =====================================================================

async function loadSaved() {
  try {
    renderSaved(await getSavedQueries());
  } catch (e) {
    setStatus(e.message, true);
  }
}

function formatTime(iso) {
  var d = new Date(iso);
  return isNaN(d.getTime()) ? '' : d.toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' });
}

function renderSaved(list) {
  var box = $('saved');
  box.textContent = '';
  $('savedEmpty').style.display = list.length ? 'none' : '';
  $('refreshAll').style.display = list.length > 1 ? '' : 'none';

  list.forEach(function (q) {
    var row = document.createElement('div');
    row.className = 'srow';

    var info = document.createElement('div');
    info.className = 'sinfo';
    var name = document.createElement('div');
    name.className = 'sname';
    name.textContent = q.name;
    name.title = q.name;
    var meta = document.createElement('div');
    meta.className = 'small';
    meta.textContent = q.where + ' · ' + q.summary;
    var last = document.createElement('div');
    last.className = 'small';
    last.textContent = q.lastRun ? 'Zuletzt: ' + formatTime(q.lastRun) : '';
    info.appendChild(name); info.appendChild(meta); info.appendChild(last);

    var rb = document.createElement('button');
    rb.textContent = '↻ Aktualisieren';
    rb.title = 'Frische Daten holen und an die gespeicherte Stelle schreiben';
    rb.onclick = function () { refreshOne(q.id); };

    var db = document.createElement('button');
    db.className = 'sec';
    db.textContent = '✕';
    db.title = 'Gespeicherte Abfrage löschen (die Daten in der Tabelle bleiben)';
    var armed = false, timer = null;
    db.onclick = function () {
      if (!armed) {
        armed = true;
        db.textContent = 'Löschen?';
        timer = setTimeout(function () { armed = false; db.textContent = '✕'; }, 3000);
        return;
      }
      clearTimeout(timer);
      removeOne(q.id);
    };

    row.appendChild(info); row.appendChild(rb); row.appendChild(db);
    box.appendChild(row);
  });
}

// Führt eine Aktion mit Statusanzeige aus und sperrt währenddessen die Buttons
async function withBusy(startMsg, fn) {
  setBusy(true);
  setStatus(startMsg || '');
  try {
    var msg = await fn();
    setStatus(msg || '', /Fehler bei/.test(msg || ''));
    return true;
  } catch (e) {
    console.error(e);
    setStatus(e.message, true);
    return false;
  } finally {
    setBusy(false);
  }
}

async function refreshOne(id) {
  await withBusy('Aktualisiere …', async function () {
    var list = await loadSavedList();
    var q = list.filter(function (x) { return x.id === id; })[0];
    if (!q) throw new Error('Gespeicherte Abfrage nicht gefunden.');
    var msg = '„' + q.name + '“: ' + (await refreshSaved(q));
    await storeSavedList(list);
    return msg;
  });
  await loadSaved();
}

// Aktualisiert alle gespeicherten Abfragen; ein Fehler stoppt die übrigen nicht.
async function refreshAll() {
  await withBusy('Aktualisiere alle …', async function () {
    var list = await loadSavedList();
    if (!list.length) return 'Keine gespeicherten Abfragen vorhanden.';
    var ok = 0, failed = [], truncated = [];
    for (var i = 0; i < list.length; i++) {
      try {
        var r = await refreshSaved(list[i]);
        ok++;
        if (r.indexOf('⚠') !== -1) truncated.push('„' + list[i].name + '“');
      } catch (e) {
        failed.push('„' + list[i].name + '“ (' + e.message + ')');
      }
    }
    await storeSavedList(list);
    var msg = ok + ' von ' + list.length + ' Abfragen aktualisiert';
    if (truncated.length) msg += '. ⚠ Zeilenlimit erreicht bei: ' + truncated.join('; ');
    if (failed.length) msg += '. Fehler bei: ' + failed.join('; ');
    return msg;
  });
  await loadSaved();
}

async function removeOne(id) {
  await withBusy('', async function () { await deleteSavedQuery(id); return ''; });
  await loadSaved();
}

// =====================================================================
// Start der Oberfläche: Cube-Modell laden
// =====================================================================

async function startApp() {
  if (started) return;
  started = true;
  setStatus('Lade Cube-Modell …');
  loadSaved();
  try {
    init((await runScript('getMeta')) || []);
  } catch (e) {
    started = false; // beim nächsten Versuch neu laden
    setStatus(e.message, true);
  }
}

function init(meta) {
  META = meta;
  DIMS = {}; FMT = {}; FIELDS = {}; DEFAULT_AGG = {};
  META.forEach(function (c) {
    c.dimensions.forEach(function (d) {
      DIMS[d.name] = { title: c.title + ' › ' + d.title, short: d.title, type: d.type };
      FIELDS[d.name] = { kind: 'dim', title: d.title };
    });
    c.measures.forEach(function (m) {
      if (m.format) FMT[m.name] = m.format;
      FIELDS[m.name] = { kind: 'measure', title: m.title };
      if (m.aggregation && AGG[m.aggregation]) DEFAULT_AGG[m.name] = m.aggregation;
    });
  });
  // Case-Modus zeigt nur Views mit meta.case
  VIEWS = META.filter(function (c) { return c.type === 'view' && c.caseName; });
  fillViewSelect();
  onModeChange();
}

// Anzeigename einer Dimension: im Case-Modus der Kurztitel, sonst mit Cube-Namen davor
function dimTitle(n) {
  var d = DIMS[n];
  return !d ? n : (MODE === 'case' ? d.short : d.title);
}

// =====================================================================
// Modus, Auswertung (View) und Vorlagen
// =====================================================================

// Auswahl: erste Ebene = meta.case (nach case_order), zweite Ebene = meta.view_label (nach view_order)
function fillViewSelect() {
  var sel = $('viewSel');
  sel.textContent = '';
  if (!VIEWS.length) {
    var none = document.createElement('option');
    none.value = ''; none.textContent = '– keine Views mit meta.case –';
    sel.appendChild(none);
    return;
  }
  var cases = {};
  VIEWS.forEach(function (v) { (cases[v.caseName] = cases[v.caseName] || []).push(v); });
  var minOrder = function (name) {
    return Math.min.apply(null, cases[name].map(function (v) { return v.caseOrder; }));
  };
  Object.keys(cases)
    .sort(function (a, b) { return minOrder(a) - minOrder(b) || a.localeCompare(b); })
    .forEach(function (name) {
      var g = document.createElement('optgroup');
      g.label = name;
      cases[name]
        .sort(function (a, b) { return a.viewOrder - b.viewOrder || a.viewLabel.localeCompare(b.viewLabel); })
        .forEach(function (v) {
          var o = document.createElement('option');
          o.value = v.name; o.textContent = v.viewLabel;
          g.appendChild(o);
        });
      sel.appendChild(g);
    });
}

function currentView() {
  var name = $('viewSel').value;
  return VIEWS.filter(function (v) { return v.name === name; })[0] || null;
}

function onModeChange() {
  MODE = $('mode').value;
  var dev = MODE === 'dev';
  $('caseBox').style.display = dev ? 'none' : '';
  $('devHint').style.display = dev ? '' : 'none';
  $('granBox').style.display = (dev && OUTPUT !== 'pivot') ? '' : 'none';
  CUR = dev ? null : currentView();
  renderLists();
  resetForm();
  fillPresets();
  if (!dev && !VIEWS.length) {
    setStatus('Im Modell wurden keine Views mit meta.case gefunden. Über „Alle Cubes und Views“ stehen alle Felder zur Verfügung.', true);
  }
}

function onViewChange() {
  CUR = currentView();
  renderLists();
  resetForm();
  fillPresets();
}

function fillPresets() {
  var sel = $('preset');
  sel.textContent = '';
  var list = (MODE === 'case' && CUR) ? CUR.presets : [];
  if (!list.length) {
    var o = document.createElement('option');
    o.value = ''; o.textContent = '– keine Vorlagen –';
    sel.appendChild(o);
  }
  list.forEach(function (p) {
    var o = document.createElement('option');
    o.value = p.name; o.textContent = p.name;
    sel.appendChild(o);
  });
  fillPresetsState();
}

function fillPresetsState() {
  var has = !!(MODE === 'case' && CUR && CUR.presets.length);
  $('preset').disabled = !has;
  $('loadPresetBtn').disabled = !has;
}

function setPresetFilters(f) {
  PRESET_FILTERS = f || [];
  $('presetInfo').style.display = PRESET_FILTERS.length ? '' : 'none';
  $('usePresetFilters').checked = true;
}

// Vorlage laden: füllt Kennzahlen, Spalten, Zeitraum und (unsichtbar) die Vorlagenfilter.
// Danach kann die Nutzerin alles anpassen; die Vorlage selbst bleibt unverändert.
function loadPreset() {
  if (!CUR) return;
  var name = $('preset').value;
  var p = (CUR.presets || []).filter(function (x) { return x.name === name; })[0];
  if (!p) { setStatus('Keine Vorlage gewählt.', true); return; }

  resetForm();

  var wanted = (p.measures || []).concat(p.dimensions || []);
  var want = {}, found = {};
  wanted.forEach(function (n) { want[n] = true; });
  Array.prototype.forEach.call(document.querySelectorAll('.list input'), function (i) {
    if (want[i.value]) { i.checked = true; found[i.value] = true; }
  });
  var missing = wanted.filter(function (n) { return !found[n]; });

  if (p.time_dimension) $('timedim').value = p.time_dimension;
  // Zeitraum tagesgenau und unverändert übernehmen
  if (p.date_range && p.date_range.length === 2) {
    $('d1').value = String(p.date_range[0]).slice(0, 10);
    $('d2').value = String(p.date_range[1]).slice(0, 10);
  }
  $('limit').value = PRESET_LIMIT;
  // Name vorbelegen, damit sich das Ergebnis später per Klick aktualisieren lässt
  $('saveName').value = p.name + ' (' + CUR.viewLabel + ')';
  setPresetFilters(p.filters || []);

  if (OUTPUT === 'pivot') syncZones(wanted);
  updateCounts();
  applyFilter();

  if (missing.length) {
    setStatus('Vorlage „' + p.name + '“ geladen, aber diese Felder gibt es in der View nicht: ' + missing.join(', '), true);
  } else {
    setStatus('Vorlage „' + p.name + '“ geladen. Mit „In Tabelle schreiben“ ausführen; der Name unter „Abfrage merken“ ist vorbelegt, damit du das Ergebnis später aktualisieren kannst.');
  }
}

// =====================================================================
// Felder, Filter, Zeitraum
// =====================================================================

// Felder einer View nach Ordnern (folders) gruppieren, Reihenfolge wie geliefert.
// Felder ohne Ordner kommen unter „Weitere“ ans Ende.
function folderGroups(view, kind) {
  var byName = {}, used = {}, out = [];
  view[kind].forEach(function (it) { byName[it.name] = it; });
  (view.folders || []).forEach(function (f) {
    var items = [];
    (f.members || []).forEach(function (n) {
      if (byName[n] && !used[n]) { items.push(byName[n]); used[n] = true; }
    });
    if (items.length) out.push({ label: f.name, items: items });
  });
  var rest = view[kind].filter(function (it) { return !used[it.name]; });
  if (rest.length) out.push({ label: 'Weitere', items: rest });
  return out;
}

// Baut Measures, Dimensions und Zeitdimensionen für den aktuellen Modus neu auf
function renderLists() {
  var measures = $('measures'), dims = $('dimensions'), td = $('timedim');
  measures.textContent = ''; dims.textContent = '';
  while (td.options.length > 1) td.remove(1);

  var members = {};
  META.forEach(function (c) {
    if (c.group) (members[c.group] = members[c.group] || []).push(c.title);
  });

  var cubes = MODE === 'case' ? (CUR ? [CUR] : []) : META;
  cubes.forEach(function (c) {
    if (MODE === 'case') {
      folderGroups(c, 'measures').forEach(function (g) { addGroup(measures, g.label, c.group, g.items, g.label); });
      folderGroups(c, 'dimensions').forEach(function (g) { addGroup(dims, g.label, c.group, g.items, g.label); });
    } else {
      var hint;
      if (c.type === 'view') hint = 'View: nur Felder dieser View kombinieren';
      else if (!c.group) hint = 'Keine Join-Information verfügbar';
      else {
        var others = members[c.group].filter(function (t) { return t !== c.title; });
        hint = others.length ? 'Kombinierbar mit: ' + others.join(', ') : 'Keine Joins zu anderen Cubes';
      }
      addGroup(measures, c.title, c.group, c.measures, hint);
      addGroup(dims, c.title, c.group, c.dimensions, hint);
    }
    c.dimensions.forEach(function (d) {
      if (d.type === 'time') {
        var o = document.createElement('option');
        o.value = d.name;
        o.setAttribute('data-group', c.group);
        o.textContent = MODE === 'case' ? d.title : c.title + ' › ' + d.title;
        td.appendChild(o);
      }
    });
  });
  updateCounts();
}

function addGroup(box, label, group, items, hint) {
  if (!items.length) return;
  var g = document.createElement('div');
  g.className = 'grp';
  g.textContent = label;
  g.title = hint || '';
  g.setAttribute('data-grp', '1');
  g.setAttribute('data-group', group);
  box.appendChild(g);
  items.forEach(function (it) {
    var l = document.createElement('label');
    l.title = it.name;
    l.draggable = true;
    l.addEventListener('dragstart', function (e) {
      e.dataTransfer.setData('text', it.name);
      e.dataTransfer.effectAllowed = 'copyMove';
    });
    l.setAttribute('data-name', it.name);
    l.setAttribute('data-group', group);
    l.setAttribute('data-text', (label + ' ' + it.title + ' ' + it.name + ' ' + (it.explanation || '')).toLowerCase());
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = it.name;
    l.appendChild(cb);
    l.appendChild(document.createTextNode(' ' + it.title));
    if (it.explanation) {
      var ex = document.createElement('div');
      ex.className = 'small expl';
      ex.textContent = it.explanation;
      l.appendChild(ex);
    }
    box.appendChild(l);
  });
}

function applyFilter() {
  var q = $('search').value.trim().toLowerCase();
  ['measures', 'dimensions'].forEach(function (id) {
    var box = $(id), anyVisible = false, kids = box.children;
    for (var i = kids.length - 1; i >= 0; i--) {
      var el = kids[i];
      if (el.getAttribute('data-grp')) {
        el.style.display = anyVisible || !q ? '' : 'none';
        anyVisible = false;
      } else {
        var show = !q || el.getAttribute('data-text').indexOf(q) !== -1 || el.querySelector('input').checked;
        el.style.display = show ? '' : 'none';
        if (show) anyVisible = true;
      }
    }
  });
}

function checked(id) {
  return Array.prototype.map.call($(id).querySelectorAll('input:checked'), function (i) { return i.value; });
}

function updateCounts() {
  $('cm').textContent = '(' + checked('measures').length + ' gewählt)';
  $('cd').textContent = '(' + checked('dimensions').length + ' gewählt)';
  refreshCompat();
  refreshFilters();
}

// Gruppe der bisherigen Auswahl (Felder oder Zeitdimension); leer = noch nichts gewählt
function selectedGroup() {
  var boxes = document.querySelectorAll('.list input:checked');
  for (var i = 0; i < boxes.length; i++) {
    var g = boxes[i].parentNode.getAttribute('data-group');
    if (g) return g;
  }
  var opt = $('timedim').options[$('timedim').selectedIndex];
  return opt ? opt.getAttribute('data-group') || '' : '';
}

// Felder anderer Gruppen (kein Join-Pfad) ausgrauen und sperren.
// Im Case-Modus gehören alle Felder zur selben View, hier wird also nichts gesperrt.
function refreshCompat() {
  var g = selectedGroup();
  Array.prototype.forEach.call(document.querySelectorAll('.list label, .list .grp'), function (el) {
    var eg = el.getAttribute('data-group');
    var ok = !g || !eg || eg === g;
    el.style.opacity = ok ? '' : '0.4';
    var input = el.querySelector('input');
    if (input) {
      input.disabled = !ok;
      el.title = ok ? el.getAttribute('data-name') : 'Nicht kombinierbar mit den gewählten Feldern';
    }
  });
  var td = $('timedim');
  Array.prototype.forEach.call(td.options, function (o) {
    var og = o.getAttribute('data-group');
    o.disabled = !!(g && og && og !== g);
  });
  if (td.options[td.selectedIndex].disabled) td.value = '';
}

// Standard-Zeitdimension: im Case-Modus view.report_date (sonst die erste Zeitdimension)
function defaultTimeDim() {
  if (MODE !== 'case' || !CUR) return '';
  var opts = $('timedim').options, first = '';
  for (var i = 0; i < opts.length; i++) {
    var v = opts[i].value;
    if (!v) continue;
    if (/\.report_date$/.test(v)) return v;
    if (!first) first = v;
  }
  return first;
}

// ---------- Filter ----------
function addFilter() {
  var row = document.createElement('div');
  row.className = 'frow';

  var dim = document.createElement('select'); dim.className = 'f-dim';
  var op = document.createElement('select'); op.className = 'f-op';
  var val = document.createElement('input');
  val.type = 'text'; val.className = 'f-val'; val.placeholder = 'Wert(e), mit ; trennen';
  var rm = document.createElement('button');
  rm.className = 'sec'; rm.textContent = 'Filter entfernen';
  rm.onclick = function () { row.parentNode.removeChild(row); };

  dim.onchange = function () { fillOps(row); };
  op.onchange = function () { toggleVal(row); };

  row.appendChild(dim); row.appendChild(op); row.appendChild(val); row.appendChild(rm);
  $('filters').appendChild(row);
  fillDims(row);
}

// Dimension-Auswahl aus den angehakten Dimensions (ohne Zeitdimensionen) füllen
function fillDims(row) {
  var dim = row.querySelector('.f-dim');
  var old = dim.value;
  var names = checked('dimensions').filter(function (n) {
    return DIMS[n] && DIMS[n].type !== 'time';
  });
  dim.textContent = '';
  if (!names.length) {
    var o = document.createElement('option');
    o.value = ''; o.textContent = '– erst Dimension auswählen –';
    dim.appendChild(o);
  }
  names.forEach(function (n) {
    var o = document.createElement('option');
    o.value = n; o.textContent = dimTitle(n);
    dim.appendChild(o);
  });
  if (names.indexOf(old) !== -1) dim.value = old;
  fillOps(row);
}

function fillOps(row) {
  var dim = row.querySelector('.f-dim').value;
  var op = row.querySelector('.f-op');
  var old = op.value;
  var list = OPS[(DIMS[dim] || {}).type] || OPS.string;
  op.textContent = '';
  list.forEach(function (p) {
    var o = document.createElement('option');
    o.value = p[0]; o.textContent = p[1];
    op.appendChild(o);
  });
  if (list.some(function (p) { return p[0] === old; })) op.value = old;
  toggleVal(row);
}

function toggleVal(row) {
  var op = row.querySelector('.f-op').value;
  var val = row.querySelector('.f-val');
  var noValue = op === 'set' || op === 'notSet';
  val.disabled = noValue;
  if (noValue) val.value = '';
}

function refreshFilters() {
  Array.prototype.forEach.call($('filters').children, function (row) { fillDims(row); });
}

function collectFilters() {
  var out = [];
  var rows = $('filters').children;
  for (var i = 0; i < rows.length; i++) {
    var member = rows[i].querySelector('.f-dim').value;
    var op = rows[i].querySelector('.f-op').value;
    var raw = rows[i].querySelector('.f-val').value;
    if (!member) return { error: 'Jeder Filter braucht eine ausgewählte Dimension.' };
    if (op === 'set' || op === 'notSet') {
      out.push({ member: member, operator: op });
    } else {
      var values = raw.split(';').map(function (s) { return s.trim(); }).filter(Boolean);
      if (!values.length) return { error: 'Bitte einen Wert für den Filter „' + dimTitle(member) + '“ eingeben.' };
      out.push({ member: member, operator: op, values: values });
    }
  }
  return { filters: out };
}

// ---------- Zurücksetzen, Abfrage bauen, ausführen ----------
// Setzt Auswahl, Filter, Zeitraum, Limit, Name und Vorlagenfilter zurück
// (Modus, gewählte View und Vorlagenliste bleiben erhalten).
function resetForm() {
  Array.prototype.forEach.call(document.querySelectorAll('.list input'), function (i) { i.checked = false; });
  $('d1').value = ''; $('d2').value = ''; $('gran').value = '';
  $('search').value = '';
  $('saveName').value = '';
  $('filters').textContent = '';
  $('limit').value = 1000;
  setPresetFilters([]);
  $('timedim').value = defaultTimeDim();
  PIVOT = { rows: [], cols: [], vals: [] };
  renderZones();
  applyFilter(); updateCounts(); setStatus('');
}

function clearAll() { resetForm(); }

// d3-Formate (meta.format) der gewählten Measures; das Backend macht daraus die Excel-Zahlenformate
function formatsFor(measures) {
  var out = {};
  measures.forEach(function (m) { if (FMT[m]) out[m] = FMT[m]; });
  return out;
}

function buildQuery() {
  var measures = checked('measures'), dimensions = checked('dimensions');
  if (!measures.length && !dimensions.length) {
    return { error: 'Wähle mindestens ein Measure oder eine Dimension.' };
  }

  var limit = parseInt($('limit').value, 10) || 1000;
  if (limit > MAX_LIMIT) return { error: 'Max. Zeilen: höchstens 50.000 (Limit der Instanz).' };
  var query = { measures: measures, dimensions: dimensions, limit: limit };

  // Zeitraum: tagesgenau, ohne Rundung. Im Case-Modus ohne Granularität
  // (Jahr, Monat und Tag kommen als Dimensions aus der Kalender-Gruppe).
  var dim = $('timedim').value, d1 = $('d1').value, d2 = $('d2').value;
  var gran = (MODE === 'dev' && OUTPUT !== 'pivot') ? $('gran').value : '';
  if ((d1 || d2 || gran) && !dim) {
    return { error: 'Wähle eine Zeitdimension für Zeitraum oder Gruppierung.' };
  }
  if ((d1 && !d2) || (!d1 && d2)) {
    return { error: 'Bitte Anfang („von“) und Ende („bis“) des Zeitraums angeben.' };
  }
  if (d1 && d2 && d2 < d1) return { error: '„bis“ darf nicht vor „von“ liegen.' };
  if (dim && (d1 || gran)) {
    var td = { dimension: dim };
    if (d1) td.dateRange = [d1, d2];
    if (gran) td.granularity = gran;
    query.timeDimensions = [td];
  }

  // Filter: Vorlagenfilter unverändert, dazu die Filter der Nutzerin (alle gelten zusammen: „und“)
  var f = collectFilters();
  if (f.error) return { error: f.error };
  var all = ($('usePresetFilters').checked ? PRESET_FILTERS : []).concat(f.filters);
  if (all.length) query.filters = all;

  return { query: query };
}

// Schreibt das Ergebnis ab der aktiven Zelle.
// Mit Namen wird die Abfrage samt Zielposition gespeichert (gleicher Name = überschreiben).
async function run() {
  var b = buildQuery();
  if (b.error) { setStatus(b.error, true); return; }
  var name = $('saveName').value.trim();
  var formats = formatsFor(b.query.measures);
  if (OUTPUT === 'pivot') return runPivot(b.query, name, formats);

  var ok = await withBusy('Abfrage läuft …', async function () {
    var table = await fetchTable(b.query, formats);
    if (!table) return 'Keine Daten' + (name ? ' (nicht gespeichert)' : '');
    var target = await writeTable(table.values, null, table.formats);
    var msg = (table.values.length - 1) + ' Zeilen geschrieben';
    if (name) {
      await saveQuery(name, b.query, target, formats);
      msg += ' · als „' + name + '“ gespeichert';
    }
    if (table.warn) msg += '\n' + table.warn;
    return msg;
  });
  if (ok && name) await loadSaved();
}

// ---------- Werkzeuge: Cube-Übersicht in die Tabelle ----------
function groupLabel(c) {
  if (c.group && c.group.indexOf('cc:') === 0) return 'Gruppe ' + c.group.slice(3);
  if (c.type === 'view') return 'View ' + c.name;
  return 'Unbekannt';
}

async function writeMeta() {
  await withBusy('Schreibe Cube-Übersicht …', async function () {
    if (!META.length) META = (await runScript('getMeta')) || [];
    var rows = [['Cube', 'Art', 'Kombinierbar (Gruppe)', 'Feldart', 'Name', 'Datentyp']];
    META.forEach(function (c) {
      var head = [c.name, c.type, groupLabel(c)];
      c.measures.forEach(function (m) { rows.push(head.concat(['Measure', m.name, ''])); });
      c.dimensions.forEach(function (d) { rows.push(head.concat(['Dimension', d.name, d.type])); });
    });
    await writeTable(rows);
    return (rows.length - 1) + ' Einträge geschrieben';
  });
}

// Zeigt, welche Cubes sich miteinander kombinieren lassen (gleiche Gruppe = Join-Pfad vorhanden)
async function writeJoinGroups() {
  await withBusy('Schreibe Gruppen …', async function () {
    if (!META.length) META = (await runScript('getMeta')) || [];
    var groups = {};
    META.forEach(function (c) {
      var label = groupLabel(c);
      (groups[label] = groups[label] || []).push(c.name);
    });
    var rows = [['Gruppe', 'Cubes (untereinander kombinierbar)', 'Anzahl']];
    Object.keys(groups).sort().forEach(function (label) {
      rows.push([label, groups[label].sort().join(', '), groups[label].length]);
    });
    await writeTable(rows);
    return (rows.length - 1) + ' Gruppen geschrieben';
  });
}

// =====================================================================
// Pivot-Editor: Felder in Zeilen / Spalten / Werte ziehen, Excel baut die Pivot-Tabelle
// =====================================================================
// Ablauf: Cube-Abfrage wie gewohnt (flache Daten, nur im Speicher) -> das Add-in pivotiert selbst
// -> nur das fertige Ergebnis landet im Blatt. Aktualisieren wiederholt die Abfrage und ersetzt den Bereich.

function zoneOf(name) {
  if (PIVOT.rows.indexOf(name) !== -1) return 'rows';
  if (PIVOT.cols.indexOf(name) !== -1) return 'cols';
  for (var i = 0; i < PIVOT.vals.length; i++) if (PIVOT.vals[i].name === name) return 'vals';
  return null;
}

// Haken in der Feldliste setzen (die Haken bestimmen, welche Felder in der Abfrage stehen)
function setBox(name, state) {
  var boxes = document.querySelectorAll('.list input');
  for (var i = 0; i < boxes.length; i++) {
    if (boxes[i].value === name) { boxes[i].checked = state; return true; }
  }
  return false;
}

function removeFromZones(name) {
  var ne = function (n) { return n !== name; };
  PIVOT.rows = PIVOT.rows.filter(ne);
  PIVOT.cols = PIVOT.cols.filter(ne);
  PIVOT.vals = PIVOT.vals.filter(function (v) { return v.name !== name; });
}

function defaultZone(name) {
  return FIELDS[name] && FIELDS[name].kind === 'measure' ? 'vals' : 'rows';
}

function addToZone(zone, name) {
  if (!FIELDS[name] || !setBox(name, true)) return; // Feld gehört nicht zur gewählten Auswertung
  removeFromZones(name);
  if (zone === 'vals') PIVOT.vals.push({ name: name, agg: DEFAULT_AGG[name] || 'sum' });
  else PIVOT[zone].push(name);
  renderZones();
  updateCounts();
}

function dropField(zone, name) {
  if (!name || !FIELDS[name]) return;
  var kind = FIELDS[name].kind;
  if (kind === 'measure' && zone !== 'vals') { setStatus('Kennzahlen gehören in „Werte“.', true); return; }
  if (kind === 'dim' && zone === 'vals') { setStatus('Dimensionen gehören in „Zeilen“ oder „Spalten“.', true); return; }
  setStatus('');
  addToZone(zone, name);
}

// Bereiche an die angehakten Felder anpassen (neue Felder in den Standardbereich, entfernte raus).
// order: gewünschte Reihenfolge für neu hinzukommende Felder (z. B. aus einer Vorlage).
function syncZones(order) {
  var sel = {};
  checked('measures').concat(checked('dimensions')).forEach(function (n) { sel[n] = true; });
  PIVOT.rows = PIVOT.rows.filter(function (n) { return sel[n]; });
  PIVOT.cols = PIVOT.cols.filter(function (n) { return sel[n]; });
  PIVOT.vals = PIVOT.vals.filter(function (v) { return sel[v.name]; });
  (order || []).concat(checked('measures'), checked('dimensions')).forEach(function (n) {
    if (!sel[n] || zoneOf(n)) return;
    if (defaultZone(n) === 'vals') PIVOT.vals.push({ name: n, agg: DEFAULT_AGG[n] || 'sum' });
    else PIVOT.rows.push(n);
  });
  renderZones();
}

// Haken in der Feldliste: im Pivot-Modus wandert das Feld in den Standardbereich bzw. wieder heraus
function onFieldToggle(e) {
  var t = e && e.target;
  if (OUTPUT === 'pivot' && t && t.type === 'checkbox') {
    if (t.checked) { if (!zoneOf(t.value)) addToZone(defaultZone(t.value), t.value); }
    else removeFromZones(t.value);
    renderZones();
  }
  updateCounts();
}

function moveInZone(zone, idx, delta) {
  var list = PIVOT[zone], j = idx + delta;
  if (j < 0 || j >= list.length) return;
  var tmp = list[idx]; list[idx] = list[j]; list[j] = tmp;
  renderZones();
}

function renderZones() {
  ['rows', 'cols', 'vals'].forEach(function (z) {
    var body = $('zone-' + z).querySelector('.zbody');
    body.textContent = '';
    var list = PIVOT[z];
    if (!list.length) {
      var empty = document.createElement('div');
      empty.className = 'zempty';
      empty.textContent = z === 'vals' ? 'Kennzahlen hierher ziehen' : 'Dimensionen hierher ziehen';
      body.appendChild(empty);
      return;
    }
    list.forEach(function (item, idx) {
      var name = z === 'vals' ? item.name : item;
      var chip = document.createElement('div');
      chip.className = 'chip';
      chip.draggable = true;
      chip.addEventListener('dragstart', function (e) {
        e.dataTransfer.setData('text', name);
        e.dataTransfer.effectAllowed = 'move';
      });

      var title = document.createElement('span');
      title.className = 'ctitle';
      title.textContent = (FIELDS[name] && FIELDS[name].title) || name;
      title.title = name;
      chip.appendChild(title);

      if (z === 'vals') {
        var sel = document.createElement('select');
        Object.keys(AGG_LABEL).forEach(function (k) {
          var o = document.createElement('option');
          o.value = k; o.textContent = AGG_LABEL[k];
          sel.appendChild(o);
        });
        sel.value = item.agg;
        sel.onchange = function () { item.agg = sel.value; };
        chip.appendChild(sel);
      }

      var up = document.createElement('button');
      up.textContent = '▲'; up.title = 'Nach oben';
      up.onclick = function () { moveInZone(z, idx, -1); };
      var down = document.createElement('button');
      down.textContent = '▼'; down.title = 'Nach unten';
      down.onclick = function () { moveInZone(z, idx, 1); };
      var rm = document.createElement('button');
      rm.textContent = '✕'; rm.title = 'Entfernen';
      rm.onclick = function () {
        removeFromZones(name);
        setBox(name, false);
        renderZones();
        updateCounts();
      };
      chip.appendChild(up); chip.appendChild(down); chip.appendChild(rm);
      body.appendChild(chip);
    });
  });
}

function initZones() {
  ['rows', 'cols', 'vals'].forEach(function (z) {
    var el = $('zone-' + z);
    el.addEventListener('dragover', function (e) { e.preventDefault(); el.classList.add('over'); });
    el.addEventListener('dragleave', function () { el.classList.remove('over'); });
    el.addEventListener('drop', function (e) {
      e.preventDefault();
      el.classList.remove('over');
      dropField(z, e.dataTransfer.getData('text'));
    });
  });
  renderZones();
}

function onOutputChange() {
  OUTPUT = $('output').value;
  var piv = OUTPUT === 'pivot';
  $('pivotBox').style.display = piv ? '' : 'none';
  $('granBox').style.display = (MODE === 'dev' && !piv) ? '' : 'none';
  if (piv) $('gran').value = '';
  $('go').textContent = piv ? 'Pivot-Tabelle erstellen' : 'In Tabelle schreiben';
  if (piv) syncZones();
}

// ---------- Pivotieren im Hintergrund ----------
// Es gibt weder ein Datenblatt noch eine Excel-PivotTable: Das Add-in holt die flachen Daten von Cube,
// bildet die Kreuztabelle selbst (Zeilen x Spalten x Werte, Gesamtergebnis mit der gewählten Aggregation)
// und schreibt nur das fertige Ergebnis ab der markierten Zelle ins Blatt.

var SEP = '\u0001';   // Trenner für Schlüssel aus mehreren Dimensionen
var SEP2 = '\u0002';

// d3-Format (meta.format) -> Excel-Zahlenformat (Office.js numberFormat, sprachunabhängig)
function excelFormat(d3fmt) {
  var m = /^(,?)\.(\d+)([f%])$/.exec(String(d3fmt || ''));
  if (!m) return null;
  var dec = parseInt(m[2], 10);
  var f = m[1] ? '#,##0' : '0';
  if (dec) f += '.' + new Array(dec + 1).join('0');
  if (m[3] === '%') f += '%';
  return f;
}

function cmpVal(a, b) {
  var ea = (a === '' || a === null || a === undefined), eb = (b === '' || b === null || b === undefined);
  if (ea || eb) return ea === eb ? 0 : (ea ? 1 : -1); // leere Werte ans Ende
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a).localeCompare(String(b), 'de', { numeric: true });
}

function cmpParts(a, b) {
  for (var i = 0; i < a.length; i++) {
    var c = cmpVal(a[i], b[i]);
    if (c) return c;
  }
  return 0;
}

// Beginnt auf Ebene `level` eine neue Gruppe (unterscheidet sich ein Wert bis zu dieser Ebene vom Vorgänger)?
function groupStart(parts, prev, level) {
  if (!prev) return true;
  for (var k = 0; k <= level; k++) if (parts[k] !== prev[k]) return true;
  return false;
}

function dimLabel(v) { return (v === '' || v === null || v === undefined) ? '(leer)' : v; }

// Aggregation einer Liste von Zahlen (wie die Zusammenfassung in einer Excel-Pivot)
function aggList(list, agg) {
  var n = list.length, i, r;
  if (!n) return '';
  if (agg === 'count') return n;
  if (agg === 'min') { r = list[0]; for (i = 1; i < n; i++) if (list[i] < r) r = list[i]; return r; }
  if (agg === 'max') { r = list[0]; for (i = 1; i < n; i++) if (list[i] > r) r = list[i]; return r; }
  r = 0;
  for (i = 0; i < n; i++) r += list[i];
  return agg === 'average' ? r / n : r;
}

// Spalte im Ergebnis je Feld: Neues Backend liefert keys (Feldname je Spalte), sonst Zuordnung über den Titel
function columnIndexMap(table, query) {
  var orig = table.values[0], map = {};
  query.measures.concat(query.dimensions).forEach(function (m) {
    var i = table.keys ? table.keys.indexOf(m) : orig.indexOf(FIELDS[m] && FIELDS[m].title);
    if (i < 0) throw new Error('Die Spalte für „' + ((FIELDS[m] && FIELDS[m].title) || m) + '“ fehlt im Ergebnis.');
    map[m] = i;
  });
  return map;
}

function fieldTitle(n) { return (FIELDS[n] && FIELDS[n].title) || n; }

// Baut aus den flachen Cube-Daten die Kreuztabelle.
// pv: { rows: [Feld], cols: [Feld], vals: [{ name, agg }], totals: bool, repeat: bool }
// Layout: Kopfzeilen (je Spaltenfeld eine Ebene, zuletzt die Kennzahlen), links die Zeilenfelder,
// optional Gesamtergebnis als letzte Spalte und letzte Zeile.
function buildGrid(table, query, pv) {
  var rowDims = pv.rows, colDims = pv.cols, vals = pv.vals;
  var R = rowDims.length, C = colDims.length, M = vals.length;
  var idx = columnIndexMap(table, query);
  var rowKeys = {}, colKeys = {}, cells = {}, rowTot = {}, colTot = {}, grand = [];
  var m, i, j, k;
  var emptyLists = function () { var a = []; for (var q = 0; q < M; q++) a.push([]); return a; };
  for (m = 0; m < M; m++) grand.push([]);

  // 1. Zeilen den Zellen zuordnen (Cube liefert je Dimensionskombination eine Zeile)
  for (i = 1; i < table.values.length; i++) {
    var row = table.values[i];
    var rp = [], cp = [];
    for (k = 0; k < R; k++) rp.push(row[idx[rowDims[k]]]);
    for (k = 0; k < C; k++) cp.push(row[idx[colDims[k]]]);
    var rk = rp.join(SEP), ck = cp.join(SEP);
    if (!rowKeys[rk]) { rowKeys[rk] = rp; rowTot[rk] = emptyLists(); }
    if (!colKeys[ck]) { colKeys[ck] = cp; colTot[ck] = emptyLists(); }
    var cell = cells[rk + SEP2 + ck] || (cells[rk + SEP2 + ck] = emptyLists());
    for (m = 0; m < M; m++) {
      var v = row[idx[vals[m].name]];
      if (typeof v === 'number') {
        cell[m].push(v); rowTot[rk][m].push(v); colTot[ck][m].push(v); grand[m].push(v);
      }
    }
  }

  var rks = Object.keys(rowKeys).sort(function (a, b) { return cmpParts(rowKeys[a], rowKeys[b]); });
  var cks = Object.keys(colKeys).sort(function (a, b) { return cmpParts(colKeys[a], colKeys[b]); });

  var totCol = !!pv.totals && C > 0;
  var totRow = !!pv.totals && R > 0;
  var H = C > 0 ? C + 1 : 1;                        // Kopfzeilen
  var width = R + cks.length * M + (totCol ? M : 0);
  var nRows = H + rks.length + (totRow ? 1 : 0);
  if (width > 16384) {
    throw new Error('Die Pivot hätte ' + width + ' Spalten (Excel erlaubt 16.384). Verschiebe Felder von „Spalten“ nach „Zeilen“.');
  }
  if (nRows * width > 4000000) {
    throw new Error('Die Pivot wäre zu groß (' + nRows + ' Zeilen × ' + width + ' Spalten). Schränke Zeitraum oder Felder ein.');
  }

  var blank = function () { var a = new Array(width); for (var q = 0; q < width; q++) a[q] = ''; return a; };
  var mlabel = vals.map(function (v) {
    var t = fieldTitle(v.name);
    return v.agg === 'sum' ? t : t + ' (' + AGG_LABEL[v.agg] + ')';
  });
  var fmtFor = function (v) { return v.agg === 'count' ? '#,##0' : excelFormat(FMT[v.name]); };

  // 2. Kopfzeilen
  var values = [];
  for (j = 0; j < H; j++) values.push(blank());
  var hdr = values[H - 1];
  for (k = 0; k < R; k++) hdr[k] = fieldTitle(rowDims[k]);
  if (C > 0) {
    if (R > 0) for (k = 0; k < C; k++) values[k][R - 1] = fieldTitle(colDims[k]);
    cks.forEach(function (ck, jj) {
      var parts = colKeys[ck], prev = jj > 0 ? colKeys[cks[jj - 1]] : null;
      for (var h = 0; h < C; h++) {
        if (groupStart(parts, prev, h)) values[h][R + jj * M] = dimLabel(parts[h]);
      }
      for (var mm = 0; mm < M; mm++) hdr[R + jj * M + mm] = mlabel[mm];
    });
    if (totCol) {
      values[0][R + cks.length * M] = 'Gesamtergebnis';
      for (m = 0; m < M; m++) hdr[R + cks.length * M + m] = mlabel[m];
    }
  } else {
    for (m = 0; m < M; m++) hdr[R + m] = mlabel[m];
  }

  // Zahlenformat je Spalte
  var formats = [];
  for (k = 0; k < R; k++) formats.push(null);
  for (j = 0; j < cks.length + (totCol ? 1 : 0); j++) vals.forEach(function (v) { formats.push(fmtFor(v)); });

  // 3. Datenzeilen
  rks.forEach(function (rk, ii) {
    var parts = rowKeys[rk], prev = ii > 0 ? rowKeys[rks[ii - 1]] : null;
    var out = blank();
    for (var kk = 0; kk < R; kk++) {
      out[kk] = (pv.repeat || groupStart(parts, prev, kk)) ? dimLabel(parts[kk]) : '';
    }
    cks.forEach(function (ck, jj) {
      var cell = cells[rk + SEP2 + ck];
      for (var mm = 0; mm < M; mm++) out[R + jj * M + mm] = cell ? aggList(cell[mm], vals[mm].agg) : '';
    });
    if (totCol) for (var m2 = 0; m2 < M; m2++) out[R + cks.length * M + m2] = aggList(rowTot[rk][m2], vals[m2].agg);
    values.push(out);
  });

  // 4. Gesamtergebnis (Zeile)
  if (totRow) {
    var t = blank();
    t[0] = 'Gesamtergebnis';
    cks.forEach(function (ck, jj) {
      for (var mm = 0; mm < M; mm++) t[R + jj * M + mm] = aggList(colTot[ck][mm], vals[mm].agg);
    });
    if (totCol) for (m = 0; m < M; m++) t[R + cks.length * M + m] = aggList(grand[m], vals[m].agg);
    values.push(t);
  }

  return { values: values, headerRows: H, width: width, rows: values.length, M: M, totRow: totRow, totCol: totCol, formats: formats };
}

// Kopfzeilen und Gesamtergebnis fett, Linie unter der Kopfzeile (nur Komfort: Fehler hier brechen nichts ab)
async function styleGrid(t, grid) {
  try {
    await Excel.run(async function (ctx) {
      var sheet = ctx.workbook.worksheets.getItem(t.sheetId);
      var head = sheet.getRangeByIndexes(t.row, t.col, grid.headerRows, grid.width);
      head.format.font.bold = true;
      head.format.borders.getItem('EdgeBottom').style = 'Continuous';
      if (grid.totRow) {
        var tr = sheet.getRangeByIndexes(t.row + grid.rows - 1, t.col, 1, grid.width);
        tr.format.font.bold = true;
        tr.format.borders.getItem('EdgeTop').style = 'Continuous';
      }
      if (grid.totCol) {
        sheet.getRangeByIndexes(t.row, t.col + grid.width - grid.M, grid.rows, grid.M).format.font.bold = true;
      }
      // Spaltenbreite nur an den ersten Zeilen ausrichten (schnell, auch bei vielen Zeilen)
      sheet.getRangeByIndexes(t.row, t.col, Math.min(grid.rows, 200), grid.width).format.autofitColumns();
      await ctx.sync();
    });
  } catch (e) { console.warn('Formatierung der Pivot nicht möglich:', e); }
}

// „Aktualisieren“ für eine gespeicherte Pivot: gleiche Abfrage, gleiche Belegung, gleiche Stelle.
// Erst abfragen und die Kreuztabelle bauen, dann den alten Bereich leeren: bei einem Fehler bleibt alles stehen.
async function refreshPivot(q) {
  if (!q.pivot || q.pivot.version !== 2) {
    throw new Error('Diese gespeicherte Pivot stammt aus einer älteren Version des Add-ins. Bitte neu erstellen.');
  }
  var table = await fetchTable(q.query, q.formats);
  if (!table) return 'Keine Daten';
  var grid = buildGrid(table, q.query, q.pivot);
  await clearArea(q, true);
  var t = await writeTable(grid.values, q, grid.formats, grid.headerRows);
  await styleGrid(t, grid);
  q.rows = t.rows; q.cols = t.cols;
  var msg = (table.values.length - 1) + ' Datenzeilen pivotiert · Pivot aktualisiert';
  if (table.warn) msg += '\n⚠ Die Pivot beruht auf unvollständigen Daten, die Summen stimmen nicht.\n' + table.warn;
  return msg;
}

async function runPivot(query, name, formats) {
  if (!PIVOT.vals.length || !(PIVOT.rows.length || PIVOT.cols.length)) {
    setStatus('Zieh mindestens eine Kennzahl nach „Werte“ und ein Feld nach „Zeilen“ oder „Spalten“.', true);
    return;
  }
  var pv = {
    version: 2,
    rows: PIVOT.rows.slice(),
    cols: PIVOT.cols.slice(),
    vals: PIVOT.vals.map(function (v) { return { name: v.name, agg: v.agg }; }),
    totals: $('grandTotals').checked,
    repeat: $('repeatLabels').checked
  };

  var ok = await withBusy('Abfrage läuft …', async function () {
    var table = await fetchTable(query, formats);
    if (!table) return 'Keine Daten' + (name ? ' (nicht gespeichert)' : '');
    var grid = buildGrid(table, query, pv);
    var t = await writeTable(grid.values, null, grid.formats, grid.headerRows);
    await styleGrid(t, grid);
    var msg = (table.values.length - 1) + ' Datenzeilen pivotiert · ' + (grid.rows - grid.headerRows) + ' Zeilen × ' + grid.width + ' Spalten geschrieben';
    if (name) {
      await saveQuery(name, query, t, formats, pv);
      msg += ' · als „' + name + '“ gespeichert';
    }
    if (table.warn) msg += '\n⚠ Die Pivot beruht auf unvollständigen Daten, die Summen stimmen nicht.\n' + table.warn;
    return msg;
  });
  if (ok && name) await loadSaved();
}

// =====================================================================
// Start
// =====================================================================

Office.onReady(function () {
  var id = (window.CONFIG && CONFIG.DEPLOYMENT_ID) || '';
  $('build').textContent = 'Build ' + BUILD + ' · Deployment-ID: ' + (id ? id.slice(0, 10) + '…' : 'FEHLT');

  $('measures').addEventListener('change', onFieldToggle);
  $('dimensions').addEventListener('change', onFieldToggle);
  initZones();

  $('login').disabled = false;
  renderAuth();

  $('login').onclick = async function () {
    setStatus('');
    try {
      await login();
      renderAuth();
      startApp();
    } catch (e) {
      setStatus('Fehler: ' + e.message, true);
    }
  };

  $('logout').onclick = function () {
    clearSession();
    started = false;
    setStatus('Abgemeldet.');
    renderAuth();
  };

  if (getSession()) startApp();
});
