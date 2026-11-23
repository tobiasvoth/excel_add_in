// Excel-Add-in "Cube": gleiche Funktionen wie das Google-Sheets-Add-on.
// Ablauf: Google-Login (Dialog) -> Apps Script API (scripts.run) -> Backend holt Daten von Cube
// -> Ergebnis wird mit Office.js in die Tabelle geschrieben.
// Benötigt config.js (window.CONFIG) sowie login.html und auth-callback.html im selben Ordner.

var BUILD = "cube1";
var SESSION_KEY = "cube.session"; // eigener Schlüssel: alte Tokens mit anderen Berechtigungen werden nicht genutzt
var SAVED_KEY = "cube.saved";     // Name der Einstellung in der Arbeitsmappe
var MAX_CELLS_PER_WRITE = 20000;  // große Ergebnisse in Portionen schreiben

var META = [];
var DIMS = {}; // Dimension-Name -> { title, type }
var started = false;

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
  s.textContent = msg || '';
  s.className = isErr ? 'err' : '';
}

function setBusy(busy) {
  Array.prototype.forEach.call(document.querySelectorAll('button'), function (b) { b.disabled = busy; });
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

// Abfrage ausführen: liefert [Kopfzeile, Zeilen …] oder null (keine Daten)
async function fetchTable(query) {
  var res = await runScript('runCubeQuery', [query]);
  return res && res.values && res.values.length ? res.values : null;
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

// Ohne target: ab der aktiven Zelle. Mit target ({ sheetId, row, col }, 0-basiert): an dieser Position.
// Liefert die geschriebene Position und Größe zurück.
async function writeTable(values, target) {
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
    return { sheetId: sheetId, row: row, col: col, rows: rows, cols: cols };
  });
}

// Alten Ergebnisbereich leeren (nur Inhalte, Formatierung bleibt)
async function clearArea(q) {
  if (!q.rows || !q.cols) return;
  await Excel.run(async function (ctx) {
    var sheet = ctx.workbook.worksheets.getItemOrNullObject(q.sheetId);
    await ctx.sync();
    if (sheet.isNullObject) throw new Error('Das Zielblatt existiert nicht mehr.');
    sheet.getRangeByIndexes(q.row, q.col, q.rows, q.cols).clear('Contents');
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

async function saveQuery(name, query, target) {
  var list = await loadSavedList();
  var existing = list.filter(function (q) { return q.name === name; })[0];
  var entry = {
    id: existing ? existing.id : crypto.randomUUID(),
    name: name,
    query: query,
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
    return { id: q.id, name: q.name, where: 'Blatt nicht gefunden', lastRun: q.lastRun, summary: summary(q.query) };
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

// Holt frische Daten und schreibt sie an die gespeicherte Position.
// Erst abfragen, dann den alten Bereich leeren: bei einem Fehler bleiben die alten Daten stehen.
async function refreshSaved(q) {
  if (!(await sheetExists(q.sheetId))) throw new Error('Das Zielblatt existiert nicht mehr.');
  var values = await fetchTable(q.query);
  await clearArea(q);

  var msg;
  if (!values) {
    q.rows = 0; q.cols = 0;
    msg = 'Keine Daten';
  } else {
    var t = await writeTable(values, q);
    q.rows = t.rows; q.cols = t.cols;
    msg = (values.length - 1) + ' Zeilen aktualisiert';
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
    var ok = 0, failed = [];
    for (var i = 0; i < list.length; i++) {
      try {
        await refreshSaved(list[i]);
        ok++;
      } catch (e) {
        failed.push('„' + list[i].name + '“ (' + e.message + ')');
      }
    }
    await storeSavedList(list);
    var msg = ok + ' von ' + list.length + ' Abfragen aktualisiert';
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
// Felder, Filter, Zeitraum (Oberfläche wie im Sheets-Add-on)
// =====================================================================

async function startApp() {
  if (started) return;
  started = true;
  setStatus('Lade Cube-Modell …');
  loadSaved();
  try {
    META = (await runScript('getMeta')) || [];
    renderFields();
    setStatus('');
  } catch (e) {
    started = false; // beim nächsten Versuch neu laden
    setStatus(e.message, true);
  }
}

function renderFields() {
  var measures = $('measures'), dims = $('dimensions'), td = $('timedim');
  measures.textContent = ''; dims.textContent = '';
  var members = {};
  META.forEach(function (c) {
    if (c.group) (members[c.group] = members[c.group] || []).push(c.title);
  });
  META.forEach(function (c) {
    var hint;
    if (c.type === 'view') hint = 'View: nur Felder dieser View kombinieren';
    else if (!c.group) hint = 'Keine Join-Information verfügbar';
    else {
      var others = members[c.group].filter(function (t) { return t !== c.title; });
      hint = others.length ? 'Kombinierbar mit: ' + others.join(', ') : 'Keine Joins zu anderen Cubes';
    }
    addGroup(measures, c, c.measures, hint);
    addGroup(dims, c, c.dimensions, hint);
    c.dimensions.forEach(function (d) {
      DIMS[d.name] = { title: c.title + ' › ' + d.title, type: d.type };
      if (d.type === 'time') {
        var o = document.createElement('option');
        o.value = d.name;
        o.setAttribute('data-group', c.group);
        o.textContent = c.title + ' › ' + d.title;
        td.appendChild(o);
      }
    });
  });
  measures.addEventListener('change', updateCounts);
  dims.addEventListener('change', updateCounts);
  updateCounts();
}

function addGroup(box, cube, items, hint) {
  if (!items.length) return;
  var cubeTitle = cube.title;
  var g = document.createElement('div');
  g.className = 'grp';
  g.textContent = cubeTitle;
  g.title = hint;
  g.setAttribute('data-grp', '1');
  g.setAttribute('data-group', cube.group);
  box.appendChild(g);
  items.forEach(function (it) {
    var l = document.createElement('label');
    l.title = it.name;
    l.setAttribute('data-name', it.name);
    l.setAttribute('data-group', cube.group);
    l.setAttribute('data-text', (cubeTitle + ' ' + it.title + ' ' + it.name).toLowerCase());
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = it.name;
    l.appendChild(cb);
    l.appendChild(document.createTextNode(' ' + it.title));
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

// Felder anderer Gruppen (kein Join-Pfad) ausgrauen und sperren
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
    o.value = n; o.textContent = DIMS[n].title;
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
      if (!values.length) return { error: 'Bitte einen Wert für den Filter „' + DIMS[member].title + '“ eingeben.' };
      out.push({ member: member, operator: op, values: values });
    }
  }
  return { filters: out };
}

// ---------- Zurücksetzen, Abfrage bauen, ausführen ----------
function clearAll() {
  Array.prototype.forEach.call(document.querySelectorAll('.list input'), function (i) { i.checked = false; });
  $('timedim').value = ''; $('y1').value = ''; $('y2').value = ''; $('gran').value = '';
  $('search').value = '';
  $('saveName').value = '';
  $('filters').textContent = '';
  applyFilter(); updateCounts(); setStatus('');
}

function buildQuery() {
  var measures = checked('measures'), dimensions = checked('dimensions');
  if (!measures.length && !dimensions.length) {
    return { error: 'Wähle mindestens ein Measure oder eine Dimension.' };
  }
  var query = { measures: measures, dimensions: dimensions, limit: parseInt($('limit').value, 10) || 1000 };

  var dim = $('timedim').value, y1 = $('y1').value.trim(), y2 = $('y2').value.trim(), gran = $('gran').value;
  if ((y1 || y2 || gran) && !dim) {
    return { error: 'Wähle eine Zeitdimension für Jahr oder Gruppierung.' };
  }
  if (dim) {
    var td = { dimension: dim };
    if (y1 || y2) {
      var from = parseInt(y1 || y2, 10), to = parseInt(y2 || y1, 10);
      if (!(from >= 1900 && from <= 2100 && to >= 1900 && to <= 2100)) {
        return { error: 'Jahr bitte vierstellig angeben (1900–2100).' };
      }
      if (to < from) return { error: '„bis Jahr“ muss gleich oder größer als „Jahr“ sein.' };
      td.dateRange = [from + '-01-01', to + '-12-31'];
    }
    if (gran) td.granularity = gran;
    query.timeDimensions = [td];
  }

  var f = collectFilters();
  if (f.error) return { error: f.error };
  if (f.filters.length) query.filters = f.filters;

  return { query: query };
}

// Schreibt das Ergebnis ab der aktiven Zelle.
// Mit Namen wird die Abfrage samt Zielposition gespeichert (gleicher Name = überschreiben).
async function run() {
  var b = buildQuery();
  if (b.error) { setStatus(b.error, true); return; }
  var name = $('saveName').value.trim();

  var ok = await withBusy('Abfrage läuft …', async function () {
    var values = await fetchTable(b.query);
    if (!values) return 'Keine Daten' + (name ? ' (nicht gespeichert)' : '');
    var target = await writeTable(values);
    var msg = (values.length - 1) + ' Zeilen geschrieben';
    if (name) {
      await saveQuery(name, b.query, target);
      msg += ' · als „' + name + '“ gespeichert';
    }
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
// Start
// =====================================================================

Office.onReady(function () {
  var id = (window.CONFIG && CONFIG.DEPLOYMENT_ID) || '';
  $('build').textContent = 'Build ' + BUILD + ' · Deployment-ID: ' + (id ? id.slice(0, 10) + '…' : 'FEHLT');

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
