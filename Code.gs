// ============================================================
//  EOD Waste / Production — Google Apps Script (API version 2, rev 7)
//
//  Paste into your Google Sheet: Extensions > Apps Script, replace everything, Save.
//  PINs live in Project Settings > Script properties (never in this code):
//    STAFF_PIN    – staff entry form
//    MANAGER_PIN  – manager page (also works on the entry form)
//  Update: Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
//  Undo a bad update: same screen, pick the previous version, Deploy.
//
//  Design rules that keep things from breaking:
//   - Numbers are matched to columns by HEADER NAME, never by position.
//     New items (from items.js) get a new column at the end automatically.
//   - Old actions keep working when new ones are added (pages and script
//     can be updated at slightly different times).
//   - Every save is also copied to the History tab, which is never overwritten.
// ============================================================

const API_VERSION = 2;
const START_DATE = '2026-09-01';
const SHEET_NAME = 'Submissions';
const HISTORY_NAME = 'History';

const FIXED = ['Last saved', 'Type', 'Date', 'Name', 'Status', 'Notes', 'Extra items', 'Edit log'];
const COL = { saved: 0, type: 1, date: 2, name: 3, status: 4, notes: 5, extras: 6, log: 7 };
const HISTORY_HEADER = ['Saved at', 'Type', 'Date', 'Name', 'Status after save', 'Notes', 'Extra items', 'Counts (JSON)'];
const KEY_RE = /^[a-z][a-z0-9_]{1,40}$/;
const MAX_COLUMNS = 250;

// ---------- sheet helpers ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet_() {
  let sh = ss_().getSheetByName(SHEET_NAME);
  if (!sh) sh = ss_().insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(FIXED);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, FIXED.length).setFontWeight('bold');
    sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}

function getHistory_() {
  let sh = ss_().getSheetByName(HISTORY_NAME);
  if (!sh) {
    sh = ss_().insertSheet(HISTORY_NAME);
    sh.appendRow(HISTORY_HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HISTORY_HEADER.length).setFontWeight('bold');
    sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}

// Header row -> { names: [...], index: {name: col0} }
function header_(sh) {
  const n = sh.getLastColumn();
  const names = n ? sh.getRange(1, 1, 1, n).getValues()[0].map(String) : [];
  const index = {};
  names.forEach(function (h, i) { if (h && !(h in index)) index[h] = i; });
  return { names: names, index: index };
}

// Make sure every item key has a column (added at the end, in the order given).
function ensureItemColumns_(sh, keys) {
  const h = header_(sh);
  const missing = [];
  keys.forEach(function (k) { if (KEY_RE.test(k) && !(k in h.index) && missing.indexOf(k) < 0) missing.push(k); });
  if (!missing.length) return h;
  if (h.names.length + missing.length > MAX_COLUMNS) throw new Error('too_many_columns');
  if (sh.getMaxColumns() < h.names.length + missing.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), h.names.length + missing.length - sh.getMaxColumns());
  }
  sh.getRange(1, h.names.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  return header_(sh);
}

// Sheets have a fixed number of rows (1,000 by default). Writing past the end fails, so grow first.
function ensureRows_(sh, rowNum) {
  const max = sh.getMaxRows();
  if (rowNum > max) sh.insertRowsAfter(max, Math.max(500, rowNum - max));
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function validDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(d) && d >= START_DATE && d <= today_(); }

function findRow_(sh, date, type) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, FIXED.length).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][COL.type] === type && dateKey_(vals[i][COL.date]) === date) return i + 2;
  }
  return 0;
}

function parseExtras_(s) {
  return String(s || '').split(';').map(function (p) {
    const m = p.trim().match(/^(.+?):\s*(.+?)\s*=\s*(\d+)$/);
    return m ? { group: m[1].trim(), name: m[2].trim(), count: Number(m[3]) } : null;
  }).filter(Boolean);
}

// A row (array) -> object, reading item numbers by header name.
function rowToObj_(r, h) {
  const counts = {};
  h.names.forEach(function (name, i) {
    if (i < FIXED.length || !KEY_RE.test(name)) return;
    const n = Number(r[i]);
    if (n > 0) counts[name] = n;
  });
  const saved = r[COL.saved] instanceof Date ? r[COL.saved].toISOString() : String(r[COL.saved] || '');
  return {
    date: dateKey_(r[COL.date]), type: String(r[COL.type] || ''),
    savedAt: saved, preparedBy: String(r[COL.name] || ''), status: String(r[COL.status] || ''),
    notes: String(r[COL.notes] || ''), extras: parseExtras_(r[COL.extras]),
    log: String(r[COL.log] || ''), counts: counts,
    // Changes on every save (time + length of the edit log), even two saves in the same second.
    version: saved + '|' + String(r[COL.log] || '').length
  };
}

function readRow_(sh, rowNum, h) {
  return rowToObj_(sh.getRange(rowNum, 1, 1, Math.max(h.names.length, FIXED.length)).getValues()[0], h);
}

function json_(o) {
  o.apiVersion = API_VERSION;
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- PINs ----------
// 'manager' | 'staff' | '' (wrong) | 'locked' (too many wrong tries)
function role_(pin) {
  const props = PropertiesService.getScriptProperties();
  const staff = String(props.getProperty('STAFF_PIN') || '');
  const mgr = String(props.getProperty('MANAGER_PIN') || '');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 30) return 'locked';
  pin = String(pin || '');
  if (mgr && pin === mgr) return 'manager';
  if (staff && pin === staff) return 'staff';
  cache.put('pinFails', String(fails + 1), 600); // 30 wrong tries in 10 min -> locked for 10 min
  Utilities.sleep(1200);
  return '';
}

// ---------- entry points ----------
function doGet() {
  return json_({ ok: true, service: 'eod-waste' });
}

// Every request: POST text/plain JSON { v, action, pin, ... }
function doPost(e) {
  let data;
  try { data = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  // Pages from before PINs (no action) get a clear answer instead of writing junk.
  if (!data.action) return json_({ ok: false, error: 'old_client' });

  const role = role_(data.pin);
  if (role === 'locked') return json_({ ok: false, error: 'locked' });
  if (!role) return json_({ ok: false, error: 'bad_pin' });

  try {
    switch (data.action) {
      case 'check': return json_({ ok: true, role: role, startDate: START_DATE });
      case 'load': return json_(load_(data));
      case 'save': return json_(save_(data));
      case 'list':
        if (role !== 'manager') return json_({ ok: false, error: 'manager_only' });
        return json_(list_(data));
      case 'setStatus':
        if (role !== 'manager') return json_({ ok: false, error: 'manager_only' });
        return json_(setStatus_(data));
      case 'addBatch': return json_(addBatch_(data));
      case 'removeBatch': return json_(removeBatch_(data));
      case 'items': return json_(items_());
      case 'setItems':
        if (role !== 'manager') return json_({ ok: false, error: 'manager_only' });
        return json_(setItems_(data));
      default: return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// { date } -> { ok, date, produced, waste }
function load_(data) {
  const date = dateKey_(data.date);
  if (!validDate_(date)) return { ok: false, error: 'bad_date' };
  const sh = getSheet_();
  const h = header_(sh);
  const out = { ok: true, date: date, produced: null, waste: null };
  [['Produced', 'produced'], ['Waste', 'waste']].forEach(function (t) {
    const row = findRow_(sh, date, t[0]);
    if (row) out[t[1]] = readRow_(sh, row, h);
  });
  out.batches = batchesFor_(date);
  return out;
}

// { type, date, preparedBy, notes, counts, extras, itemKeys } -> { ok, updated, record }
function save_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const counts = data.counts || {};
    const date = dateKey_(data.date);
    if (!validDate_(date)) return { ok: false, error: 'bad_date' };
    const type = data.type === 'produced' ? 'Produced' : 'Waste';
    const name = String(data.preparedBy || '').slice(0, 80);
    if (!name) return { ok: false, error: 'no_name' };

    const sh = getSheet_();
    // New items in items.js get their columns here, in list order.
    const keys = (Array.isArray(data.itemKeys) ? data.itemKeys : []).concat(Object.keys(counts)).map(String);
    const h = ensureItemColumns_(sh, keys);

    const now = new Date();
    const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d h:mm a');
    const extras = (Array.isArray(data.extras) ? data.extras : []).slice(0, 30).map(function (x) {
      const n = Math.floor(Number(x.count));
      return String(x.group || '').slice(0, 20) + ': ' + String(x.name || '').slice(0, 40) + ' = ' + (n > 0 && n < 1000 ? n : 0);
    }).join('; ');

    const rowNum = findRow_(sh, date, type);
    const width = h.names.length;

    // Two devices on the same day: refuse to overwrite a save this device never saw.
    // `base` = the version the page loaded ("" = it saw no saved row). Older pages don't send it.
    if (typeof data.base === 'string' && rowNum) {
      const cur = readRow_(sh, rowNum, h);
      if (data.base !== cur.version) return { ok: false, error: 'changed', record: cur };
    }
    let row, status, log;
    if (rowNum) {
      row = sh.getRange(rowNum, 1, 1, width).getValues()[0];
      const prev = String(row[COL.status] || '');
      if (type === 'Waste') status = (prev === 'Entered' || prev === 'Changed after entry') ? 'Changed after entry' : 'Submitted';
      else status = 'Recorded';
      log = String(row[COL.log] || '') + '; ' + stamp + ' ' + name;
    } else {
      row = new Array(width).fill('');
      status = type === 'Waste' ? 'Submitted' : 'Recorded';
      log = stamp + ' ' + name;
    }
    row[COL.saved] = now; row[COL.type] = type; row[COL.date] = date; row[COL.name] = name;
    row[COL.status] = status; row[COL.notes] = String(data.notes || '').slice(0, 1000);
    row[COL.extras] = extras; row[COL.log] = log.slice(-2000);

    // Every item column: the new number, or blank. Items no longer on the form are left as they were.
    const sent = {};
    keys.forEach(function (k) { sent[k] = true; });
    const clean = {};
    h.names.forEach(function (name, i) {
      if (i < FIXED.length || !KEY_RE.test(name)) return;
      const n = Math.floor(Number(counts[name]));
      if (n > 0 && n < 1000) { row[i] = n; clean[name] = n; }
      else if (sent[name]) row[i] = '';
    });

    let written;
    if (rowNum) { sh.getRange(rowNum, 1, 1, width).setValues([row]); written = rowNum; }
    else {
      written = sh.getLastRow() + 1;
      ensureRows_(sh, written);
      sh.getRange(written, COL.date + 1).setNumberFormat('@');
      sh.getRange(written, 1, 1, width).setValues([row]);
    }

    // History: one line per save, never overwritten.
    const hist = getHistory_();
    const hr = hist.getLastRow() + 1;
    ensureRows_(hist, hr);
    hist.getRange(hr, 3).setNumberFormat('@');
    hist.getRange(hr, 1, 1, HISTORY_HEADER.length).setValues([[now, type, date, name, status, row[COL.notes], extras, JSON.stringify(clean)]]);

    SpreadsheetApp.flush();
    return { ok: true, updated: !!rowNum, record: readRow_(sh, written, h) };
  } finally {
    lock.releaseLock();
  }
}

// { days } -> { ok, rows } (newest first)
function list_(data) {
  const days = Math.min(Math.max(Number(data.days) || 21, 1), 400);
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  const h = header_(sh);
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - days);
  let cut = dateKey_(cutoff);
  if (cut < START_DATE) cut = START_DATE;
  const rows = sh.getRange(2, 1, last - 1, Math.max(h.names.length, FIXED.length)).getValues()
    .map(function (r) { return rowToObj_(r, h); })
    .filter(function (r) { return r.date >= cut && (r.type === 'Produced' || r.type === 'Waste'); });
  rows.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : (a.type < b.type ? -1 : 1); });
  return { ok: true, rows: rows };
}

// { date, type: 'waste', status: 'Approved' | 'Submitted', savedAt } -> { ok, record }
function setStatus_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const date = dateKey_(data.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'bad_date' };
    if (data.type !== 'waste') return { ok: false, error: 'only_waste' };
    const status = String(data.status || '');
    if (['Approved', 'Submitted'].indexOf(status) < 0) return { ok: false, error: 'bad_status' };
    const sh = getSheet_();
    const h = header_(sh);
    const rowNum = findRow_(sh, date, 'Waste');
    if (!rowNum) return { ok: false, error: 'not_found' };
    const cur = readRow_(sh, rowNum, h);
    // Don't approve numbers the manager hasn't seen: staff saved again since the page loaded.
    if (data.savedAt && cur.savedAt && data.savedAt !== cur.savedAt) return { ok: false, error: 'changed', record: cur };
    if (cur.status === 'Entered') return { ok: false, error: 'already_entered', record: cur };
    sh.getRange(rowNum, COL.status + 1).setValue(status);
    SpreadsheetApp.flush();
    return { ok: true, record: readRow_(sh, rowNum, h) };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  BATCHES (Produced and Waste)
//  Staff add what they just made / threw out; it is added to the day's row.
//  Every batch is kept in the Batches tab. A wrong batch is removed (its
//  numbers are taken back off the total) and stays listed as removed.
// ============================================================
const BATCHES_NAME = 'Batches';
const BATCHES_HEADER = ['Saved at', 'Date', 'Name', 'Batch', 'Counts (JSON)', 'Extras (JSON)', 'Removed', 'Type'];
const BATCH_RE = /^[A-Za-z0-9_-]{6,40}$/;

function getBatches_() {
  let sh = ss_().getSheetByName(BATCHES_NAME);
  if (sh && sh.getLastColumn() < BATCHES_HEADER.length) {
    sh.getRange(1, 1, 1, BATCHES_HEADER.length).setValues([BATCHES_HEADER]).setFontWeight('bold'); // adds the Type column
  }
  if (!sh) {
    sh = ss_().insertSheet(BATCHES_NAME);
    sh.appendRow(BATCHES_HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, BATCHES_HEADER.length).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
  }
  return sh;
}

function batchRows_() {
  const sh = ss_().getSheetByName(BATCHES_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, BATCHES_HEADER.length).getValues().map(function (r, i) {
    let counts = {}, extras = [];
    try { counts = JSON.parse(r[4] || '{}'); } catch (e) {}
    try { extras = JSON.parse(r[5] || '[]'); } catch (e) {}
    return { row: i + 2, savedAt: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''), date: dateKey_(r[1]),
             name: String(r[2] || ''), id: String(r[3] || ''), counts: counts, extras: extras, removed: String(r[6] || ''),
             type: String(r[7] || '') === 'Waste' ? 'waste' : 'produced' };
  });
}

function batchesFor_(date) {
  return batchRows_().filter(function (b) { return b.date === date; })
    .map(function (b) { return { id: b.id, type: b.type, savedAt: b.savedAt, name: b.name, counts: b.counts, extras: b.extras, removed: b.removed }; });
}

function cleanExtras_(list) {
  return (Array.isArray(list) ? list : []).slice(0, 30).map(function (x) {
    return { group: String(x.group || '').slice(0, 20), name: String(x.name || '').trim().slice(0, 40), count: Math.floor(Number(x.count)) };
  }).filter(function (x) { return x.group && x.name && x.count > 0 && x.count < 1000; });
}

// Add (sign 1) or take off (sign -1) a batch's numbers on the day's Produced or Waste row.
function applyToRow_(type, date, counts, extras, sign, name, note, itemKeys) {
  const sh = getSheet_();
  const h = ensureItemColumns_(sh, (itemKeys || []).concat(Object.keys(counts)));
  const width = h.names.length;
  const rowNum = findRow_(sh, date, type);
  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d h:mm a');
  let row;
  if (rowNum) row = sh.getRange(rowNum, 1, 1, width).getValues()[0];
  else { row = new Array(width).fill(''); row[COL.log] = ''; }
  h.names.forEach(function (k, i) {
    if (i < FIXED.length || !(k in counts)) return;
    const n = Math.max(0, (Number(row[i]) || 0) + sign * counts[k]);
    row[i] = n > 0 ? n : '';
  });
  // Extras ("Group: Name = n; ...") are added up by group + name.
  const ex = parseExtras_(row[COL.extras]);
  extras.forEach(function (x) {
    const f = ex.filter(function (e) { return e.group === x.group && e.name.toLowerCase() === x.name.toLowerCase(); })[0];
    if (f) f.count = Math.max(0, f.count + sign * x.count);
    else if (sign > 0) ex.push({ group: x.group, name: x.name, count: x.count });
  });
  row[COL.extras] = ex.filter(function (e) { return e.count > 0; }).map(function (e) { return e.group + ': ' + e.name + ' = ' + e.count; }).join('; ');
  let status = 'Recorded';
  if (type === 'Waste') {
    const prev = String(row[COL.status] || '');
    status = (prev === 'Entered' || prev === 'Changed after entry') ? 'Changed after entry' : 'Submitted';
  }
  row[COL.saved] = now; row[COL.type] = type; row[COL.date] = date; row[COL.name] = name; row[COL.status] = status;
  if (note) row[COL.notes] = (String(row[COL.notes] || '') ? String(row[COL.notes]) + ' | ' : '') + name + ': ' + note;
  row[COL.notes] = String(row[COL.notes] || '').slice(-1000);
  const pieces = Object.keys(counts).reduce(function (t, k) { return t + counts[k]; }, 0) + extras.reduce(function (t, x) { return t + x.count; }, 0);
  row[COL.log] = ((String(row[COL.log] || '') ? String(row[COL.log]) + '; ' : '') + stamp + ' ' + name + (sign > 0 ? (pieces < 0 ? ' ' : ' +') : ' removed ') + pieces).slice(-2000);
  let written;
  if (rowNum) { sh.getRange(rowNum, 1, 1, width).setValues([row]); written = rowNum; }
  else {
    written = sh.getLastRow() + 1;
    ensureRows_(sh, written);
    sh.getRange(written, COL.date + 1).setNumberFormat('@');
    sh.getRange(written, 1, 1, width).setValues([row]);
  }
  const hist = getHistory_();
  const hr = hist.getLastRow() + 1;
  ensureRows_(hist, hr);
  hist.getRange(hr, 3).setNumberFormat('@');
  hist.getRange(hr, 1, 1, HISTORY_HEADER.length).setValues([[now, type + (sign > 0 ? ' batch' : ' batch removed'), date, name, status, note || '', JSON.stringify(extras), JSON.stringify(counts)]]);
  return readRow_(sh, written, h);
}

// { type: 'produced'|'waste', date, preparedBy, batchId, counts, extras, notes, itemKeys } -> { ok, record, batches }
function addBatch_(data) {
  const type = data.type === 'waste' ? 'Waste' : 'Produced';
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const date = dateKey_(data.date);
    if (!validDate_(date)) return { ok: false, error: 'bad_date' };
    const name = String(data.preparedBy || '').trim().slice(0, 80);
    if (!name) return { ok: false, error: 'no_name' };
    const id = String(data.batchId || '');
    if (!BATCH_RE.test(id)) return { ok: false, error: 'bad_batch' };
    const sh = getSheet_();
    // Sent twice (slow Wi-Fi, tapped again): add it only once.
    if (batchRows_().some(function (b) { return b.id === id; })) {
      const r = findRow_(sh, date, type);
      return { ok: true, duplicate: true, record: r ? readRow_(sh, r, header_(sh)) : null, batches: batchesFor_(date) };
    }
    const counts = {};
    Object.keys(data.counts || {}).forEach(function (k) {
      const n = Math.floor(Number(data.counts[k]));
      // Minus numbers correct an earlier batch (e.g. -2 when 2 too many were entered).
      if (KEY_RE.test(k) && n !== 0 && n > -1000 && n < 1000) counts[k] = n;
    });
    const extras = cleanExtras_(data.extras);
    if (!Object.keys(counts).length && !extras.length) return { ok: false, error: 'empty' };
    // A correction can't take an item below 0.
    const neg = Object.keys(counts).filter(function (k) { return counts[k] < 0; });
    if (neg.length) {
      const pr = findRow_(sh, date, type);
      const cur = pr ? readRow_(sh, pr, header_(sh)).counts : {};
      const bad = neg.filter(function (k) { return (cur[k] || 0) + counts[k] < 0; });
      if (bad.length) return { ok: false, error: 'below_zero', items: bad, totals: cur };
    }
    const note = String(data.notes || '').trim().slice(0, 300);
    const record = applyToRow_(type, date, counts, extras, 1, name, note, Array.isArray(data.itemKeys) ? data.itemKeys.map(String) : []);
    const b = getBatches_();
    const br = b.getLastRow() + 1;
    ensureRows_(b, br);
    b.getRange(br, 2, 1, 3).setNumberFormat('@');
    b.getRange(br, 1, 1, BATCHES_HEADER.length).setValues([[new Date(), date, name, id, JSON.stringify(counts), JSON.stringify(extras), '', type]]);
    SpreadsheetApp.flush();
    return { ok: true, record: record, batches: batchesFor_(date) };
  } finally {
    lock.releaseLock();
  }
}

// { date, batchId, preparedBy } -> { ok, record, batches }
function removeBatch_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const date = dateKey_(data.date);
    const name = String(data.preparedBy || '').trim().slice(0, 80);
    if (!name) return { ok: false, error: 'no_name' };
    const b = batchRows_().filter(function (x) { return x.id === String(data.batchId || '') && x.date === date; })[0];
    if (!b) return { ok: false, error: 'not_found' };
    const sh = getSheet_();
    const btype = b.type === 'waste' ? 'Waste' : 'Produced';
    if (b.removed) {
      const r = findRow_(sh, date, btype);
      return { ok: true, record: r ? readRow_(sh, r, header_(sh)) : null, batches: batchesFor_(date) };
    }
    const record = applyToRow_(btype, date, b.counts, cleanExtras_(b.extras), -1, name, '', []);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d h:mm a');
    getBatches_().getRange(b.row, 7).setValue('Removed ' + stamp + ' by ' + name);
    SpreadsheetApp.flush();
    return { ok: true, record: record, batches: batchesFor_(date) };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  ITEM LIST (managed from the Manager page, stored in the Items tab)
//  Columns: Order, Category, Key, Label, Note, Active, ClearView row
//  - Key never changes once saved (it names the Submissions column).
//  - Items are hidden, never deleted, so old numbers keep their history.
//  - Until the Items tab exists, pages use their built-in items.js list.
// ============================================================
const ITEMS_NAME = 'Items';
const ITEMS_HEADER = ['Order', 'Category', 'Key', 'Label', 'Note', 'Active', 'ClearView row'];
const MAX_ITEMS = 300;

function itemsVersion_() {
  return String(PropertiesService.getScriptProperties().getProperty('ITEMS_VERSION') || '0');
}

// -> [{ group, key, label, note, active, clearview }] in order, or null if no Items tab yet
function readItems_() {
  const sh = ss_().getSheetByName(ITEMS_NAME);
  if (!sh || sh.getLastRow() < 2) return null;
  return sh.getRange(2, 1, sh.getLastRow() - 1, ITEMS_HEADER.length).getValues()
    .filter(function (r) { return KEY_RE.test(String(r[2])); })
    .sort(function (a, b) { return Number(a[0]) - Number(b[0]); })
    .map(function (r) {
      return { group: String(r[1]), key: String(r[2]), label: String(r[3]), note: String(r[4] || ''),
               active: !(r[5] === false || String(r[5]).toUpperCase() === 'FALSE'), clearview: String(r[6] || '') };
    });
}

function items_() {
  return { ok: true, items: readItems_(), version: itemsVersion_() };
}

// { items, version } -> { ok, items, version }   (manager only)
function setItems_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Two managers editing at once: refuse to overwrite a list this page never saw.
    if (String(data.version) !== itemsVersion_()) return { ok: false, error: 'items_changed', items: readItems_(), version: itemsVersion_() };
    const list = Array.isArray(data.items) ? data.items : [];
    if (!list.length || list.length > MAX_ITEMS) return { ok: false, error: 'bad_items', message: 'The list must have 1 to ' + MAX_ITEMS + ' items.' };
    const seen = {};
    const clean = [];
    for (let i = 0; i < list.length; i++) {
      const it = list[i] || {};
      const key = String(it.key || ''), group = String(it.group || '').trim(), label = String(it.label || '').trim();
      if (!KEY_RE.test(key)) return { ok: false, error: 'bad_items', message: 'Bad item key: ' + key };
      if (seen[key]) return { ok: false, error: 'bad_items', message: 'Duplicate item: ' + label };
      if (!group || group.length > 20) return { ok: false, error: 'bad_items', message: 'Category names must be 1–20 characters.' };
      if (!label || label.length > 40) return { ok: false, error: 'bad_items', message: 'Item names must be 1–40 characters.' };
      seen[key] = true;
      clean.push([i + 1, group, key, label, String(it.note || '').trim().slice(0, 30), it.active !== false, String(it.clearview || '').trim().slice(0, 80)]);
    }
    // Saved items can be hidden but never removed, so their history stays readable.
    const before = readItems_() || [];
    const dropped = before.filter(function (b) { return !seen[b.key]; });
    if (dropped.length) return { ok: false, error: 'bad_items', message: 'Items can be hidden but not deleted: ' + dropped.map(function (d) { return d.label; }).join(', ') };

    let sh = ss_().getSheetByName(ITEMS_NAME);
    if (!sh) sh = ss_().insertSheet(ITEMS_NAME);
    sh.clearContents();
    ensureRows_(sh, clean.length + 1);
    sh.getRange(1, 1, clean.length + 1, ITEMS_HEADER.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, ITEMS_HEADER.length).setValues([ITEMS_HEADER]).setFontWeight('bold');
    sh.getRange(2, 1, clean.length, ITEMS_HEADER.length).setValues(clean.map(function (r) { return [String(r[0]), r[1], r[2], r[3], r[4], r[5] ? 'TRUE' : 'FALSE', r[6]]; }));
    sh.setFrozenRows(1);
    // New items get their Submissions column now, in list order.
    ensureItemColumns_(getSheet_(), clean.map(function (r) { return r[2]; }));
    const v = String(Number(itemsVersion_()) + 1);
    PropertiesService.getScriptProperties().setProperty('ITEMS_VERSION', v);
    SpreadsheetApp.flush();
    return { ok: true, items: readItems_(), version: v };
  } finally {
    lock.releaseLock();
  }
}

// Run once from the editor: creates both tabs if they're missing.
function setup() { getSheet_(); getHistory_(); }

// ============================================================
//  MORNING CHECK (runs on Google's schedule — no Claude needed)
//
//  Every morning around 7:00 AM it emails a short summary:
//   - yesterday's Produced / Waste missing?
//   - waste days waiting for approval
//   - days changed after they were entered in ClearView
//   - approved days not entered yet
//  Turn on once: pick setupMorningCheck in the toolbar, Run, Allow.
//  It sends one email right away so you know it works.
//  Recipients: the Sheet owner, or Script property ALERT_EMAILS
//  (comma-separated) to send to other people instead.
//  Turn off: run stopMorningCheck.
// ============================================================
const CHECK_HOUR = 7;
const MANAGER_URL = 'https://rupat001.github.io/thb-eod-waste/manager.html';
const FORM_URL = 'https://rupat001.github.io/thb-eod-waste/';

function setupMorningCheck() {
  stopMorningCheck();
  ScriptApp.newTrigger('morningCheck').timeBased()
    .everyDays(1).atHour(CHECK_HOUR).nearMinute(0)
    .inTimezone(Session.getScriptTimeZone()).create();
  morningCheck();
}

function stopMorningCheck() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'morningCheck') ScriptApp.deleteTrigger(t);
  });
}

function fmtDay_(d) {
  const p = d.split('-').map(Number);
  return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2], 12), Session.getScriptTimeZone(), 'EEE MMM d');
}

// Pure summary from rows (kept separate so it can be tested).
function buildMorningSummary_(rows, today) {
  const tz = Session.getScriptTimeZone();
  const y = new Date(today.getTime()); y.setDate(y.getDate() - 1);
  const yesterday = Utilities.formatDate(y, tz, 'yyyy-MM-dd');
  const f = new Date(today.getTime()); f.setDate(f.getDate() - 14);
  const from = Utilities.formatDate(f, tz, 'yyyy-MM-dd');

  const byKey = {};
  rows.forEach(function (r) { byKey[r.date + '|' + r.type] = r; });
  const missing = [];
  if (yesterday >= START_DATE) {
    if (!byKey[yesterday + '|Produced']) missing.push('Produced');
    if (!byKey[yesterday + '|Waste']) missing.push('Waste');
  }
  const waste = rows.filter(function (r) { return r.type === 'Waste' && r.date >= from; })
    .sort(function (a, b) { return a.date < b.date ? -1 : 1; });
  const waiting = waste.filter(function (r) { return r.status === 'Submitted'; });
  const changed = waste.filter(function (r) { return r.status === 'Changed after entry'; });
  const approved = waste.filter(function (r) { return r.status === 'Approved'; });
  const notes = ['Produced', 'Waste'].map(function (t) { const r = byKey[yesterday + '|' + t]; return r && r.notes ? t + ': ' + r.notes : ''; }).filter(String);

  const day = fmtDay_(yesterday);
  const problems = [];
  if (missing.length) problems.push('Missing: ' + missing.join(' and ') + ' for ' + day);
  if (waiting.length) problems.push(waiting.length + ' waiting for approval');
  if (changed.length) problems.push(changed.length + ' changed after entry');
  if (approved.length) problems.push(approved.length + ' ready to enter');
  const subject = 'EOD Waste – ' + (problems.length ? problems.join(' · ') : 'All good for ' + day);

  const lines = [];
  lines.push(missing.length ? 'MISSING for ' + day + ': ' + missing.join(' and ') + '. Staff can still enter it on the form.' : 'Yesterday (' + day + '): Produced and Waste both entered.');
  const list = function (title, arr, tail) {
    if (!arr.length) return;
    lines.push('');
    lines.push(title + ':');
    arr.forEach(function (r) { lines.push('  • ' + fmtDay_(r.date) + ' – ' + (r.preparedBy || 'no name')); });
    if (tail) lines.push('  ' + tail);
  };
  list('Waiting for your approval', waiting, 'Approve on the Manager page.');
  list('Changed after entry (ClearView needs fixing)', changed, 'Approve again, then ask Claude to "enter approved waste" or fix ClearView by hand.');
  list('Approved, not entered in ClearView yet', approved, 'Ask Claude to "enter approved waste", or enter them by hand.');
  if (notes.length) { lines.push(''); lines.push('Notes from yesterday:'); notes.forEach(function (n) { lines.push('  • ' + n); }); }
  lines.push('');
  lines.push('Manager page: ' + MANAGER_URL);
  lines.push('Staff form: ' + FORM_URL);
  return { subject: subject, body: lines.join('\n'), needsAction: problems.length > 0 };
}

function morningCheck() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const h = header_(sh);
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, Math.max(h.names.length, FIXED.length)).getValues()
    .map(function (r) { return rowToObj_(r, h); });
  const s = buildMorningSummary_(rows, new Date());
  const to = String(PropertiesService.getScriptProperties().getProperty('ALERT_EMAILS') || '').trim()
    || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to: to, subject: s.subject, body: s.body, name: 'EOD Waste' });
}
