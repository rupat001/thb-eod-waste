// ============================================================
//  EOD Waste / Production — Google Apps Script (API version 2, rev 3)
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

// Run once from the editor: creates both tabs if they're missing.
function setup() { getSheet_(); getHistory_(); }
