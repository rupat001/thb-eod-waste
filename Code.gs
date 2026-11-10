// Google Apps Script for the EOD Waste Sheet.
// Paste into your Google Sheet: Extensions > Apps Script, replace everything, Save,
// then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
//
// One row per Date + Type (Produced / Waste). Saving the same day again updates that row,
// so staff can come back later and add to or correct the numbers.

const SHEET_NAME = 'Submissions';

const ITEMS = [
  'apple_fritter','oreo','vanilla_cream','boston_cream','canadian_maple','raspberry',
  'vanilla_dip','chocolate_dip','maple_dip','honey_dip','honey_cruller','chocolate_cruller',
  'sour_cream','blueberry_sour_cream','chocolate_glazed','double_chocolate','old_fashioned_plain',
  'yeast_rings','yeast_fills',
  'tb_honey_dip','tb_chocolate_glazed','tb_blueberry_sour_cream','tb_birthday_cake',
  'chocolate_chunk','salted_caramel','muffin_chocolate_chip','muffin_blueberry',
  'eggs','hashbrown','sausage','bacon_egg_bites','spinach_egg_bites','bacon','chicken','steak',
  'bagel_plain','bagel_everything',
  'biscuit','english_muffin','croissant','sourdough','artisan_bun'
];
const FIXED = ['Last saved', 'Type', 'Date', 'Name', 'Status', 'Notes', 'Extra items', 'Edit log'];
const HEADER = FIXED.concat(ITEMS);
const COL = { saved: 0, type: 1, date: 2, name: 3, status: 4, notes: 5, extras: 6, log: 7 };

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sh.getRange('C:C').setNumberFormat('@'); // keep dates as plain text like 2026-11-10
  }
  return sh;
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// Returns the 1-based sheet row for this date + type, or 0.
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
  // "Donuts: Pumpkin Spice = 4; Timbits: Apple Cider = 20"
  return String(s || '').split(';').map(function (p) {
    const m = p.trim().match(/^(.+?):\s*(.+?)\s*=\s*(\d+)$/);
    return m ? { group: m[1].trim(), name: m[2].trim(), count: Number(m[3]) } : null;
  }).filter(Boolean);
}

function rowToObj_(r) {
  const counts = {};
  ITEMS.forEach(function (k, i) { const n = Number(r[FIXED.length + i]); if (n > 0) counts[k] = n; });
  const saved = r[COL.saved] instanceof Date ? r[COL.saved].toISOString() : String(r[COL.saved] || '');
  return {
    savedAt: saved, preparedBy: String(r[COL.name] || ''), status: String(r[COL.status] || ''),
    notes: String(r[COL.notes] || ''), extras: parseExtras_(r[COL.extras]), counts: counts
  };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// GET ?date=2026-11-10  ->  { ok, date, produced: {...} | null, waste: {...} | null }
function doGet(e) {
  const date = dateKey_((e && e.parameter && e.parameter.date) || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'bad date' });
  const sh = getSheet_();
  const out = { ok: true, date: date, produced: null, waste: null };
  [['Produced', 'produced'], ['Waste', 'waste']].forEach(function (t) {
    const row = findRow_(sh, date, t[0]);
    if (row) out[t[1]] = rowToObj_(sh.getRange(row, 1, 1, HEADER.length).getValues()[0]);
  });
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const counts = data.counts || {};
    const date = dateKey_(data.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('bad date');
    const type = data.type === 'produced' ? 'Produced' : 'Waste';
    const name = String(data.preparedBy || '').slice(0, 80);
    const now = new Date();
    const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d h:mm a');
    const extras = (Array.isArray(data.extras) ? data.extras : []).slice(0, 30).map(function (x) {
      const n = Math.floor(Number(x.count));
      return String(x.group || '').slice(0, 20) + ': ' + String(x.name || '').slice(0, 40) + ' = ' + (n > 0 && n < 1000 ? n : 0);
    }).join('; ');

    const sh = getSheet_();
    const rowNum = findRow_(sh, date, type);
    let status, log;
    if (rowNum) {
      const old = sh.getRange(rowNum, 1, 1, FIXED.length).getValues()[0];
      const prev = String(old[COL.status] || '');
      // Waste edited after approval or entry goes back to the manager for a second look.
      if (type === 'Waste') status = prev === 'Entered' ? 'Changed after entry' : 'Submitted';
      else status = 'Recorded';
      log = String(old[COL.log] || '') + '; ' + stamp + ' ' + name;
    } else {
      status = type === 'Waste' ? 'Submitted' : 'Recorded';
      log = stamp + ' ' + name;
    }

    const row = [now, type, date, name, status, String(data.notes || '').slice(0, 1000), extras, log.slice(-2000)]
      .concat(ITEMS.map(function (k) {
        const n = Math.floor(Number(counts[k]));
        return n > 0 && n < 1000 ? n : '';
      }));

    if (rowNum) sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    else {
      sh.appendRow(row);
      sh.getRange(sh.getLastRow(), COL.date + 1).setNumberFormat('@').setValue(date);
    }
    return json_({ ok: true, updated: !!rowNum });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Run once from the editor to create the header row.
function setup() { getSheet_(); }
