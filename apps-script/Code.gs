/**
 * Monthly Budget API
 * Lives inside the "Monthly Budget" Google Sheet (Extensions > Apps Script).
 * The dashboard on GitHub Pages calls this to read the data and save flags.
 *
 * It never edits Transactions, Budgets or Map. It only writes to three tabs
 * it creates itself: Merchant, Flags and Markers.
 */

// The archived sheet, "Monthly Budget (2022-2024)".
const ARCHIVE_ID = '16XuIX4RY4pScWKaxFFYqtNf3lP9yCLLqyPuNlcXiGfA';

const TAB_MERCHANT = 'Merchant';
const TAB_FLAGS = 'Flags';
const TAB_MARKERS = 'Markers';

// Used only if the category order cannot be read from the Yearly Tracker tab.
const DEFAULT_ORDER = ['Food', 'Dining/Drinks', 'Entertainment', 'Shopping', 'Groceries',
  'Transportation', 'Petrol', 'Home Improv.', 'Gifts', 'Health & Beauty', 'Utilities',
  'Rent/Board', 'Other'];

/* ---------- One-time setup: run this once from the editor ---------- */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureTab_(ss, TAB_MERCHANT, ['Description', 'Merchant', 'Source']);
  ensureTab_(ss, TAB_FLAGS, ['Year', 'Month', 'Category', 'Merchant', 'Reason', 'Note', 'Updated']);
  ensureTab_(ss, TAB_MARKERS, ['Date', 'Label']);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) {
    props.setProperty('KEY', (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
  }
  const data = getData_();
  Logger.log('Setup finished.');
  Logger.log('Expense rows: ' + data.ex.length + ', income rows: ' + data.inc.length +
    ', budget rows: ' + data.bud.length + ', merchants mapped: ' + Object.keys(data.mer).length);
  Logger.log('YOUR SECRET KEY (copy it, keep it private): ' + props.getProperty('KEY'));
}

/** Run this if you ever need to see the key again. */
function showKey() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('KEY'));
}

/** Run this to replace the key. Every device then needs the new one. */
function resetKey() {
  PropertiesService.getScriptProperties()
    .setProperty('KEY', (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
  showKey();
}

function ensureTab_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- Web app entry points ---------- */

function doGet() {
  return out_({ ok: true, app: 'Monthly Budget API' });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { /* handled below */ }
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!key || req.key !== key) return out_({ ok: false, error: 'unauthorised' });
  try {
    if (req.action === 'data') return out_({ ok: true, data: getData_() });
    if (req.action === 'saveFlags') return out_({ ok: true, flags: saveFlags_(req) });
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Reading the sheets ---------- */

function getData_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const cur = readTransactions_(ss, tz);
  let arc = { ex: [], inc: [], sav: [] };
  try {
    arc = readTransactions_(SpreadsheetApp.openById(ARCHIVE_ID), tz);
  } catch (err) {
    // Archive not reachable: carry on with the current sheet only.
  }
  const ex = mergeMax_(arc.ex, cur.ex);
  const inc = mergeMax_(arc.inc, cur.inc);
  const sav = mergeMax_(arc.sav, cur.sav);
  const bud = readBudgets_(ss, tz);

  const people = [];
  // Persona is column 5 of an expense row. The most recent rows decide the two names.
  const tally = {};
  ex.slice(-600).forEach(function (r) { if (r[5]) tally[r[5]] = (tally[r[5]] || 0) + 1; });
  Object.keys(tally).sort(function (a, b) { return tally[b] - tally[a]; })
    .forEach(function (p) { people.push(p); });

  return {
    generated: new Date().toISOString(),
    people: people.slice(0, 2),
    cats: categoryOrder_(ss, bud),
    ex: ex, inc: inc, sav: sav, bud: bud,
    mer: syncMerchants_(ss, ex),
    flags: readFlags_(ss),
    markers: readMarkers_(ss, tz)
  };
}

/**
 * Transactions tab layout (header on row 4, data from row 5):
 *   Expenses: D Date, E Card, F Description, G Amount, H Category, I Persona
 *   Income:   M Date, N Description, O Amount, P Category, Q Persona
 *   Savings:  U Date, V Description, W Amount, X Category, Y Persona
 */
function readTransactions_(ss, tz) {
  const sh = ss.getSheetByName('Transactions');
  const last = sh.getLastRow();
  if (last < 5) return { ex: [], inc: [], sav: [] };
  const v = sh.getRange(5, 1, last - 4, 25).getValues();
  const ex = [], inc = [], sav = [];
  v.forEach(function (r) {
    let d = date_(r[3], tz), a = num_(r[6]);
    // Expenses are negative in the sheet; the app works with positive spend.
    if (d && a !== null && r[5] !== '') ex.push([d, String(r[4]).trim(), String(r[5]).trim(), -a, String(r[7]).trim(), String(r[8]).trim()]);
    d = date_(r[12], tz); a = num_(r[14]);
    if (d && a !== null) inc.push([d, String(r[13]).trim(), a, String(r[15]).trim(), String(r[16]).trim()]);
    d = date_(r[20], tz); a = num_(r[22]);
    if (d && a !== null) sav.push([d, String(r[21]).trim(), a, String(r[23]).trim(), String(r[24]).trim()]);
  });
  return { ex: ex, inc: inc, sav: sav };
}

/** Budgets tab: B Date, C Planned (line name), D Amount, E Section, F Persona, H Notes. */
function readBudgets_(ss, tz) {
  const sh = ss.getSheetByName('Budgets');
  const last = sh.getLastRow();
  if (last < 5) return [];
  const out = [];
  sh.getRange(5, 2, last - 4, 7).getValues().forEach(function (r) {
    const d = date_(r[0], tz), a = num_(r[2]);
    if (!d || !r[1]) return;
    out.push([d, String(r[1]).trim(), a === null ? 0 : a, String(r[3]).trim(), String(r[4]).trim(), String(r[6] || '').trim()]);
  });
  return out;
}

/**
 * The archive and the current sheet overlap. A row that appears in both is one
 * transaction, but two identical rows inside one sheet are two real purchases.
 * So for each identical row we keep whichever sheet has more copies.
 */
function mergeMax_(a, b) {
  const count = function (rows) {
    const m = {};
    rows.forEach(function (r) {
      const k = JSON.stringify(r);
      if (!m[k]) m[k] = { row: r, n: 0 };
      m[k].n++;
    });
    return m;
  };
  const ma = count(a), mb = count(b), out = [];
  Object.keys(ma).forEach(function (k) {
    const n = Math.max(ma[k].n, mb[k] ? mb[k].n : 0);
    for (let i = 0; i < n; i++) out.push(ma[k].row);
  });
  Object.keys(mb).forEach(function (k) {
    if (ma[k]) return;
    for (let i = 0; i < mb[k].n; i++) out.push(mb[k].row);
  });
  out.sort(function (x, y) { return x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0; });
  return out;
}

/** Category order as shown on the Yearly Tracker, then anything else that has a budget. */
function categoryOrder_(ss, bud) {
  let order = [];
  try {
    const sh = ss.getSheets().filter(function (s) { return s.getName().indexOf('Yearly Tracker') === 0; })[0];
    const col = sh.getRange(1, 2, Math.min(sh.getLastRow(), 120), 1).getValues().map(function (r) { return String(r[0]).trim(); });
    const start = col.indexOf('Status', col.indexOf('Expenses', 20));
    for (let i = start + 1; i < col.length && col[i]; i++) order.push(col[i]);
  } catch (err) { order = []; }
  if (order.length < 3) order = DEFAULT_ORDER.slice();
  bud.forEach(function (r) {
    if (r[3] === 'Expenses' && order.indexOf(r[1]) < 0) order.push(r[1]);
  });
  return order;
}

/* ---------- Merchant names ---------- */

const BRANDS = [
  [/WOOLWORTHS|WOOLIES|\bWW METRO/, 'Woolworths'], [/\bCOLES\b/, 'Coles'], [/\bALDI\b/, 'Aldi'],
  [/HARRIS FARM/, 'Harris Farm'], [/COSTCO/, 'Costco'], [/\bIGA\b/, 'IGA'],
  [/AMAZON|AMZN/, 'Amazon'], [/UBER\s*\*?\s*EATS|UBEREATS/, 'Uber Eats'], [/\bUBER\b/, 'Uber'],
  [/DOORDASH/, 'DoorDash'], [/MENULOG/, 'Menulog'], [/NETFLIX/, 'Netflix'], [/SPOTIFY/, 'Spotify'],
  [/DISNEY/, 'Disney+'], [/YOUTUBE/, 'YouTube'], [/ANTHROPIC|CLAUDE/, 'Anthropic'],
  [/GOOGLE/, 'Google'], [/APPLE\.COM|APPLE STORE|ITUNES/, 'Apple'], [/MICROSOFT/, 'Microsoft'],
  [/MCDONALD/, "McDonald's"], [/GUZMAN/, 'Guzman y Gomez'], [/\bKFC\b/, 'KFC'], [/SUBWAY/, 'Subway'],
  [/KMART/, 'Kmart'], [/\bTARGET\b/, 'Target'], [/BIG W\b/, 'Big W'], [/BUNNINGS/, 'Bunnings'],
  [/\bIKEA\b/, 'IKEA'], [/OFFICEWORKS/, 'Officeworks'], [/JB HI-?FI/, 'JB Hi-Fi'],
  [/HARVEY NORMAN/, 'Harvey Norman'], [/\bMYER\b/, 'Myer'], [/DAVID JONES/, 'David Jones'],
  [/UNIQLO/, 'Uniqlo'], [/\bASOS\b/, 'ASOS'], [/COTTON ON/, 'Cotton On'], [/THE ICONIC/, 'The Iconic'],
  [/BABY BUNTING/, 'Baby Bunting'], [/CHEMIST WAREHOUSE/, 'Chemist Warehouse'], [/PRICELINE/, 'Priceline'],
  [/\bMECCA\b/, 'Mecca'], [/FITNESS PASSPORT/, 'Fitness Passport'],
  [/DAN MURPHY/, "Dan Murphy's"], [/\bBWS\b/, 'BWS'],
  [/7-?ELEVEN/, '7-Eleven'], [/AMPOL/, 'Ampol'], [/^BP\b/, 'BP'], [/SPEEDWAY/, 'Speedway'],
  [/\bSHELL\b|COLES EXPRESS/, 'Shell'], [/CALTEX/, 'Caltex'],
  [/TRANSPORT ?FOR ?NSW|\bOPAL\b/, 'Opal'], [/LINKT/, 'Linkt'], [/QANTAS/, 'Qantas'], [/JETSTAR/, 'Jetstar'],
  [/TELSTRA/, 'Telstra'], [/OPTUS/, 'Optus'], [/VODAFONE/, 'Vodafone'], [/\bAGL\b/, 'AGL'],
  [/ORIGIN ENERGY/, 'Origin Energy'], [/SYDNEY WATER/, 'Sydney Water'], [/MORTGAGE/, 'Mortgage'],
  [/\bNRMA\b/, 'NRMA'], [/MEDIBANK/, 'Medibank'], [/\bBUPA\b/, 'Bupa'], [/\bATO\b/, 'ATO']
];

/** Turns a bank description into a merchant name. Known brands first, then a general tidy-up. */
function merchantName_(desc) {
  let s = String(desc).toUpperCase().trim();
  s = s.replace(/^(SQ|ZLR|SMP|LS|TST|PP|PAYPAL|DNH|EZI|SP|IZ|CKO|FS|SUMUP)\s*\*\s*/, '');
  for (let i = 0; i < BRANDS.length; i++) if (BRANDS[i][0].test(s)) return BRANDS[i][1];
  s = s.replace(/\b(PTY|LTD|LIMITED|PTY\.|P\/L|AUSTRALIA|AUS|AU|NSW|VIC|QLD)\b/g, ' ')
    .replace(/\S*\d\S*/g, ' ').replace(/[*_]/g, ' ').replace(/\s+/g, ' ').trim();
  const words = s.split(' ').filter(Boolean);
  let take = words[0] === 'THE' ? 2 : 2;
  if (words[take - 1] === '&') take++;
  const name = words.slice(0, take).join(' ') || String(desc).trim();
  return name.toLowerCase().replace(/(^|[\s&\-])([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); });
}

/**
 * Keeps the Merchant tab complete. Any description that is not there yet gets a
 * row with an automatic name. To correct one, edit column B and set column C to
 * "manual"; this function never changes existing rows.
 */
function syncMerchants_(ss, ex) {
  const sh = ensureTab_(ss, TAB_MERCHANT, ['Description', 'Merchant', 'Source']);
  const map = {};
  const last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (r[0] !== '') map[String(r[0]).trim()] = String(r[1]).trim();
    });
  }
  const add = [];
  ex.forEach(function (r) {
    const d = r[2];
    if (map[d] === undefined) {
      map[d] = merchantName_(d);
      add.push([d, map[d], 'auto']);
    }
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  return map;
}

/* ---------- Flags and markers ---------- */

function readFlags_(ss) {
  const sh = ss.getSheetByName(TAB_FLAGS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(function (r) { return r[0] !== '' && r[3] !== ''; })
    .map(function (r) { return [Number(r[0]), Number(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5])]; });
}

/** Replaces every flag for one category in one month with the list sent by the app. */
function saveFlags_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ensureTab_(ss, TAB_FLAGS, ['Year', 'Month', 'Category', 'Merchant', 'Reason', 'Note', 'Updated']);
    const y = Number(req.y), m = Number(req.m), cat = String(req.cat);
    const last = sh.getLastRow();
    if (last > 1) {
      const rows = sh.getRange(2, 1, last - 1, 3).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (Number(rows[i][0]) === y && Number(rows[i][1]) === m && String(rows[i][2]) === cat) sh.deleteRow(i + 2);
      }
    }
    const items = (req.items || []).filter(function (it) { return it.merchant && it.note; });
    if (items.length) {
      const now = new Date();
      sh.getRange(sh.getLastRow() + 1, 1, items.length, 7).setValues(items.map(function (it) {
        return [y, m, cat, String(it.merchant), String(it.reason || 'One-off'), String(it.note), now];
      }));
    }
    return readFlags_(ss);
  } finally {
    lock.releaseLock();
  }
}

function readMarkers_(ss, tz) {
  const sh = ss.getSheetByName(TAB_MARKERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(function (r) { return [date_(r[0], tz), String(r[1]).trim()]; })
    .filter(function (r) { return r[0] && r[1]; });
}

/* ---------- Small helpers ---------- */

const MONTHS_ = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };

/** Returns yyyy-MM-dd, or '' when the cell is not a date. */
function date_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v || '').match(/(\d{1,2})\D+([A-Za-z]{3})[A-Za-z]*\D+(\d{4})/);
  if (!m || !MONTHS_[m[2].toLowerCase()]) return '';
  return m[3] + '-' + MONTHS_[m[2].toLowerCase()] + '-' + ('0' + m[1]).slice(-2);
}

/** Returns a number, or null when the cell is empty or not numeric. */
function num_(v) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return null;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}
