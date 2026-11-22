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
const TAB_DRAFT = 'Plan Draft';
const TAB_LOG = 'Change Log';
const BACKUPS_KEPT = 5;

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
  Logger.log('Expense rows: ' + data.ex.length + ' (from ' + (data.ex[0] || [''])[0] + ' to ' +
    (data.ex[data.ex.length - 1] || [''])[0] + '), income rows: ' + data.inc.length +
    ', budget rows: ' + data.bud.length + ', merchants mapped: ' + Object.keys(data.mer).length);
  if (data.archive.error) Logger.log('ARCHIVE NOT READ: ' + data.archive.error);
  else Logger.log('Archive read: ' + data.archive.expenses + ' expense rows found, older ones added.');
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
    if (req.action === 'reallocate') return out_({ ok: true, result: reallocate_(req) });
    if (req.action === 'saveDraft') return out_({ ok: true, drafts: saveDraft_(req) });
    if (req.action === 'publish') return out_({ ok: true, result: publish_(req) });
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
  const arc = readArchive_(tz);
  // The archive only fills in history from before the current sheet starts, so
  // the overlap (2024) is never counted twice, even if a row was edited later.
  const ex = before_(arc.ex, cur.ex).concat(cur.ex);
  const inc = before_(arc.inc, cur.inc).concat(cur.inc);
  const sav = before_(arc.sav, cur.sav).concat(cur.sav);
  const bud = readBudgets_(ss, tz);

  const people = [];
  // Persona is column 5 of an expense row. The most recent rows decide the two names.
  const tally = {};
  ex.slice(-600).forEach(function (r) { if (r[5]) tally[r[5]] = (tally[r[5]] || 0) + 1; });
  Object.keys(tally).sort(function (a, b) { return tally[b] - tally[a]; })
    .forEach(function (p) { people.push(p); });

  return {
    generated: new Date().toISOString(),
    archive: { expenses: arc.ex.length, error: arc.error || '' },
    people: people.slice(0, 2),
    cats: categoryOrder_(ss, bud),
    ex: ex, inc: inc, sav: sav, bud: bud,
    mer: syncMerchants_(ss, ex),
    flags: readFlags_(ss),
    markers: readMarkers_(ss, tz),
    drafts: readDrafts_(ss)
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

function readArchive_(tz) {
  try {
    const out = readTransactions_(SpreadsheetApp.openById(ARCHIVE_ID), tz);
    out.error = '';
    return out;
  } catch (err) {
    return { ex: [], inc: [], sav: [], error: String(err) };
  }
}

/** Archive rows dated before the first row of the current sheet. Rows are sorted by date. */
function before_(archive, current) {
  if (!current.length) return archive.slice();
  let first = current[0][0];
  current.forEach(function (r) { if (r[0] < first) first = r[0]; });
  return archive.filter(function (r) { return r[0] < first; });
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

/* ---------- Writing budgets: reallocate, drafts, publish ---------- */

/**
 * Budgets layout: B Date, C Planned (line), D Amount, E Section, F Persona, G blank, H Notes.
 * Rows are matched exactly on month + line + section + persona.
 */
function budgetIndex_(sh) {
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  const idx = {};
  if (last < 5) return { idx: idx, last: last };
  const v = sh.getRange(5, 2, last - 4, 5).getValues();
  v.forEach(function (r, i) {
    if (!(r[0] instanceof Date) || !r[1]) return;
    // Read the month in the sheet's own time zone; the script's zone can differ.
    const k = Utilities.formatDate(r[0], tz, 'yyyy-M') + '|' + String(r[1]).trim() + '|' +
      String(r[3]).trim() + '|' + String(r[4]).trim();
    if (!idx[k]) idx[k] = i + 5;
  });
  return { idx: idx, last: last };
}

function log_(ss, action, detail) {
  const sh = ensureTab_(ss, TAB_LOG, ['When', 'Action', 'Detail']);
  sh.appendRow([new Date(), action, detail]);
}

/** Moves money from the December pot of a drawdown line into another month, per person. */
function reallocate_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Budgets');
    const y = Number(req.y), m = Number(req.m), line = String(req.line), sec = String(req.section);
    if (m === 12) throw new Error('December is the pot itself; pick another month.');
    const bi = budgetIndex_(sh);
    const changes = [];
    (req.items || []).forEach(function (it) {
      const amt = Math.round(Number(it.amount) * 100) / 100;
      if (!amt) return;
      const per = String(it.person);
      const rm = bi.idx[y + '-' + m + '|' + line + '|' + sec + '|' + per];
      const rd = bi.idx[y + '-12|' + line + '|' + sec + '|' + per];
      if (!rm || !rd) throw new Error('No budget row found for ' + line + ' (' + per + ') in ' + (rm ? 'December' : 'that month') + ' ' + y + '.');
      [rm, rd].forEach(function (r) {
        if (sh.getRange(r, 4).getFormula()) throw new Error('Row ' + r + ' holds a formula, so it was left alone.');
      });
      const cm = sh.getRange(rm, 4), cd = sh.getRange(rd, 4);
      const om = Number(cm.getValue()) || 0, od = Number(cd.getValue()) || 0;
      cm.setValue(om - amt);
      cd.setValue(od + amt);
      changes.push(per + ': ' + y + '-' + m + ' ' + om + ' -> ' + (om - amt) + ', Dec ' + od + ' -> ' + (od + amt));
      if (req.note) {
        const ch = sh.getRange(rm, 8), old = String(ch.getValue() || '').trim();
        if (old.indexOf(req.note) < 0) ch.setValue(old ? old + ', ' + req.note : req.note);
      }
    });
    if (!changes.length) throw new Error('Nothing to move.');
    log_(ss, 'Reallocate ' + line, changes.join(' | ') + (req.note ? ' | note: ' + req.note : ''));
    return changes;
  } finally {
    lock.releaseLock();
  }
}

function readDrafts_(ss) {
  const sh = ss.getSheetByName(TAB_DRAFT);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
    .filter(function (r) { return r[0] !== '' && r[3] !== ''; })
    .map(function (r) {
      return [Number(r[0]), String(r[1]), Number(r[2]), String(r[3]), String(r[4]), Number(r[5]) || 0, String(r[6] || ''),
        r[7] instanceof Date ? r[7].toISOString() : String(r[7] || '')];
    });
}

/** Replaces one person's draft for one year. Rows: [month, line, section, amount, note]. */
function saveDraft_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ensureTab_(ss, TAB_DRAFT, ['Year', 'Persona', 'Month', 'Line', 'Section', 'Amount', 'Note', 'Saved']);
    const y = Number(req.y), per = String(req.person);
    const last = sh.getLastRow();
    if (last > 1) {
      const keep = sh.getRange(2, 1, last - 1, 8).getValues()
        .filter(function (r) { return !(Number(r[0]) === y && String(r[1]) === per); });
      sh.getRange(2, 1, last - 1, 8).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, 8).setValues(keep);
    }
    const now = new Date();
    const rows = (req.rows || []).map(function (r) {
      return [y, per, Number(r[0]), String(r[1]), String(r[2]), Number(r[3]) || 0, String(r[4] || ''), now];
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
    return readDrafts_(ss);
  } finally {
    lock.releaseLock();
  }
}

/** Writes saved drafts into Budgets: updates matching rows, appends missing ones. Backs up first. */
function publish_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Budgets');
    const y = Number(req.y), people = (req.persons || []).map(String);
    const drafts = readDrafts_(ss).filter(function (r) { return r[0] === y && people.indexOf(r[1]) >= 0; });
    if (!drafts.length) throw new Error('No saved draft for ' + y + '.');

    // Backup, keeping only the most recent few.
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmm');
    const bak = sh.copyTo(ss).setName('Budgets backup ' + stamp);
    ss.setActiveSheet(bak); ss.moveActiveSheet(ss.getNumSheets());
    const olds = ss.getSheets().filter(function (x) { return x.getName().indexOf('Budgets backup ') === 0; })
      .sort(function (a, b) { return a.getName() < b.getName() ? 1 : -1; });
    olds.slice(BACKUPS_KEPT).forEach(function (x) { ss.deleteSheet(x); });

    const bi = budgetIndex_(sh);
    const n = Math.max(bi.last - 4, 0);
    const dCol = n ? sh.getRange(5, 4, n, 1) : null, hCol = n ? sh.getRange(5, 8, n, 1) : null;
    const dv = n ? dCol.getValues() : [], df = n ? dCol.getFormulas() : [], hv = n ? hCol.getValues() : [];
    let updated = 0, skipped = 0;
    const add = [];
    drafts.forEach(function (r) {
      const k = y + '-' + r[2] + '|' + r[3] + '|' + r[4] + '|' + r[1];
      const row = bi.idx[k];
      if (row) {
        const i = row - 5;
        if (df[i][0]) { skipped++; return; }
        dv[i][0] = r[5];
        if (r[6] && !String(hv[i][0] || '').trim()) hv[i][0] = r[6];
        updated++;
      } else {
        add.push(r);
      }
    });
    if (n) {
      dCol.setValues(dv.map(function (x, i) { return [df[i][0] || x[0]]; }));
      hCol.setValues(hv);
    }
    if (add.length) {
      // Same layout as existing years: one block per person, month by month.
      add.sort(function (a, b) {
        return people.indexOf(a[1]) - people.indexOf(b[1]) || a[2] - b[2];
      });
      const start = sh.getLastRow() + 1;
      const tz = ss.getSpreadsheetTimeZone();
      // Midnight on the 1st in the sheet's time zone, so the trackers' date ranges pick the rows up.
      const rows = add.map(function (r) {
        return [Utilities.parseDate(y + '-' + ('0' + r[2]).slice(-2) + '-01', tz, 'yyyy-MM-dd'), r[3], r[5], r[4], r[1], '', r[6]];
      });
      sh.getRange(start, 2, rows.length, 7).setValues(rows);
      if (start > 5) sh.getRange(start - 1, 2, 1, 7).copyTo(sh.getRange(start, 2, rows.length, 7), { formatOnly: true });
    }
    log_(ss, 'Publish ' + y, people.join(' + ') + ': ' + updated + ' updated, ' + add.length + ' added' +
      (skipped ? ', ' + skipped + ' formula cells left alone' : '') + '. Backup: ' + bak.getName());
    return { updated: updated, added: add.length, skipped: skipped, backup: bak.getName() };
  } finally {
    lock.releaseLock();
  }
}
