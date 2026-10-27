/**
 * PLATFORM PAFUN - Backend (Google Apps Script + Google Sheets)
 *
 * SETUP
 * 1. Create a Google Sheet, then Extensions > Apps Script. Paste this file as Code.gs.
 * 2. Project Settings > set time zone to (GMT+07:00) Bangkok.
 * 3. Run `setup()` once from the editor (authorize when asked). It creates the
 *    "Bookings" sheet and the daily midnight trigger for autoPurgeOldBookings().
 * 4. Deploy > New deployment > type "Web app"
 *      Execute as: Me
 *      Who has access: Anyone
 *    Copy the Web App URL (ends with /exec) into API_URL in index.html.
 * 5. After any code change: Deploy > Manage deployments > Edit > New version.
 */

const SHEET_NAME = 'Bookings';
const TZ = 'Asia/Bangkok';
const MAX_EXTRA_BEDS = 2;      // resort-wide capacity
const RETENTION_DAYS = 30;     // keep history this long
const HORIZON_MONTHS = 6;      // how far ahead bookings are allowed
const API_VERSION = 6;         // frontend checks this to detect an outdated deployment
const MAINT_SHEET = 'Maintenance';   // rooms closed for repair
const MHEADERS = ['maintenance_id', 'house_number', 'start_date', 'end_date', 'note', 'created_at'];
const RATES = { single: 890, double: 990 };   // THB / night: houses 1-7 = single, houses 8-12 = double/special
const INCOME_SHEET = 'Income';                // permanent archive: one row per booking per month
const INCOME_MONTHLY = 'Income_Monthly';      // totals per month (rebuilt every night)
const IHEADERS = ['booking_id', 'month', 'house_number', 'house_type', 'nights', 'rate_per_night', 'income', 'archived_at'];
const HEADERS = ['booking_id', 'customer_name', 'phone_number', 'house_number', 'house_type',
  'check_in_date', 'check_in_time', 'check_out_date', 'check_out_time',
  'extra_bed', 'purpose', 'created_at', 'rate_per_night', 'total_price'];

/* ---------- One-time setup ---------- */
function setup() {
  maintSheet_();
  textFormat_(sheet_()); // creates the "Bookings" sheet if missing; forces text format on key columns
  backfillPrices();      // fill rate/total for bookings saved before prices were stored
  incomeSheet_();
  rebuildIncomeMonthly_();

  // Daily trigger at midnight (replace existing ones to avoid duplicates)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'autoPurgeOldBookings')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoPurgeOldBookings').timeBased().everyDays(1).atHour(0).create();
}

/* ---------- Helpers ---------- */
// Returns the "Bookings" sheet, creating it (headers + text formats) if it is missing.
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script is not bound to a Google Sheet. Open the Sheet > Extensions > Apps Script.');
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    textFormat_(sh);
  } else if (!sh.getRange(1, 13).getValue()) {
    // upgrade an older sheet: add the price columns
    sh.getRange(1, 13, 1, 2).setValues([['rate_per_night', 'total_price']]).setFontWeight('bold');
  }
  return sh;
}

// Returns the "Maintenance" sheet (rooms under repair), creating it if missing.
function maintSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MAINT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MAINT_SHEET);
    sh.getRange(1, 1, 1, MHEADERS.length).setValues([MHEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    ['A', 'C', 'D', 'F'].forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));
  }
  return sh;
}

function readMaint_() {
  const sh = maintSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, MHEADERS.length).getDisplayValues().map(r => {
    const o = {};
    MHEADERS.forEach((h, i) => o[h] = r[i]);
    o.house_number = Number(o.house_number);
    return o;
  });
}

/* ---------- Pricing & income ---------- */
const rateFor_ = n => n <= 7 ? RATES.single : RATES.double;

function addDaysStr_(s, n) {
  const d = new Date(s + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
function nightsBetween_(ci, co) {
  let n = 0;
  for (let d = ci; d < co && n < 1000; d = addDaysStr_(d, 1)) n++;
  return n;
}

// Splits a stay into per-month pieces by NIGHT: [{month:'2026-10', nights, rate, income}, ...]
function allocate_(b) {
  if (!b.house_number) return [];
  const rate = b.rate_per_night || rateFor_(b.house_number);
  const by = {};
  for (let d = b.check_in_date, i = 0; d < b.check_out_date && i < 1000; d = addDaysStr_(d, 1), i++) {
    const m = d.slice(0, 7);
    by[m] = (by[m] || 0) + 1;
  }
  return Object.keys(by).sort().map(m => ({ month: m, nights: by[m], rate: rate, income: by[m] * rate }));
}

// Writes one booking row. Columns A-L are text (keeps leading 0 in phones), M-N are numbers.
function putBooking_(sh, rowIdx, b) {
  sh.getRange(rowIdx, 1, 1, 12).setNumberFormat('@');
  sh.getRange(rowIdx, 13, 1, 2).setNumberFormat('0');
  sh.getRange(rowIdx, 1, 1, HEADERS.length).setValues([HEADERS.map(h => b[h] == null ? '' : b[h])]);
}

// One-time: give bookings saved before prices existed a rate and total (current rates).
function backfillPrices() {
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  const out = rows.map(r => {
    if (r[12] !== '') return [Number(r[12]), r[13] === '' ? '' : Number(r[13])];
    const house = Number(r[3]) || 0;
    if (!house) return ['', ''];
    const rate = rateFor_(house);
    return [rate, nightsBetween_(r[5], r[7]) * rate];
  });
  sh.getRange(2, 13, last - 1, 2).setNumberFormat('0').setValues(out);
}

// Gets a sheet by name or creates it with headers; textCols are formatted as plain text.
function ensureSheet_(name, headers, textCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    textCols.forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));
  }
  return sh;
}
const incomeSheet_ = () => ensureSheet_(INCOME_SHEET, IHEADERS, ['A', 'B', 'H']);

// Copies finished bookings into the permanent Income sheet (skips ones already archived).
function archiveBookings_(list) {
  if (!list.length) return;
  const sh = incomeSheet_(), last = sh.getLastRow();
  const done = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(r => done[r[0]] = true);
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const rows = [];
  list.forEach(b => {
    if (done[b.booking_id]) return;
    allocate_(b).forEach(a => rows.push([b.booking_id, a.month, b.house_number, houseType_(b.house_number),
      a.nights, a.rate, a.income, stamp]));
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, IHEADERS.length).setValues(rows);
}

// Monthly income = archived income (Income sheet) + bookings still in the Bookings sheet.
// Returns [{month, nights_sold, room_income, bookings, occupancy_pct, single_income, double_income}].
// Future months are expected income from confirmed bookings (cancellations correct it).
function computeIncomeMonthly_(bookings) {
  const tot = {};
  const add = (m, house, nights, income) => {
    const o = tot[m] || (tot[m] = { nights: 0, income: 0, bookings: 0, single: 0, double: 0 });
    o.nights += nights; o.income += income; o.bookings++;
    if (house <= 7) o.single += income; else o.double += income;
  };
  const ish = incomeSheet_(), il = ish.getLastRow();
  if (il >= 2) {
    const months = ish.getRange(2, 2, il - 1, 1).getDisplayValues();
    const nums = ish.getRange(2, 3, il - 1, 5).getValues();   // house, type, nights, rate, income
    months.forEach((m, i) => add(m[0], Number(nums[i][0]) || 0, Number(nums[i][2]) || 0, Number(nums[i][4]) || 0));
  }
  bookings.forEach(b => allocate_(b).forEach(a => add(a.month, b.house_number, a.nights, a.income)));
  return Object.keys(tot).sort().map(m => {
    const days = new Date(Date.UTC(Number(m.slice(0, 4)), Number(m.slice(5, 7)), 0)).getUTCDate();
    return { month: m, nights_sold: tot[m].nights, room_income: tot[m].income, bookings: tot[m].bookings,
      occupancy_pct: Math.round(tot[m].nights / (12 * days) * 1000) / 10,
      single_income: tot[m].single, double_income: tot[m].double };
  });
}

const IM_HEADERS = ['month', 'nights_sold', 'room_income', 'bookings', 'occupancy_pct', 'single_income', 'double_income'];
function rebuildIncomeMonthly_() {
  const list = computeIncomeMonthly_(readAll_());
  const sh = ensureSheet_(INCOME_MONTHLY, IM_HEADERS, ['A']);
  sh.clearContents();
  sh.getRange(1, 1, 1, IM_HEADERS.length).setValues([IM_HEADERS]).setFontWeight('bold');
  if (list.length) sh.getRange(2, 1, list.length, IM_HEADERS.length)
    .setValues(list.map(x => IM_HEADERS.map(h => x[h])));
}

// Called after every change so the Income sheets exist and stay current (never blocks a save).
function refreshIncome_() {
  try { rebuildIncomeMonthly_(); } catch (err) { console.error('Income refresh failed: ' + err); }
}

// Run from the editor any time to refresh the monthly income table immediately.
function rebuildIncome() { rebuildIncomeMonthly_(); }

// Plain-text format on whole columns so Sheets never converts phones/dates/times
// (a phone like 0918655134 would otherwise lose its leading 0).
function textFormat_(sh) {
  ['A', 'C', 'F', 'G', 'H', 'I', 'L'].forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));
}

// Run once from the editor to repair phones that already lost their leading 0
// (Thai mobile numbers saved as 9 digits, e.g. 918655134 -> 0918655134).
function fixExistingPhones() {
  const sh = sheet_();
  textFormat_(sh);
  const last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, 3, last - 1, 1);
  const fixed = rng.getDisplayValues().map(r => {
    const v = String(r[0]).trim();
    return [/^\d{9}$/.test(v) ? '0' + v : v];
  });
  rng.setNumberFormat('@').setValues(fixed);
}

function houseType_(n) {
  if (!n) return '';
  if (n >= 1 && n <= 7) return 'Single Bed';
  if (n >= 8 && n <= 11) return 'Double Bed';
  return 'Special Room';
}

function readAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues().map(r => {
    const o = {};
    HEADERS.forEach((h, i) => o[h] = r[i]);
    o.house_number = o.house_number ? Number(o.house_number) : null;
    o.extra_bed = Number(o.extra_bed) || 0;
    o.rate_per_night = o.rate_per_night === '' ? null : Number(o.rate_per_night);
    o.total_price = o.total_price === '' ? null : Number(o.total_price);
    return o;
  });
}

// "YYYY-MM-DD HH:MM" strings compare correctly with plain < and >
const start_ = b => b.check_in_date + ' ' + b.check_in_time;
const end_ = b => b.check_out_date + ' ' + b.check_out_time;
const overlaps_ = (a, b) => start_(a) < end_(b) && end_(a) > start_(b);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dayStr_(offsetDays, offsetMonths) {
  const d = new Date();
  d.setDate(d.getDate() + (offsetDays || 0));
  d.setMonth(d.getMonth() + (offsetMonths || 0));
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/* ---------- API ---------- */
// GET -> health check only. Booking data is NEVER returned without a login token.
function doGet(e) {
  return json_({ ok: true, service: 'PLATFORM PAFUN', auth: 'required', api_version: API_VERSION });
}

/* ---------- Admin authentication ---------- */
// Credentials default to admin / admin1234. To change them without editing code:
// Project Settings > Script properties > add ADMIN_USER and ADMIN_PASS.
const TOKEN_TTL_SEC = 6 * 60 * 60; // 6 hours (CacheService maximum)

function login_(p) {
  const props = PropertiesService.getScriptProperties();
  const user = props.getProperty('ADMIN_USER') || 'admin';
  const pass = props.getProperty('ADMIN_PASS') || 'admin1234';
  if (String(p.username || '') === user && String(p.password || '') === pass) {
    const token = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put('tok_' + token, '1', TOKEN_TTL_SEC);
    return json_({ ok: true, token: token });
  }
  Utilities.sleep(1000); // slow down password guessing
  return json_({ ok: false, error: 'INVALID_LOGIN' });
}

function validToken_(token) {
  return !!token && CacheService.getScriptCache().get('tok_' + token) === '1';
}

// POST (JSON body sent as text/plain to avoid CORS preflight) -> { ok, booking | error }
// Actions: login | logout | list | cancel | update | create (default). All but login need a valid token.
function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Bad request' }); }
  const action = p.action || 'create';
  if (action === 'login') return login_(p);
  if (!validToken_(p.token)) return json_({ ok: false, error: 'UNAUTHORIZED' });
  if (action === 'logout') { CacheService.getScriptCache().remove('tok_' + p.token); return json_({ ok: true }); }
  if (action === 'list') {
    try {
      if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INCOME_MONTHLY)) refreshIncome_();
      const bks = readAll_();
      return json_({ ok: true, bookings: bks, max_extra_beds: MAX_EXTRA_BEDS, api_version: API_VERSION,
        maintenance: readMaint_(), income_monthly: computeIncomeMonthly_(bks) }); }
    catch (err) { return json_({ ok: false, error: String(err) }); }
  }

  // Unknown actions must never fall through to "create"
  if (['create', 'update', 'cancel', 'create_maint', 'update_maint', 'delete_maint'].indexOf(action) === -1) return json_({ ok: false, error: 'Unknown action: ' + action });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // serialise writes so two people can't double-book

    // --- Maintenance (room under repair) actions ---
    if (action === 'delete_maint') {
      const id = String(p.maintenance_id || '').trim();
      const msh = maintSheet_(), ml = msh.getLastRow();
      const ids = ml < 2 ? [] : msh.getRange(2, 1, ml - 1, 1).getDisplayValues();
      const i = ids.findIndex(r => r[0] === id);
      if (i === -1) throw new Error('Repair period ' + id + ' was not found.');
      msh.deleteRow(i + 2);
      SpreadsheetApp.flush();
      return json_({ ok: true, deleted: id });
    }
    if (action === 'create_maint' || action === 'update_maint') return saveMaint_(p, action);

    // --- Cancel: delete the booking row with the given booking_id ---
    if (action === 'cancel') {
      const id = String(p.booking_id || '').trim();
      if (!id) throw new Error('booking_id is required.');
      const sh0 = sheet_(), last0 = sh0.getLastRow();
      const ids = last0 < 2 ? [] : sh0.getRange(2, 1, last0 - 1, 1).getDisplayValues();
      const idx = ids.findIndex(r => r[0] === id);
      if (idx === -1) throw new Error('Booking ' + id + ' was not found (maybe already cancelled).');
      sh0.deleteRow(idx + 2);
      refreshIncome_();
      SpreadsheetApp.flush();
      return json_({ ok: true, cancelled: id });
    }

    const b = {
      customer_name: String(p.customer_name || '').trim(),
      phone_number: String(p.phone_number || '').trim(),
      house_number: p.house_number ? Number(p.house_number) : null,
      check_in_date: p.check_in_date,
      check_in_time: p.check_in_time || '14:00',
      check_out_date: p.check_out_date,
      check_out_time: p.check_out_time || '11:00',
      extra_bed: Number(p.extra_bed) === 1 ? 1 : 0,
      purpose: String(p.purpose || '').trim()
    };

    // --- Validation ---
    if (!b.customer_name) throw new Error('Customer name is required.');
    if (!b.phone_number) throw new Error('Phone number is required.');
    if (b.house_number !== null && !(Number.isInteger(b.house_number) && b.house_number >= 1 && b.house_number <= 12))
      throw new Error('House number must be 1-12.');
    const dateRe = /^\d{4}-\d{2}-\d{2}$/, timeRe = /^\d{2}:\d{2}$/;
    if (!dateRe.test(b.check_in_date || '') || !dateRe.test(b.check_out_date || ''))
      throw new Error('Invalid date format (use YYYY-MM-DD).');
    if (!timeRe.test(b.check_in_time) || !timeRe.test(b.check_out_time))
      throw new Error('Invalid time format (use HH:MM).');
    if (end_(b) <= start_(b)) throw new Error('Check-out must be after check-in.');
    if (b.check_in_date < dayStr_(-RETENTION_DAYS) || b.check_in_date > dayStr_(0, HORIZON_MONTHS))
      throw new Error('Check-in date is outside the allowed range (30 days back to 6 months ahead).');
    if (b.check_out_date > dayStr_(0, HORIZON_MONTHS + 1)) throw new Error('Check-out date is too far ahead.');

    // On edit, ignore the booking being edited so it never clashes with itself
    const editId = action === 'update' ? String(p.booking_id || '').trim() : '';
    const all = readAll_();
    const orig = editId ? all.find(x => x.booking_id === editId) : null;
    if (editId && !orig) throw new Error('Booking ' + editId + ' was not found.');
    const existing = all.filter(x => x.booking_id !== editId);

    // --- Rule 1: double-booking (same house, overlapping range) ---
    if (b.house_number) {
      const clash = existing.find(x => x.house_number === b.house_number && overlaps_(b, x));
      if (clash) throw new Error('House ' + b.house_number + ' is already booked ' +
        start_(clash) + ' to ' + end_(clash) + ' (' + clash.customer_name + ').');
    }

    // --- Rule 1b: house must not be under repair during the stay ---
    if (b.house_number) {
      const mc = readMaint_().find(m => m.house_number === b.house_number &&
        b.check_in_date <= m.end_date && b.check_out_date > m.start_date);
      if (mc) throw new Error('House ' + b.house_number + ' is under repair ' + mc.start_date + ' to ' + mc.end_date + '.');
    }

    // --- Rule 2: extra-bed capacity (max 2 overlapping) ---
    if (b.extra_bed === 1) {
      const used = existing.filter(x => x.extra_bed === 1 && overlaps_(b, x)).length;
      if (used + 1 > MAX_EXTRA_BEDS) throw new Error('No extra beds left for these dates (' + used + '/' + MAX_EXTRA_BEDS + ' in use).');
    }

    b.house_type = houseType_(b.house_number);
    // Price snapshot, calculated on the server (never trusted from the page).
    // When editing the same house the originally saved rate is kept, even if rates changed since.
    if (b.house_number) {
      const keep = orig && orig.house_number === b.house_number && orig.rate_per_night;
      b.rate_per_night = keep ? orig.rate_per_night : rateFor_(b.house_number);
      b.total_price = nightsBetween_(b.check_in_date, b.check_out_date) * b.rate_per_night;
    } else { b.rate_per_night = null; b.total_price = null; }
    const sh = sheet_();

    if (editId) {
      // --- Update in place: keep booking_id and created_at ---
      b.booking_id = editId;
      b.created_at = orig.created_at;
      const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
      const idx = ids.findIndex(r => r[0] === editId);
      if (idx === -1) throw new Error('Booking ' + editId + ' was not found.');
      putBooking_(sh, idx + 2, b);
    } else {
      // --- Create: ID format BK-YYYYMMDD-XXX (XXX = running number for today) ---
      const today = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
      const prefix = 'BK-' + today + '-';
      const seq = existing.filter(x => x.booking_id.indexOf(prefix) === 0)
        .reduce((m, x) => Math.max(m, Number(x.booking_id.slice(-3)) || 0), 0) + 1;
      b.booking_id = prefix + ('00' + seq).slice(-3);
      b.created_at = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
      putBooking_(sh, sh.getLastRow() + 1, b);
    }
    refreshIncome_();
    SpreadsheetApp.flush();
    return json_({ ok: true, booking: b, updated: !!editId });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- Maintenance create / update ---------- */
// A repair period closes the house for the NIGHTS start_date..end_date (inclusive).
function saveMaint_(p, action) {
  const m = { house_number: Number(p.house_number), start_date: p.start_date, end_date: p.end_date,
              note: String(p.note || '').trim() };
  if (!(Number.isInteger(m.house_number) && m.house_number >= 1 && m.house_number <= 12))
    throw new Error('House number must be 1-12.');
  const dateRe = /^\d{4}-\d{2}-\d{2}$/;
  if (!dateRe.test(m.start_date || '') || !dateRe.test(m.end_date || '')) throw new Error('Invalid date format (use YYYY-MM-DD).');
  if (m.end_date < m.start_date) throw new Error('End date must be on or after the start date.');
  if (m.start_date < dayStr_(-RETENTION_DAYS) || m.start_date > dayStr_(0, HORIZON_MONTHS))
    throw new Error('Start date is outside the allowed range (30 days back to 6 months ahead).');
  if (m.end_date > dayStr_(0, HORIZON_MONTHS + 1)) throw new Error('End date is too far ahead.');

  const editId = action === 'update_maint' ? String(p.maintenance_id || '').trim() : '';
  const allM = readMaint_();
  const orig = editId ? allM.find(x => x.maintenance_id === editId) : null;
  if (editId && !orig) throw new Error('Repair period ' + editId + ' was not found.');

  const rc = allM.find(x => x.maintenance_id !== editId && x.house_number === m.house_number &&
    m.start_date <= x.end_date && m.end_date >= x.start_date);
  if (rc) throw new Error('House ' + m.house_number + ' already has a repair period ' + rc.start_date + ' to ' + rc.end_date + '.');

  const bc = readAll_().find(x => x.house_number === m.house_number &&
    x.check_in_date <= m.end_date && x.check_out_date > m.start_date);
  if (bc) throw new Error('House ' + m.house_number + ' has a booking (' + bc.customer_name + ', ' +
    bc.check_in_date + ' to ' + bc.check_out_date + ') in these dates. Cancel or move it first.');

  const sh = maintSheet_();
  if (editId) {
    m.maintenance_id = editId; m.created_at = orig.created_at;
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    const i = ids.findIndex(r => r[0] === editId);
    const r = sh.getRange(i + 2, 1, 1, MHEADERS.length);
    r.setNumberFormat('@'); r.setValues([MHEADERS.map(h => m[h])]);
  } else {
    const prefix = 'MT-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '-';
    const seq = allM.filter(x => x.maintenance_id.indexOf(prefix) === 0)
      .reduce((mx, x) => Math.max(mx, Number(x.maintenance_id.slice(-3)) || 0), 0) + 1;
    m.maintenance_id = prefix + ('00' + seq).slice(-3);
    m.created_at = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const r = sh.getRange(sh.getLastRow() + 1, 1, 1, MHEADERS.length);
    r.setNumberFormat('@'); r.setValues([MHEADERS.map(h => m[h])]);
  }
  SpreadsheetApp.flush();
  return json_({ ok: true, maintenance: m, updated: !!editId });
}

/* ---------- Daily cleanup (time-driven trigger, midnight) ---------- */
function autoPurgeOldBookings() {
  const cutoff0 = dayStr_(-RETENTION_DAYS);
  const msh = maintSheet_(), ml = msh.getLastRow();   // repair periods that ended > 30 days ago
  if (ml >= 2) {
    const ends = msh.getRange(2, 4, ml - 1, 1).getDisplayValues();
    for (let i = ends.length - 1; i >= 0; i--) if (ends[i][0] && ends[i][0] < cutoff0) msh.deleteRow(i + 2);
  }
  const sh = sheet_();
  const cutoff = dayStr_(-RETENTION_DAYS);
  const all = readAll_();   // element i = sheet row i + 2
  archiveBookings_(all.filter(b => b.check_out_date && b.check_out_date < cutoff));   // keep income BEFORE deleting
  for (let i = all.length - 1; i >= 0; i--) {   // bottom-up so row indexes stay valid
    if (all[i].check_out_date && all[i].check_out_date < cutoff) sh.deleteRow(i + 2);
  }
  rebuildIncomeMonthly_();
}