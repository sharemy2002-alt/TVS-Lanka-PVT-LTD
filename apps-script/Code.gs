/**
 * Product NPS Survey - data API (Google Apps Script)
 * No Google Cloud project, API key or Client ID needed. Access is by a private access code.
 *
 * Install: open the "Product NPS Survey (Responses)" sheet
 *          -> Extensions -> Apps Script -> replace everything with this file -> Save.
 * Then: run setup() once (Run button), open View -> Logs / Execution log, and copy your access code.
 * Then: Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone.
 *
 * The browser sends the access code. If it matches, this script returns CLEANED, ANONYMISED rows
 * (no customer name, phone number or VIN is ever sent). The sheet itself stays private.
 */

const SHEET_RESPONSES = 'Product NPS Survey';
const MAX_FAILS_PER_10_MIN = 30;

// 0-based column positions in "Form responses 1" (A = 0).
// If you ever insert a column in the sheet, update the numbers here.
const C = {
  ts: 0,          // A  Timestamp
  exec: 1,        // B  Executive
  area1: 3,       // D  District (older rows)
  brand: 4,       // E  Brand
  vin: 5,         // F  VIN (used only to remove duplicates, never sent)
  model1: 6,      // G  Brand Name (older rows)
  odo: 7,         // H  Odometer (used only to remove duplicates, never sent)
  mode: 8,        // I  Mode of purchase
  purpose: 15,    // P  Purpose of use
  tenure: 17,     // R  How long using
  recommend: 20,  // U  Will you recommend (Yes/No)
  score: 21,      // V  0-10 score
  likes: 22,      // W  Why recommend
  improve: 23,    // X  Areas of dissatisfaction / improvement
  whyNot: 24,     // Y  Why not recommend
  area2: 25,      // Z  District (newer rows)
  model2: 26,     // AA Model of the bike (newer rows)
  competitor: 27  // AB Competitor model (newer rows)
};
const NUM_COLS = 28;

// Spelling fixes for areas typed by hand. Add more as you spot them.
const AREA_ALIASES = {
  jaffa: 'Jaffna', jafna: 'Jaffna',
  kurunegla: 'Kurunegala', kuunegala: 'Kurunegala',
  rathnapura: 'Ratnapura',
  kky: 'Kaluwanchikudy',
  moneragala: 'Monaragala',
  thambuttegams: 'Thambuttegama'
};

/* ---------- access code ---------- */

/** Run once from the editor. Creates a random access code (if none exists) and prints it in the log. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let code = props.getProperty('ACCESS_CODE');
  if (!code) { code = newCode_(); props.setProperty('ACCESS_CODE', code); }
  console.log('Your access code is: ' + code);
  console.log('Share it only with people who should see the dashboard.');
}

/** Run this to replace the code (for example when someone leaves). The old code stops working at once. */
function changeAccessCode() {
  const code = newCode_();
  PropertiesService.getScriptProperties().setProperty('ACCESS_CODE', code);
  console.log('New access code is: ' + code);
}

function newCode_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';   // no 0/O/1/I to avoid typing mistakes
  let s = '';
  for (let i = 0; i < 10; i++) s += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return s.slice(0, 5) + '-' + s.slice(5);
}

/* ---------- web app entry points ---------- */

function doGet() {
  return ContentService.createTextOutput('Product NPS API is running.');
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad_code' }); }

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS_PER_10_MIN) return json_({ error: 'too_many_attempts' });

  const real = PropertiesService.getScriptProperties().getProperty('ACCESS_CODE');
  const given = String(body.code || '').trim().toUpperCase();
  if (!real || !safeEqual_(given, real.toUpperCase())) {
    cache.put('fails', String(fails + 1), 600);
    Utilities.sleep(1500);                           // slows down guessing
    return json_({ error: 'bad_code' });
  }

  try {
    const parsed = parseResponses_(readValues_());
    return json_({
      rows: parsed.rows,
      meta: Object.assign({}, parsed.meta, { generatedAt: new Date().toISOString() })
    });
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ error: 'server_error' });
  }
}

function safeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/* ---------- reading and cleaning ---------- */

function readValues_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RESPONSES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, NUM_COLS).getValues();
}

/**
 * Pure function (no Google services): turns raw sheet rows into clean dashboard rows.
 * @param {Array<Array>} values raw rows WITHOUT the header row
 */
function parseResponses_(values) {
  const out = [];
  const seen = {};
  const meta = { rowsInSheet: values.length, used: 0, duplicatesRemoved: 0, invalidScore: 0, mismatches: 0 };

  values.forEach(function (r) {
    if (r.join('').trim() === '') return;           // blank row
    const score = parseScore_(r[C.score]);
    const ts = toIso_(r[C.ts]);
    if (score === null || !ts) { meta.invalidScore++; return; }

    const product = productOf_(r);
    const recommend = /^y/i.test(clean_(r[C.recommend])) ? 'Yes' : (/^n/i.test(clean_(r[C.recommend])) ? 'No' : '');
    const flag = (recommend === 'Yes' && score <= 6) || (recommend === 'No' && score >= 9);

    const row = {
      ts: ts,
      executive: clean_(r[C.exec]) || 'Unknown',
      area: areaOf_(r),
      brand: product.brand,
      model: product.model,
      mode: modeOf_(r[C.mode]),
      purpose: clean_(r[C.purpose]) || 'Unknown',
      tenure: clean_(r[C.tenure]) || 'Unknown',
      score: score,
      cat: score >= 9 ? 'P' : (score >= 7 ? 'N' : 'D'),
      recommend: recommend,
      likes: text_(r[C.likes]),
      improve: text_(r[C.improve]),
      whyNot: text_(r[C.whyNot]),
      flag: flag
    };

    // Same VIN + same odometer = the same visit submitted twice. Keep the latest.
    // Executives often type the letter O for zero (and I for 1); real VINs never contain O, I or Q.
    const rawVin = String(r[C.vin] || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    const vin = rawVin.replace(/[OQ]/g, '0').replace(/I/g, '1');
    const odo = String(r[C.odo] || '').replace(/\D/g, '');
    const usable = rawVin.length >= 7 && /\d/.test(rawVin) && rawVin.indexOf('NO') !== 0 && odo !== '';
    if (usable) {
      const key = vin + '|' + odo;
      if (seen[key] !== undefined) { out[seen[key]] = null; meta.duplicatesRemoved++; }
      seen[key] = out.length;
    }
    out.push(row);
  });

  const rows = out.filter(function (x) { return x; });
  meta.used = rows.length;
  meta.mismatches = rows.filter(function (x) { return x.flag; }).length;
  return { rows: rows, meta: meta };
}

function parseScore_(v) {
  const m = String(v == null ? '' : v).match(/\d+/);
  if (!m) return null;
  const n = parseInt(m[0], 10);
  return n >= 0 && n <= 10 ? n : null;
}

/** Form timestamps arrive as Date objects, or as "dd/mm/yyyy hh:mm:ss" text. Returns ISO text. */
function toIso_(v) {
  if (v instanceof Date) return isNaN(v) ? null : v.toISOString();
  const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  const p = function (n) { return ('0' + n).slice(-2); };
  return m[3] + '-' + p(m[2]) + '-' + p(m[1]) + 'T' + p(m[4] || 0) + ':' + p(m[5] || 0) + ':' + p(m[6] || 0) + '+05:30';
}

/** Brand + model, from whichever columns this row used. */
function productOf_(r) {
  const hay = [r[C.brand], r[C.model1], r[C.model2], r[C.competitor]].map(clean_).join(' ').toLowerCase();
  if (/jupit/.test(hay)) return { brand: 'TVS', model: 'TVS Jupiter' };
  if (/raider|rider/.test(hay)) return { brand: 'TVS', model: 'TVS Raider 125' };
  if (/sport|sprt|s\/s/.test(hay)) return { brand: 'TVS', model: 'TVS Sport 110' };
  if (/\bdio\b|honda/.test(hay)) return { brand: 'Honda', model: 'Honda Dio' };
  if (/discover|discower|disecower|bajaj/.test(hay)) return { brand: 'Bajaj', model: 'Bajaj Discover 125' };
  return { brand: 'TVS', model: 'TVS (model not recorded)' };
}

function areaOf_(r) {
  const s = clean_(r[C.area2]) || clean_(r[C.area1]);
  if (!s) return 'Unknown';
  const key = s.toLowerCase().replace(/[^a-z]/g, '');
  if (AREA_ALIASES[key]) return AREA_ALIASES[key];
  return s.toLowerCase().replace(/\b[a-z]/g, function (c) { return c.toUpperCase(); });
}

function modeOf_(v) {
  const s = clean_(v);
  if (/second/i.test(s)) return 'Secondhand';
  if (/new/i.test(s)) return 'New';
  return s || 'Unknown';
}

/* ---------- small helpers ---------- */

function clean_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
function norm_(v) { return clean_(v).toLowerCase(); }

function isBlank_(s) {
  const t = String(s).trim().toLowerCase();
  return !t || /^[\s\-_.=,\/\\]+$/.test(t) ||
    /^(no|none|nil|not|n|np|na|n\/a|nothing|no comment|no comments)$/.test(t);
}

/** Free-text comment: line breaks become " / ", filler like "-" or "No" becomes empty. */
function text_(v) {
  let s = String(v == null ? '' : v).replace(/\s*[\r\n]+\s*/g, ' / ').replace(/\s+/g, ' ').trim();
  s = s.replace(/^(\s*\/\s*)+/, '').replace(/(\s*\/\s*)+$/, '');
  return isBlank_(s) ? '' : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
