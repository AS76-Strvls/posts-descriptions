/**
 * 76ο Σύστημα — Σημειώσεις / Περιγραφές Δημοσιεύσεων (backend)
 * Google Apps Script Web App. Πηγή αλήθειας: το συνδεδεμένο Google Sheet (φύλλο «Πρώτη Ύλη»).
 *
 *  • «Πρώτη Ύλη» (Sheet) — η πρώτη ύλη κάθε τμήματος ανά Σάββατο, για να φτιάχνεται το description.
 *  • «Τελευταία δράσεις» — διαβάζονται ΖΩΝΤΑΝΑ & ΔΗΜΟΣΙΑ από την ιστοσελίδα 76.life
 *    (WordPress REST API, κατηγορία «ΔΡΑΣΕΙΣ») — ΧΩΡΙΣ σύνδεση/token.
 *
 * ΜΟΝΤΕΛΟ πρώτης ύλης: append (κάθε καταχώρηση = μία γραμμή με μοναδικό ID) → ταυτόχρονες
 * εγγραφές τμημάτων δεν σβήνουν η μία την άλλη.
 *
 * API (όλα μέσω doGet ώστε να δουλεύει JSONP από GitHub Pages χωρίς CORS):
 *   ?action=list                 -> { ok, raw:[...] }
 *   ?action=add&<πεδία>          -> { ok, id }
 *   ?action=update&ID=..&<πεδία> -> { ok }
 *   ?action=delete&ID=..         -> { ok }
 *   ?action=web                  -> { ok, items:[{id,date,link,title,excerpt}] }  (από 76.life)
 *   &callback=fn  -> JSONP wrap
 *
 * Εγκατάσταση: δες posts-SETUP.md
 */

const SHEET_ID = 'PASTE_SHEET_ID_HERE'; // ← το ID του Google Sheet (μεταξύ /d/ και /edit)

const RAW_SHEET   = 'Πρώτη Ύλη';
const RAW_HEADERS = ['ID', 'Καταχώρηση', 'Ημερομηνία', 'Έως', 'Τμήμα', 'Τίτλος', 'Κείμενο', 'Συντάκτης'];

// ---- Ιστοσελίδα (τελευταία δράσεις — δημόσιο WordPress REST, ΧΩΡΙΣ token) ----
const SITE_URL     = 'https://76.life';
const WEB_CATEGORY = '29';   // κατηγορία «ΔΡΑΣΕΙΣ» στο 76.life (κενό '' = όλες οι δημοσιεύσεις)
const WEB_LIMIT    = 12;

function ss_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById(SHEET_ID);
}

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(RAW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RAW_SHEET);
    sh.getRange(1, 1, 1, RAW_HEADERS.length).setValues([RAW_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, RAW_HEADERS.length);
  }
  return sh;
}

/** Τρέξε ΜΙΑ φορά από τον editor (Run > initSheet) για να δημιουργηθεί το φύλλο. */
function initSheet() {
  sheet_();
}

function read_() {
  const sh = sheet_();
  const vals = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!String(r[0] || '').trim()) continue; // απαιτείται ID
    const o = {};
    RAW_HEADERS.forEach((h, j) => o[h] = (r[j] === undefined || r[j] === null) ? '' : String(r[j]));
    rows.push(o);
  }
  return rows;
}

function addRow_(obj) {
  const sh = sheet_();
  const id = 'id' + Date.now() + Math.floor(Math.random() * 1000);
  obj.ID = id;
  obj['Καταχώρηση'] = new Date().toISOString();
  sh.appendRow(RAW_HEADERS.map(h => (obj[h] !== undefined && obj[h] !== null) ? obj[h] : ''));
  return id;
}

function updateRow_(id, obj) {
  const sh = sheet_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) {
      RAW_HEADERS.forEach((h, j) => {
        if (h === 'ID' || h === 'Καταχώρηση') return;
        if (obj[h] !== undefined) sh.getRange(i + 1, j + 1).setValue(obj[h]);
      });
      return true;
    }
  }
  return false;
}

function deleteRow_(id) {
  const sh = sheet_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) { sh.deleteRow(i + 1); return true; }
  }
  return false;
}

/** Δημόσια ανάγνωση δράσεων από το 76.life (WordPress REST). Χωρίς token. */
function fetchWebPosts_() {
  let url = SITE_URL + '/wp-json/wp/v2/posts?per_page=' + WEB_LIMIT +
            '&_fields=id,date,link,title,excerpt';
  if (WEB_CATEGORY) url += '&categories=' + encodeURIComponent(WEB_CATEGORY);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return { ok: false, error: 'HTTP ' + resp.getResponseCode() };
  const data = JSON.parse(resp.getContentText() || '[]');
  const items = (data || []).map(function (p) {
    return {
      id: p.id,
      date: p.date || '',
      link: p.link || '',
      title: (p.title && p.title.rendered) || '',
      excerpt: (p.excerpt && p.excerpt.rendered) || ''
    };
  });
  return { ok: true, items: items };
}

function out_(obj, cb) {
  const t = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + t + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON);
}

function handle_(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  const action = p.action || 'list';
  try {
    if (action === 'list') {
      return out_({ ok: true, raw: read_(), ts: new Date().toISOString() }, cb);
    }
    if (action === 'web') {
      return out_(fetchWebPosts_(), cb);
    }
    if (action === 'add') {
      const obj = {};
      RAW_HEADERS.forEach(h => { if (p[h] !== undefined) obj[h] = p[h]; });
      return out_({ ok: true, id: addRow_(obj) }, cb);
    }
    if (action === 'update') {
      const obj = {};
      RAW_HEADERS.forEach(h => { if (p[h] !== undefined) obj[h] = p[h]; });
      return out_({ ok: updateRow_(p.ID || p.id, obj) }, cb);
    }
    if (action === 'delete') {
      return out_({ ok: deleteRow_(p.ID || p.id) }, cb);
    }
    return out_({ ok: false, error: 'unknown action: ' + action }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  }
}

function doGet(e) {
  return handle_(e);
}

function doPost(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      const b = JSON.parse(e.postData.contents);
      e.parameter = Object.assign({}, e.parameter, b);
    } catch (_) {}
  }
  return handle_(e);
}
