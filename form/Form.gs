/**
 * Φόρμα ενδιαφέροντος του okpublisher.gr (ξεχωριστό Apps Script, ΔΕΝ αγγίζει τη δημοσίευση).
 *
 * Το site στέλνει POST (JSON) σε αυτό το web app. Κάθε αίτημα:
 *   - γράφεται σε μια γραμμή του Sheet στο οποίο ανήκει το script (καρτέλα «Ενδιαφέροντα»),
 *   - στέλνεται ως email στον λογαριασμό που τρέχει το script (okpublisher.gr) — σε κανέναν άλλον.
 * Προστασία: κρυφό πεδίο (website) για bots, όρια μήκους, το πολύ 30 αιτήματα την ώρα.
 *
 * Εγκατάσταση (μία φορά, στο okpublisher.gr@gmail.com):
 *   1. Νέο Google Sheet «okpublisher — ενδιαφέροντα» → Επεκτάσεις → Apps Script → επικόλληση αυτού του αρχείου.
 *   2. Ανάπτυξη → Νέα ανάπτυξη → Εφαρμογή ιστού · Εκτέλεση ως: Εγώ · Πρόσβαση: Οποιοσδήποτε → Ανάπτυξη → άδειες.
 *   3. Η διεύθυνση /exec μπαίνει στο FORM_URL του index.html του site.
 *   Δοκιμή: συνάρτηση testForm() → μία δοκιμαστική γραμμή + email.
 */
const FORM_ = {
  SHEET: 'Ενδιαφέροντα',
  HEAD: ['Χρόνος', 'Όνομα', 'Επιχείρηση', 'Email', 'Τηλέφωνο', 'Πακέτο', 'Instagram/Facebook', 'Μήνυμα', 'Γλώσσα', 'Κατάσταση'],
  MAX_PER_HOUR: 30,
  LEN: { name: 120, business: 160, email: 160, phone: 40, plan: 40, social: 160, message: 3000, lang: 4 },
};

function doPost(e) {
  let d = {};
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  if (d.website) return out_({ ok: true });                       // bot: σιωπηλά τίποτα
  const r = {};
  Object.keys(FORM_.LEN).forEach(function (k) { r[k] = String(d[k] == null ? '' : d[k]).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, FORM_.LEN[k]); });
  if (!r.name || !r.business || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r.email)) return out_({ ok: false, error: 'missing fields' });
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('n') || 0);
  if (n >= FORM_.MAX_PER_HOUR) return out_({ ok: false, error: 'busy' });
  cache.put('n', String(n + 1), 3600);
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sh = sheet_();
    // Τα κελιά γράφονται ως κείμενο (όχι τύποι): ό,τι αρχίζει με = + - @ παίρνει μπροστά απόστροφο.
    const safe = function (v) { return /^[=+\-@]/.test(v) ? "'" + v : v; };
    sh.appendRow([new Date()].concat([r.name, r.business, r.email, r.phone, r.plan, r.social, r.message, r.lang].map(safe), ['νέο']));
  } finally { lock.releaseLock(); }
  const me = Session.getEffectiveUser().getEmail();
  if (me) {
    MailApp.sendEmail({
      to: me, replyTo: r.email, name: 'okpublisher.gr',
      subject: '[okpublisher.gr] Νέο ενδιαφέρον: ' + r.business + (r.plan ? ' (' + r.plan + ')' : ''),
      body: ['Όνομα: ' + r.name, 'Επιχείρηση: ' + r.business, 'Email: ' + r.email, 'Τηλέφωνο: ' + (r.phone || '—'),
        'Πακέτο: ' + (r.plan || 'δεν ξέρει ακόμα'), 'Instagram/Facebook: ' + (r.social || '—'), 'Γλώσσα: ' + (r.lang || 'el'), '', r.message || '(χωρίς μήνυμα)',
        '', 'Απάντηση: απλώς «Απάντηση» σε αυτό το email (πάει στον ενδιαφερόμενο).'].join('\n'),
    });
  }
  return out_({ ok: true });
}

function doGet() { return out_({ ok: true, service: 'okpublisher form' }); }

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FORM_.SHEET);
  if (!sh) {
    sh = ss.insertSheet(FORM_.SHEET);
    sh.getRange(1, 1, 1, FORM_.HEAD.length).setValues([FORM_.HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function testForm() {
  const res = doPost({ postData: { contents: JSON.stringify({ name: 'Δοκιμή', business: 'Δοκιμαστική επιχείρηση', email: 'test@example.com', plan: 'Μεσαίο', message: 'Δοκιμή φόρμας', lang: 'el' }) } });
  console.log(res.getContent());
}
