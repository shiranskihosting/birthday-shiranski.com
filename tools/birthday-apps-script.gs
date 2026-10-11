/**
 * מעקב ברכות יום הולדת – שירנסקי מארח
 * Apps Script של הגיליון "מעקב ברכות יום הולדת – שירנסקי מארח".
 * כל כניסה לדף birthday.shiranski.com וכל לחיצה על "רוצה לתאם ערב" מעדכנות שורה אחת לכל לקוח: שם, תאריך שליחת ההודעה (מהקישור), פעולה (נשלח → כניסה → תיאום) בלשונית הראשונה.
 * פרסום: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone > Deploy.
 */
const ACTIONS = ['כניסה', 'תיאום'];
const RANK = {'נשלח': 0, 'כניסה': 1, 'תיאום': 2}; // שורה אחת לכל לקוח: הפעולה רק מתקדמת, לא חוזרת אחורה

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    const action = ACTIONS.indexOf(b.a) >= 0 ? b.a : '';
    if (!action) return out('bad');
    const name = clean(b.n, 30) || '(בלי שם)';
    const m = String(b.s || '').match(/^(\d{4})(\d{2})(\d{2})$/);
    const sent = m ? (+m[3]) + '.' + (+m[2]) + '.' + m[1] : '';
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const book = SpreadsheetApp.getActiveSpreadsheet();
      const sh = book.getSheets()[0];
      if (sh.getLastRow() === 0) sh.appendRow(['שם', 'תאריך שליחת הודעה', 'פעולה']);
      const last = sh.getLastRow();
      const rows = last > 1 ? sh.getRange(2, 1, last - 1, 3).getDisplayValues() : [];
      let at = -1;
      for (let i = rows.length - 1; i >= 0; i--) if (rows[i][0] === name && rows[i][1] === sent) { at = i; break; }
      if (at < 0) sh.appendRow([name, "'" + sent, action]); // the date is kept as text so it always matches
      else if ((RANK[rows[at][2]] ?? -1) < RANK[action]) sh.getRange(at + 2, 3).setValue(action);
    } finally { lock.releaseLock(); }
    return out('ok');
  } catch (err) {
    return out('error');
  }
}
function doGet() { return out('ok'); }
function clean(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f<>=+@]/g, '').trim().slice(0, max); }
function out(s) { return ContentService.createTextOutput(s); }
