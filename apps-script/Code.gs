// KnowledgeFlood eligibility form -> Google Sheet.
// Paste into the Sheet via Extensions > Apps Script, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Put the resulting /exec URL in index.html.

const SHEET_NAME = 'Submissions';
const NOTIFY_EMAIL = ''; // optional: an address to email on each submission
const FIELDS = ['name', 'title', 'email', 'phone', 'company', 'industry', 'employees', 'years', 'country', 'state', 'data', 'consent'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp'].concat(FIELDS, ['Eligible']));
      sheet.setFrozenRows(1);
    }

    const p = e.parameter || {};
    const eligible = p.employees !== 'Under 25' && p.years !== 'Under 2' && p.country !== 'Other';
    const row = [new Date()].concat(FIELDS.map(f => clean(p[f])), [eligible ? 'Yes' : 'No']);
    sheet.appendRow(row);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'New KnowledgeFlood submission: ' + clean(p.company),
        FIELDS.map(f => f + ': ' + clean(p[f])).join('\n') + '\neligible: ' + (eligible ? 'Yes' : 'No'));
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Stop submitted text from being interpreted as a spreadsheet formula.
function clean(v) {
  const s = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
