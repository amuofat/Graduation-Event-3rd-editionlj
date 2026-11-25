/**
 * DELTA 2027 · registrations → Google Sheet
 *
 * Setup (5 minutes):
 * 1. Create a new Google Sheet, e.g. "DELTA 2027 – Registrations".
 * 2. Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 * 3. Deploy → New deployment → type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Authorise, then copy the Web app URL (ends with /exec).
 * 5. In index.html, paste it into:  const SHEET_URL = '...';
 *
 * Each registration becomes one row. The Pass ID is what the QR code on the Golden Pass carries,
 * so the welcome team can search it with Ctrl+F on the day.
 */

const SHEET_NAME = 'Registrations';
const HEADERS = [
  'Submitted at', 'Pass ID', 'Full name', 'Email', 'Phone', 'Guests',
  'Guest 1 name', 'Guest 1 relation',
  'Guest 2 name', 'Guest 2 relation',
  'Guest 3 name', 'Guest 3 relation',
  'Language', 'Seats (engineer + guests)'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.type !== 'delta-registration') return json({ ok: false, error: 'wrong type' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0B1F3A').setFontColor('#FFBF00');
    }

    const guests = Math.max(0, Math.min(3, Number(d.guestCount) || 0));
    sh.appendRow([
      new Date(), safe(d.passId), safe(d.name), safe(d.email), "'" + safe(d.phone), guests,
      safe(d.guest1Name), safe(d.guest1Relation),
      safe(d.guest2Name), safe(d.guest2Relation),
      safe(d.guest3Name), safe(d.guest3Relation),
      safe(d.lang), 1 + guests
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Stops anyone typing a formula (=, +, -, @) into a cell through the form.
function safe(v) {
  const s = String(v == null ? '' : v).slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
