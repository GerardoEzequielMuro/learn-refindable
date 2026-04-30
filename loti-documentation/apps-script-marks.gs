/**
 * Loti Properties Audit — shared marks backend (Google Apps Script)
 *
 * SETUP (one-time, ~3 minutes):
 *   1. Go to https://sheets.google.com and create a new blank Sheet.
 *      Name it e.g. "Loti Properties Audit Marks".
 *   2. In that Sheet: Extensions → Apps Script.
 *   3. Delete the placeholder code and paste THIS ENTIRE FILE.
 *   4. Save (disk icon).
 *   5. Click "Deploy" → "New deployment".
 *      - Gear icon → "Web app".
 *      - Description: "Loti marks v1"
 *      - Execute as: Me (your account)
 *      - Who has access: Anyone
 *      - Click "Deploy", authorize when prompted.
 *   6. Copy the "Web app URL". It ends in /exec.
 *   7. Paste that URL into properties.html as the SCRIPT_URL constant.
 *   8. Commit + push.
 *
 * Re-deploying after edits: Deploy → Manage deployments → pencil icon → New version.
 *
 * Schema (auto-created on first call):
 *   Sheet "marks": object | internal | disposition | issue | notes | updatedAt
 *   Key = (object, internal). Last write wins.
 */

const SHEET_NAME = 'marks';
const HEADERS = ['object', 'internal', 'disposition', 'issue', 'notes', 'updatedAt'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const sh = getSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return jsonOut_({ ok: true, marks: [] });
    const headers = values[0];
    const marks = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);
      if (!obj.object || !obj.internal) continue;
      marks.push(obj);
    }
    return jsonOut_({ ok: true, marks: marks });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const body = JSON.parse(e.postData.contents);
    const sh = getSheet_();

    if (body.action === 'clear') {
      return clearObject_(sh, body.object);
    }

    return upsert_(sh, body);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function upsert_(sh, body) {
  if (!body.object || !body.internal) {
    return jsonOut_({ ok: false, error: 'Missing object or internal' });
  }
  const data = sh.getDataRange().getValues();
  const now = new Date().toISOString();
  const row = [
    body.object,
    body.internal,
    body.disposition || '',
    body.issue || '',
    body.notes || '',
    now
  ];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.object && data[i][1] === body.internal) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return jsonOut_({ ok: true, action: 'updated', updatedAt: now });
    }
  }
  sh.appendRow(row);
  return jsonOut_({ ok: true, action: 'inserted', updatedAt: now });
}

function clearObject_(sh, object) {
  if (!object) return jsonOut_({ ok: false, error: 'Missing object' });
  const data = sh.getDataRange().getValues();
  let removed = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === object) {
      sh.deleteRow(i + 1);
      removed++;
    }
  }
  return jsonOut_({ ok: true, action: 'cleared', removed: removed });
}
