/**
 * Wolfson Cardiac Surgery — Schedule Sync Backend
 * ------------------------------------------------
 * Stores app state (docs, sched, wishes, sessions, timeoff, requests) as
 * JSON blobs in a Google Sheet, one row per key. Generic by design so
 * the front-end can add new top-level keys (like "sessions") without
 * requiring further backend changes.
 *
 * DEPLOYMENT (keeps the existing GS_URL the same):
 *  1. Open the Apps Script project already bound to your Sheet
 *     (Extensions > Apps Script from the Sheet, or script.google.com).
 *  2. Select ALL existing code in Code.gs and replace it with this file's
 *     contents.
 *  3. Deploy > Manage deployments > (pencil icon) Edit > Version: "New version"
 *     > Deploy. This keeps the same /exec URL that's already hardcoded
 *     in index.html as GS_URL — no front-end change needed for this step.
 *  4. Test: open the GS_URL in a browser, you should see {"ok":true,"data":{...}}.
 */

const SHEET_NAME = 'wf_data';       // single sheet, one row per key
const KNOWN_KEYS = ['docs', 'sched', 'wishes', 'sessions', 'timeoff', 'requests'];

function _getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value_json']]);
  }
  return sh;
}

function _readAll_() {
  const sh = _getSheet_();
  const rows = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0];
    const raw = rows[i][1];
    if (!key) continue;
    try {
      out[key] = raw ? JSON.parse(raw) : (key === 'sched' || key === 'wishes' ? {} : []);
    } catch (e) {
      out[key] = (key === 'sched' || key === 'wishes') ? {} : [];
    }
  }
  return out;
}

function _writeKey_(sh, key, value) {
  const rows = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { rowIdx = i + 1; break; }
  }
  const json = JSON.stringify(value === undefined ? null : value);
  if (rowIdx === -1) {
    sh.appendRow([key, json]);
  } else {
    sh.getRange(rowIdx, 2).setValue(json);
  }
}

function _writeAll_(data) {
  const sh = _getSheet_();
  // Lock to avoid concurrent write corruption when multiple devices push at once
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    Object.keys(data).forEach(key => {
      _writeKey_(sh, key, data[key]);
    });
  } finally {
    lock.releaseLock();
  }
}

function _json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const data = _readAll_();
    return _json_({ ok: true, data: data });
  } catch (err) {
    return _json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'setAll') {
      // Only persist known keys (plus any future ones the front-end adds —
      // we don't hard-reject unknown keys, so this stays forward compatible).
      _writeAll_(body.data || {});
      return _json_({ ok: true });
    }
    return _json_({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return _json_({ ok: false, error: String(err) });
  }
}