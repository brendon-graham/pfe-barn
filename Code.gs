/**
 * PFE Barn (Rangitata Barn calf feeding) — Google Apps Script sync backend.
 *
 * Setup:
 *  1. Create a new Google Sheet, e.g. "Rangitata Barn App".
 *  2. Extensions > Apps Script. Delete the default code, paste this file, Save.
 *  3. Deploy > New deployment > type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone   (must be "Anyone", not "Anyone with a Google account")
 *  4. Copy the /exec URL. Paste it into the app: Settings > Sync endpoint.
 *  5. Any Code.gs change later: Deploy > Manage deployments > Edit > New version.
 *
 * Data model: one "Records" sheet, one row per record.
 *  kind = pen | move | lift | feed | health
 *  body = JSON string of the record's fields
 *  Row-level upsert: newer updatedAt wins. Deletes are a deleted=true flag (tombstone), never row removal.
 */

var SHEET_RECORDS = 'Records';
var SHEET_META = 'Meta';
var HEADERS = ['id', 'kind', 'date', 'body', 'updatedAt', 'deleted'];

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  var payload = {};
  try { payload = JSON.parse(e.postData.contents || '{}'); } catch (err) { payload = {}; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    upsert_(payload.rows || []);
  } finally {
    lock.releaseLock();
  }
  return json_(readAll_());
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// Sheets coerces date-looking strings to Date objects. Columns are set to plain text on
// create, and these guards turn any Date back into the string the app wrote.
function normDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}
function normStamp_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return v.toISOString();
  return String(v);
}

function readAll_() {
  var sh = sheet_(SHEET_RECORDS, HEADERS);
  var vals = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var v = vals[i];
    if (!v[0]) continue;
    rows.push({
      id: String(v[0]),
      kind: String(v[1]),
      date: normDate_(v[2]),
      body: String(v[3]),
      updatedAt: normStamp_(v[4]),
      deleted: (v[5] === true || v[5] === 'true' || v[5] === 'TRUE')
    });
  }
  return { ok: true, lastModified: getMeta_('lastModified') || '', rows: rows };
}

function upsert_(incoming) {
  var sh = sheet_(SHEET_RECORDS, HEADERS);
  var vals = sh.getDataRange().getValues();
  var idx = {};
  for (var i = 1; i < vals.length; i++) if (vals[i][0]) idx[String(vals[i][0])] = i;
  var appends = [];
  var changed = false;
  for (var j = 0; j < incoming.length; j++) {
    var r = incoming[j];
    if (!r || !r.id) continue;
    var rowArr = [String(r.id), String(r.kind || ''), String(r.date || ''),
      r.body == null ? '' : String(r.body), String(r.updatedAt || ''), r.deleted ? 'true' : 'false'];
    if (idx.hasOwnProperty(r.id)) {
      var existing = normStamp_(vals[idx[r.id]][4]);
      if (String(r.updatedAt) > existing) {
        sh.getRange(idx[r.id] + 1, 1, 1, HEADERS.length).setValues([rowArr]);
        changed = true;
      }
    } else {
      idx[r.id] = -1;
      appends.push(rowArr);
    }
  }
  if (appends.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, appends.length, HEADERS.length).setNumberFormat('@').setValues(appends);
    changed = true;
  }
  if (changed) setMeta_('lastModified', new Date().toISOString());
}

function getMeta_(k) {
  var sh = sheet_(SHEET_META, ['key', 'value']);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) if (vals[i][0] === k) return vals[i][1];
  return '';
}

function setMeta_(k, v) {
  var sh = sheet_(SHEET_META, ['key', 'value']);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === k) { sh.getRange(i + 1, 2).setValue(v); return; }
  }
  sh.appendRow([k, v]);
}
