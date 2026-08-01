/**
 * Personnel Accountability Log - Google Sheets Backend
 *
 * SETUP:
 * 1. Create a new Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Delete any starter code and paste this entire file in.
 * 4. Click Deploy > New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy, authorize the permissions it asks for, and copy the
 *    "Web app URL" it gives you.
 * 6. Paste that URL into the SHEET_API_URL constant near the top of
 *    accountability-log-google-sheets.html
 *
 * RETENTION (25-day auto-purge):
 * 1. In the Apps Script editor, click the clock icon on the left ("Triggers").
 * 2. Click "+ Add Trigger".
 * 3. Function to run: purgeOldEntries
 * 4. Event source: Time-driven > Day timer > pick any time (e.g. 2am-3am)
 * 5. Save.
 * This runs once a day and deletes any row older than 25 days.
 */

var SHEET_NAME = 'Log';

function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getOrCreateSheet();
    var action = e.parameter.action;

    if (action === 'list') {
      return respond(listEntries(sheet));
    } else if (action === 'signout') {
      return respond(doSignOut(sheet, e.parameter));
    } else if (action === 'signin') {
      return respond(doSignIn(sheet, e.parameter));
    } else {
      return respond({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return respond({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID', 'Name', 'Buddy', 'Platoon', 'Location', 'TimeOut', 'TimeIn']);
  }
  return sheet;
}

function listEntries(sheet) {
  var data = sheet.getDataRange().getValues();
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    entries.push({
      id: String(row[0]),
      name: row[1],
      buddy: row[2],
      platoon: row[3],
      location: row[4],
      timeOut: row[5] instanceof Date ? row[5].toISOString() : row[5],
      timeIn: row[6] instanceof Date ? row[6].toISOString() : (row[6] || null)
    });
  }
  return { entries: entries };
}

function doSignOut(sheet, params) {
  var id = Utilities.getUuid();
  var now = new Date().toISOString();
  sheet.appendRow([
    id,
    params.name || '',
    params.buddy || '',
    params.platoon || '',
    params.location || '',
    now,
    ''
  ]);
  return { success: true, id: id };
}

function doSignIn(sheet, params) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === params.id) {
      sheet.getRange(i + 1, 7).setValue(new Date().toISOString());
      return { success: true };
    }
  }
  return { success: false, error: 'Entry not found' };
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Deletes any row whose sign-out time is older than 25 days.
 * Set this up as a daily time-driven trigger (see setup notes above).
 */
function purgeOldEntries() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var cutoff = new Date(Date.now() - 25 * 24 * 60 * 60 * 1000);
  for (var i = data.length - 1; i >= 1; i--) {
    var t = new Date(data[i][5]);
    if (t < cutoff) {
      sheet.deleteRow(i + 1);
    }
  }
}
