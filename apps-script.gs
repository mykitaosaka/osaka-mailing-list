/**
 * Google Apps Script webhook for MYKITA OSAKA Frame Warranty Registration.
 * Deploy as Web App (Execute as: Me / Access: Anyone) and paste the
 * resulting /exec URL into SHEET_WEBHOOK_URL in index.html.
 *
 * Spreadsheet column order:
 * Timestamp | 会員番号 | Name | Email | Opt-in | Country | Language | Follow-up Sent | Staff Notes
 * (会員番号 is filled in manually in the sheet, so it is left blank here.)
 */
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  var headers = ['Timestamp', '会員番号', 'Name', 'Email', 'Opt-in', 'Country', 'Language', 'Follow-up Sent', 'Staff Notes'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  var country = data.country || '';
  if (data.countryCode) {
    country += ' (' + data.countryCode + ')';
  }

  sheet.appendRow([
    data.timestamp || '',
    data.memberNo || '',
    data.name || '',
    data.email || '',
    data.optin ? 'Yes' : 'No',
    country,
    (data.lang || '').toUpperCase(),
    '',
    ''
  ]);

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Admin API (read-only login, list, delete).
 * Set ADMIN_PASSWORD in Project Settings > Script Properties (NOT in code).
 *
 * GET params:
 *   action=admin & password=...                → list records
 *   action=admin & password=... & op=delete & row=N → delete row N
 *   action=admin & password=... & op=followup & row=N & status=pending|sent|na → set follow-up status
 *   action=admin & password=... & op=note & row=N & text=... → set staff notes
 */
function doGet(e) {
  var params = e.parameter || {};
  var out;

  if (params.action === 'admin') {
    out = handleAdminRequest(params);
  } else {
    out = { ok: false, error: 'unknown action' };
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAdminRequest(params) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!stored || params.password !== stored) {
    return { ok: false, error: 'invalid password' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var op = params.op || 'list';

  if (op === 'delete') {
    var row = parseInt(params.row, 10);
    if (!row || row < 2 || row > sheet.getLastRow()) {
      return { ok: false, error: 'invalid row' };
    }
    sheet.deleteRow(row);
    return { ok: true };
  }

  if (op === 'followup') {
    var fuRow = parseInt(params.row, 10);
    if (!fuRow || fuRow < 2 || fuRow > sheet.getLastRow()) {
      return { ok: false, error: 'invalid row' };
    }
    var value = '';
    if (params.status === 'sent') value = '✔︎';
    else if (params.status === 'na') value = 'N/A';
    sheet.getRange(fuRow, 8).setValue(value);
    return { ok: true };
  }

  if (op === 'note') {
    var noteRow = parseInt(params.row, 10);
    if (!noteRow || noteRow < 2 || noteRow > sheet.getLastRow()) {
      return { ok: false, error: 'invalid row' };
    }
    sheet.getRange(noteRow, 9).setValue(params.text || '');
    return { ok: true };
  }

  return { ok: true, records: readRecords(sheet) };
}

function readRecords(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return []; }

  var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var records = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    records.push({
      row: i + 2,
      timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ''),
      memberNo: row[1] || '',
      name: row[2] || '',
      email: row[3] || '',
      optin: row[4] || '',
      country: row[5] || '',
      lang: row[6] || '',
      followupStatus: row[7] === '✔︎' ? 'sent' : row[7] === 'N/A' ? 'na' : '',
      note: row[8] || ''
    });
  }
  return records;
}
