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
 *   action=admin & password=... & op=generate & note=... → generate EN/JA topic phrase from staff notes via Claude API
 *
 * To use op=generate, set CLAUDE_API_KEY in Project Settings > Script Properties.
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

  if (op === 'generate') {
    return generateTopicPhrase(params.note || '');
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

/**
 * Uses the Claude API to turn a short staff note into a natural clause
 * for both the English and Japanese thank-you email templates.
 */
function generateTopicPhrase(note) {
  if (!note) {
    return { ok: false, error: 'empty note' };
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'CLAUDE_API_KEY not configured' };
  }

  try {
    var prompt = 'A boutique staff member wrote this short note about a customer visit (it may be written in Japanese or English):\n"' + note + '"\n\n' +
      'Generate two short natural clauses based on this note, for use in a thank-you email:\n' +
      '- "en": a clause completing the English sentence "It was a pleasure meeting you and ___." ' +
      'It can describe a topic discussed (e.g. "talking with you about your trip to Kyoto") or something done together ' +
      '(e.g. "choosing your new frame together"), whichever fits the note best. If the note is written in Japanese, translate its meaning into natural English.\n' +
      '- "ja": a clause completing the Japanese sentence "___、大変嬉しく思っております。" ' +
      'expressed naturally in Japanese (e.g. "京都旅行のお話をお伺いできましたこと" or "一緒にフレームをお選びできましたこと").\n\n' +
      'Respond with ONLY a JSON object on a single line, no markdown, no extra text: {"en": "...", "ja": "..."}';

    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 200,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var body = JSON.parse(response.getContentText());
    if (responseCode !== 200) {
      var apiError = (body.error && body.error.message) || ('HTTP ' + responseCode);
      return { ok: false, error: 'Claude API error: ' + apiError };
    }

    var textBlock = null;
    for (var i = 0; i < (body.content || []).length; i++) {
      if (body.content[i].type === 'text') {
        textBlock = body.content[i];
        break;
      }
    }
    if (!textBlock) {
      return { ok: false, error: 'no text in response' };
    }

    var text = textBlock.text.trim();
    // Strip markdown code fences if present (e.g. ```json ... ```)
    text = text.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '').trim();

    var parsed;
    try {
      parsed = JSON.parse(text);
    } catch (err) {
      return { ok: false, error: 'invalid response format: ' + text };
    }

    if (!parsed.en || !parsed.ja) {
      return { ok: false, error: 'missing en/ja in response: ' + text };
    }

    return { ok: true, en: parsed.en, ja: parsed.ja };
  } catch (err) {
    return { ok: false, error: 'Exception: ' + err.message };
  }
}
