/**
 * "מה שישאר" — Google Apps Script webhook
 *
 * Setup:
 *   1. Open https://script.google.com → New project.
 *   2. Paste this file's contents into Code.gs.
 *   3. Deploy → New deployment → Type: Web app.
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   4. Copy the deployed Web app URL into index.html → SHEET_WEBHOOK.
 */

const SHEET_ID = '1nC7UoBpLItL0vOwDJkOE9II7YPeHp48cbBdmDGq_U_w';
const SHEET_NAME = 'data';
const HEADERS = [
  'timestamp', 'name', 'email', 'code',
  'q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7',
  'legacy_text'
];

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const isEmpty = firstRow.every(function (cell) { return cell === '' || cell === null; });
  if (isEmpty) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const viewers = [];

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    data.forEach(function (row) {
      if (!row[1] && !row[3]) return;
      viewers.push({
        name: String(row[1] || ''),
        code: String(row[3] || '').padStart(4, '0'),
        legacy_text: String(row[11] || '')
      });
    });
  }

  const result = { count: viewers.length, viewers: viewers };
  const json = JSON.stringify(result);

  if (e && e.parameter && e.parameter.callback) {
    const callback = String(e.parameter.callback).replace(/[^a-zA-Z0-9_$]/g, '');
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const sheet = getSheet_();
    ensureHeaders_(sheet);

    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        Logger.log('JSON parse failed; raw contents: ' + e.postData.contents);
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const dk = 'לא יודע/ת';
    const qVal = function (key) {
      const v = payload[key];
      if (v === null || v === undefined || v === '') return dk;
      return v;
    };

    const codeStr = String(payload.code || '').padStart(4, '0');

    const row = [
      payload.timestamp || new Date().toISOString(),
      payload.name || '',
      payload.email || '',
      "'" + codeStr,
      qVal('q1'), qVal('q2'), qVal('q3'), qVal('q4'),
      qVal('q5'), qVal('q6'), qVal('q7'),
      payload.legacy_text || ''
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, row: row }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
