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
 *
 * Endpoints:
 *   GET ?callback=fn                  → JSONP list of viewers {count, viewers:[{name,code,legacy_text}]}
 *   GET ?callback=fn&code=XXXX        → JSONP single record {record:{name,code,q1..q7,legacy_text,additions:[...]}}
 *   POST {action:'add', code, type, content}  → append to "additions" sheet
 *   POST {timestamp,name,email,code,q1..q7,legacy_text}  → append to data sheet
 */

const SHEET_ID = '1nC7UoBpLItL0vOwDJkOE9II7YPeHp48cbBdmDGq_U_w';
const SHEET_NAME = 'data';
const ADDITIONS_SHEET = 'additions';
const HEADERS = [
  'timestamp', 'name', 'email', 'code',
  'q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7',
  'legacy_text'
];
const ADDITION_HEADERS = ['timestamp', 'code', 'type', 'content'];

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function getAdditionsSheet_(createIfMissing) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(ADDITIONS_SHEET);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(ADDITIONS_SHEET);
    sheet.appendRow(ADDITION_HEADERS);
  } else if (sheet && sheet.getLastRow() === 0) {
    sheet.appendRow(ADDITION_HEADERS);
  }
  return sheet;
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

function respond_(result, e) {
  const json = JSON.stringify(result);
  if (e && e.parameter && e.parameter.callback) {
    const cb = String(e.parameter.callback).replace(/[^a-zA-Z0-9_$]/g, '');
    return ContentService
      .createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getAdditionsForCode_(code) {
  const sheet = getAdditionsSheet_(false);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, ADDITION_HEADERS.length).getValues();
  const result = [];
  data.forEach(function (row) {
    const rowCode = String(row[1] || '').padStart(4, '0');
    if (rowCode === code) {
      let ts = '';
      if (row[0]) {
        try { ts = new Date(row[0]).toISOString(); }
        catch (e) { ts = String(row[0]); }
      }
      result.push({
        timestamp: ts,
        code: rowCode,
        type: String(row[2] || ''),
        content: String(row[3] || '')
      });
    }
  });
  return result;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.code) {
    return doGetByCode_(e);
  }
  return doGetList_(e);
}

function doGetList_(e) {
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

  return respond_({ count: viewers.length, viewers: viewers }, e);
}

function doGetByCode_(e) {
  const code = String(e.parameter.code).padStart(4, '0');
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  let record = null;

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowCode = String(data[i][3] || '').padStart(4, '0');
      if (rowCode === code) {
        record = {
          name: String(data[i][1] || ''),
          code: rowCode,
          q1: String(data[i][4] || ''),
          q2: String(data[i][5] || ''),
          q3: String(data[i][6] || ''),
          q4: String(data[i][7] || ''),
          q5: String(data[i][8] || ''),
          q6: String(data[i][9] || ''),
          q7: String(data[i][10] || ''),
          legacy_text: String(data[i][11] || ''),
          additions: getAdditionsForCode_(code)
        };
        break;
      }
    }
  }

  return respond_({ record: record }, e);
}

function doPost(e) {
  try {
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

    if (payload && payload.action === 'add') {
      return handleAddition_(payload);
    }

    return handleEntry_(payload);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleEntry_(payload) {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

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
}

function handleAddition_(payload) {
  const sheet = getAdditionsSheet_(true);
  const codeStr = String(payload.code || '').padStart(4, '0');
  const row = [
    payload.timestamp || new Date().toISOString(),
    "'" + codeStr,
    String(payload.type || 'text'),
    String(payload.content || '')
  ];
  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
