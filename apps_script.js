// ======================================
// Soltri 선생산 무덤 관리 - Apps Script
// 구글시트 Apps Script 에디터에 붙여넣고
// 웹앱으로 배포 (액세스: 모든 사용자)
// ======================================

const SS_ID = '1MsmVKtz5NTxIIoj3efXYPLEhL3GaONW5LAlRNjKk7s0';
const SHEET_NAME = '무덤';

function getSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action || 'load';
  const callback = e.parameter.callback || 'cb';
  const sheet = getSheet();

  if (action === 'save') {
    const type = e.parameter.type || 'graveyard';
    const data = e.parameter.data || '{}';
    const cell = type === 'completed' ? 'A2' : 'A1';
    sheet.getRange(cell).setValue(data);
    return ContentService
      .createTextOutput(callback + '({"ok":true})')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // load: graveyard(A1) + completed(A2) 동시 반환
  let grave = '{}', comp = '{}';
  try { grave = sheet.getRange('A1').getValue() || '{}'; } catch(e) {}
  try { comp  = sheet.getRange('A2').getValue() || '{}'; } catch(e) {}

  const result = '{"graveyard":' + grave + ',"completed":' + comp + '}';
  return ContentService
    .createTextOutput(callback + '(' + result + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
