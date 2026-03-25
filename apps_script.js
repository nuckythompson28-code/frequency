// ======================================
// Soltri 선생산 무덤 관리 - Apps Script
// 구글시트 Apps Script 에디터에 붙여넣고
// 웹앱으로 배포 (액세스: 모든 사용자)
// ======================================

const SS_ID = '1MsmVKtz5NTxIIoj3efXYPLEhL3GaONW5LAlRNjKk7s0';
const SHEET_NAME = '무덤';
const TABLE_SHEET = '생산지시이력';

function getSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function getTableSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(TABLE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TABLE_SHEET);
    sheet.getRange('A1:F1').setValues([['치수', '재질', '상태', '지시일', '완료일', '비고']]);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#334155').setFontColor('#e2e8f0');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 110);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 100);
  }
  return sheet;
}

// JSON → 테이블 시트에 펼쳐서 저장
function syncTable(graveJson, compJson) {
  const sheet = getTableSheet();
  // 헤더 외 전부 지우기
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();

  const rows = [];
  const grave = JSON.parse(graveJson || '{}');
  const comp = JSON.parse(compJson || '{}');

  // graveyard → 생산지시 중
  for (const [key, val] of Object.entries(grave)) {
    const parts = key.split('|');
    const chisu = parts[0] || key;
    const jaejil = parts[1] || '';
    const graveDate = typeof val === 'string' ? val.slice(0, 10) : (val.graveDate || '').slice(0, 10);
    rows.push([chisu, jaejil, '생산지시 중', graveDate, '', '']);
  }

  // completed → 생산지시 완료
  for (const [key, val] of Object.entries(comp)) {
    const parts = key.split('|');
    const chisu = parts[0] || key;
    const jaejil = parts[1] || '';
    let graveDate = (val.graveDate || '').slice(0, 10);
    let compDate = (val.completedDate || '').slice(0, 10);
    // 6자리 날짜(YYMMDD) → YYYY-MM-DD 정규화
    if (graveDate.length === 6) graveDate = '20' + graveDate.slice(0,2) + '-' + graveDate.slice(2,4) + '-' + graveDate.slice(4,6);
    if (compDate.length === 6) compDate = '20' + compDate.slice(0,2) + '-' + compDate.slice(2,4) + '-' + compDate.slice(4,6);
    rows.push([chisu, jaejil, '완료', graveDate, compDate, '']);
  }

  if (rows.length > 0) {
    // 치수 순 정렬
    rows.sort((a, b) => a[0].localeCompare(b[0]));
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);

    // 상태별 색상
    for (let i = 0; i < rows.length; i++) {
      const range = sheet.getRange(i + 2, 1, 1, 6);
      if (rows[i][2] === '생산지시 중') {
        range.setBackground('#f1f5f9').setFontColor('#334155');
      } else {
        range.setBackground('#f0fdf4').setFontColor('#166534');
      }
    }
  }
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

    // 테이블 시트 동기화
    let grave = '{}', comp = '{}';
    try { grave = sheet.getRange('A1').getValue() || '{}'; } catch(err) {}
    try { comp  = sheet.getRange('A2').getValue() || '{}'; } catch(err) {}
    if (type === 'graveyard') grave = data;
    else comp = data;
    try { syncTable(grave, comp); } catch(err) {}

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
