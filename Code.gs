const SHEET_NAME = '點擊紀錄';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['時間', '來源', '活動', '訪客ID', '裝置', '上一頁']);
  }
  sh.appendRow([new Date(), p.src || 'direct', p.cmp || '', p.vid || '', p.dev || '', p.ref || '']);
  return ContentService.createTextOutput('ok');
}
