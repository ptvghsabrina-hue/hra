const SHEET_ID = '15ikciWJUeILsCEUtwPMtVGoq44j7XAft12IQ5aLLJqs';
const SHEET_NAME = '點擊紀錄';
const PAGE = 'https://ptvghsabrina-hue.github.io/hra/';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const sh = getLogSheet_();
  sh.appendRow([new Date(), p.src || 'direct', p.cmp || '', p.vid || '', p.dev || '', p.ref || '']);
  return ContentService.createTextOutput('ok');
}

function getLogSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['時間', '來源', '活動', '訪客ID', '裝置', '上一頁']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 只需在編輯器執行一次：建立「統計」「推廣連結」工作表
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.rename('健康風險評估推廣點擊紀錄');
  getLogSheet_();

  const L = "'" + SHEET_NAME + "'";
  let st = ss.getSheetByName('統計') || ss.getSheets()[0];
  if (st.getName() === SHEET_NAME) st = ss.insertSheet('統計');
  st.setName('統計').clear();
  try { ss.setActiveSheet(st); ss.moveActiveSheet(1); } catch (e) {}

  st.getRange('A1:A3').setValues([['總點擊數'], ['不重複訪客（估）'], ['今天點擊數']]);
  st.getRange('B1').setFormula('=COUNTA(' + L + '!A2:A)');
  st.getRange('B2').setFormula('=COUNTUNIQUE(' + L + '!D2:D)');
  st.getRange('B3').setFormula('=COUNTIFS(' + L + '!A2:A,">="&TODAY())');

  st.getRange('D1').setValue('依來源');
  st.getRange('D2').setFormula('=IFERROR(QUERY(' + L + '!A:D,"select B, count(A) where A is not null group by B order by count(A) desc label B \'來源\', count(A) \'點擊數\'",1),"尚無資料")');

  const period = (fmt, label) => '=ARRAYFORMULA(LET(d,' + L + '!A2:A,v,' + L + '!D2:D,m,IF(d="","",TEXT(d,"' + fmt + '")),u,IFERROR(SORT(UNIQUE(FILTER(m,m<>"")),1,FALSE),),IF(INDEX(u,1)="",{"' + label + '","點擊數","不重複訪客";"尚無資料","",""},VSTACK({"' + label + '","點擊數","不重複訪客"},HSTACK(u,MAP(u,LAMBDA(x,SUM(--(m=x)))),MAP(u,LAMBDA(x,IFERROR(COUNTUNIQUE(FILTER(v,m=x,v<>"")),0))))))))';
  st.getRange('G1').setValue('每月');
  st.getRange('G2').setFormula(period('yyyy-mm', '月份'));
  st.getRange('K1').setValue('每年');
  st.getRange('K2').setFormula(period('yyyy', '年份'));
  st.getRange('1:2').setFontWeight('bold');

  const ln = ss.getSheetByName('推廣連結') || ss.insertSheet('推廣連結', 1);
  ln.clear();
  ln.getRange(1, 1, 7, 3).setValues([
    ['管道', '連結（發這條）', '說明'],
    ['Facebook', PAGE + '?src=facebook', 'FB 貼文、粉專'],
    ['LINE', PAGE + '?src=line', 'LINE 群組、官方帳號'],
    ['海報 QR Code', PAGE + '?src=poster', '用這條產生 QR Code'],
    ['院網', PAGE + '?src=hospital-web', '院網頁面上的連結'],
    ['電子報 / Email', PAGE + '?src=email', ''],
    ['新增管道', PAGE + '?src=自訂名稱&cmp=活動名', 'src 填管道名、cmp 填活動名（選填）；沒帶 src 的不會記錄']
  ]);
  ln.getRange('1:1').setFontWeight('bold');
  ln.autoResizeColumns(1, 3);
}
