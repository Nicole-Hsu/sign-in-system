// ══════════════════════════════════════════════════════════
//  照顧服務員術科培訓研討會 — 簽到系統 Google Apps Script
//
//  接收方式：GET 請求（URL 參數），使用 Image 技巧繞過 CORS
//  簽名圖片：僅存於本地 admin.html，不上傳至此
//
//  部署設定：
//    執行身分：我
//    存取權限：所有人（包括匿名）
// ══════════════════════════════════════════════════════════

const SHEET_NAME = '簽到記錄';

function doGet(e) {
  // 有資料參數才寫入（區別一般瀏覽測試）
  if (e.parameter && e.parameter.name) {
    try {
      saveRecord(e.parameter);
    } catch(err) {
      // 寫入失敗也回傳 ok，避免前端報錯
      console.error('寫入失敗：', err);
    }
  }
  // 一定要回傳內容，否則 Image 請求會報錯（不影響功能但較乾淨）
  return ContentService
    .createTextOutput('ok')
    .setMimeType(ContentService.MimeType.TEXT);
}

function saveRecord(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 取得或建立工作表
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['#', '簽到時間', '姓名', '服務單位', '職稱', '聯絡電話', '身分證後4碼', '活動名稱'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setBackground('#2d6a50');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    sheet.setColumnWidth(1, 40);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 160);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(8, 200);
  }

  const rowNum = sheet.getLastRow(); // 扣掉標題列的序號
  const time   = p.time || new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

  sheet.appendRow([
    rowNum,
    time,
    p.name     || '',
    p.org      || '',
    p.jobTitle || '',
    p.phone    || '',
    p.idLast4  || '',
    p.event    || '',
  ]);
}

// POST 備用（保留，未來擴充用）
function doPost(e) {
  try {
    const raw = (e.postData && e.postData.contents)
      ? e.postData.contents
      : (e.parameter && e.parameter.data ? e.parameter.data : '{}');
    const data = JSON.parse(raw);
    if (data.name) saveRecord(data);
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    return ContentService.createTextOutput('error: ' + err).setMimeType(ContentService.MimeType.TEXT);
  }
}
