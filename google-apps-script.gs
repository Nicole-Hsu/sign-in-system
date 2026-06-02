// ══════════════════════════════════════════════════════════
//  照顧服務員術科培訓研討會 — 簽到系統 Google Apps Script
//
//  接收方式：隱藏 iframe 表單 POST（e.parameter.data）
//  簽名圖片：自動存入 Google Drive，Sheets 記錄連結
//
//  部署設定：
//    執行身分：我
//    存取權限：所有人（包括匿名）
// ══════════════════════════════════════════════════════════

const SHEET_NAME = '簽到記錄';
const SIG_FOLDER_NAME = '簽到簽名圖片';

function doPost(e) {
  try {
    // 表單 POST 的資料在 e.parameter.data（JSON 字串）
    const raw = (e.parameter && e.parameter.data)
      ? e.parameter.data
      : (e.postData && e.postData.contents ? e.postData.contents : '{}');
    const data = JSON.parse(raw);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 取得或建立工作表
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = ['#', '簽到時間', '姓名', '服務單位', '職稱', '聯絡電話', '身分證後4碼', '活動名稱', '簽名圖片連結'];
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
      sheet.setColumnWidth(9, 300);
    }

    const rowNum = sheet.getLastRow(); // 流水號
    const time = data.time || new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    // 儲存簽名圖片到 Google Drive
    let sigUrl = '';
    if (data.signature && data.signature.startsWith('data:image/png;base64,')) {
      sigUrl = saveSigImage(data.signature, data.name || 'unknown', rowNum);
    }

    // 寫入試算表
    sheet.appendRow([
      rowNum,
      time,
      data.name     || '',
      data.org      || '',
      data.jobTitle || '',
      data.phone    || '',
      data.idLast4  || '',
      data.event    || '',
      sigUrl,
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error('doPost 錯誤：', err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSigImage(dataUrl, name, idx) {
  try {
    const base64 = dataUrl.replace('data:image/png;base64,', '');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      'image/png',
      `sig_${idx}_${name}.png`
    );
    const folders = DriveApp.getFoldersByName(SIG_FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(SIG_FOLDER_NAME);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    console.error('saveSigImage 錯誤：', err);
    return '';
  }
}

// GET：回傳 ok 供測試用
function doGet(e) {
  return ContentService
    .createTextOutput('簽到系統運作正常 ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}
