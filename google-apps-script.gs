// ══════════════════════════════════════════════════════════
//  照顧服務員術科培訓研討會 — 簽到系統 Google Apps Script
//  使用說明：
//  1. 開啟 Google Sheets → 擴充功能 → Apps Script
//  2. 貼上此程式碼
//  3. 點「部署」→「新增部署作業」→ 類型選「網路應用程式」
//  4. 執行身分：「我」；存取權限：「所有人（包括匿名）」
//  5. 授權後取得 URL，填入 index.html 的 CONFIG.scriptUrl
// ══════════════════════════════════════════════════════════

const SHEET_NAME = '簽到記錄';
const SIG_FOLDER_NAME = '簽到簽名圖片';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 取得或建立工作表
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['#', '簽到時間', '姓名', '服務單位', '職稱', '聯絡電話', '身分證後4碼', '活動名稱', '簽名圖片連結']);
      // 凍結標題列
      sheet.setFrozenRows(1);
      // 設定標題樣式
      const hdr = sheet.getRange(1, 1, 1, 9);
      hdr.setBackground('#2d6a50');
      hdr.setFontColor('#ffffff');
      hdr.setFontWeight('bold');
    }

    const lastRow = sheet.getLastRow();
    const rowNum = lastRow; // 扣掉標題列

    // 儲存簽名圖片到 Google Drive
    let sigUrl = '';
    if (data.signature && data.signature.startsWith('data:image/png;base64,')) {
      sigUrl = saveSigImage(data.signature, data.name, rowNum);
    }

    // 寫入一列
    sheet.appendRow([
      rowNum,
      data.time || new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      data.name || '',
      data.org || '',
      data.jobTitle || '',
      data.phone || '',
      data.idLast4 || '',
      data.event || '',
      sigUrl,
    ]);

    // 設定剛寫入那列的格式
    const newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, 9).setVerticalAlignment('middle');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', row: newRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSigImage(dataUrl, name, idx) {
  try {
    const base64 = dataUrl.replace('data:image/png;base64,', '');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', `sig_${idx}_${name}.png`);

    // 取得或建立資料夾
    const folders = DriveApp.getFoldersByName(SIG_FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(SIG_FOLDER_NAME);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    console.warn('儲存簽名圖片失敗：', err);
    return '';
  }
}

// GET 請求：可用來測試部署是否成功
function doGet(e) {
  return ContentService
    .createTextOutput('簽到系統運作正常 ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}
