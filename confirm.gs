/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file
 * access for this script to only the current document.
 */

// 試算表工作表的名稱
const SHEET_NAME = "wordwall授權確認"; // 請確認您的工作表名稱是否正確

/**
 * 當使用者訪問網頁應用程式時，執行此函數來提供 HTML 頁面。
 * @param {Object} e - 事件參數。
 * @returns {HtmlOutput} - 從 HTML 檔案建立的 HTML 輸出物件。
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Wordwall 授權帳號確認系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * 根據使用者輸入的姓名，在試算表中尋找對應的授權帳號。
 * @param {string} name - 使用者輸入的姓名。
 * @returns {Object|null} - 如果找到，返回包含帳號和列號的物件；否則返回 null。
 */
function findAccount(name) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`找不到名稱為 "${SHEET_NAME}" 的工作表。`);
    }
    // 假設 A 欄是姓名，B 欄是帳號
    const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === name) {
        return {
          account: data[i][1],
          row: i + 2 // +2 是因為範圍從第 2 列開始，且陣列索引從 0 開始
        };
      }
    }
    return null; // 找不到姓名
  } catch (error) {
    Logger.log(error.toString());
    return null;
  }
}

/**
 * 更新試算表中的確認狀態或帳號資料，並寄送通知信。
 * @param {string} name - 要更新資料的姓名。
 * @param {string} status - 更新的狀態 ('correct' 或 'modify')。
 * @param {string|null} newEmail - 如果需要修改，提供新的電子郵件地址。
 * @returns {string} - 返回操作結果的訊息。
 */
function updateStatus(name, status, newEmail) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
     if (!sheet) {
      throw new Error(`找不到名稱為 "${SHEET_NAME}" 的工作表。`);
    }
    // 讀取姓名和 Email 欄位
    const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
    let rowToUpdate = -1;
    let originalEmail = '';

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === name) {
        rowToUpdate = i + 2;
        originalEmail = data[i][1]; // 儲存原始 Email
        break;
      }
    }

    if (rowToUpdate === -1) {
      return "找不到您的姓名，無法更新。";
    }

    const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    
    if (status === 'correct') {
      // 在 C 欄寫入 "已確認" 和時間戳
      sheet.getRange(rowToUpdate, 3).setValue('已確認 - ' + timestamp);

      // 寄送確認信
      const recipient = originalEmail;
      const subject = "Wordwall 授權帳號確認通知";
      const body = `${name} 您好，\n\n感謝您的確認，您用於登入 Wordwall 的授權帳號 (${recipient}) 已確認無誤。\n\n此為系統自動發送郵件，請勿直接回覆。`;
      GmailApp.sendEmail(recipient, subject, body);
      
      return "感謝您的確認！資料已成功記錄，一封確認信已寄至您的信箱。";

    } else if (status === 'modify' && newEmail) {
      // 在 B 欄更新 Email，在 C 欄寫入 "需修改" 和時間戳
      sheet.getRange(rowToUpdate, 2).setValue(newEmail);
      sheet.getRange(rowToUpdate, 3).setValue('已修改 - ' + timestamp);

      // 寄送修改通知信
      const recipient = newEmail;
      const subject = "Wordwall 授權帳號修改通知";
      const body = `${name} 您好，\n\n您用於登入 Wordwall 的授權帳號已成功更新為：${recipient}\n\n此為系統自動發送郵件，請勿直接回覆。`;
      GmailApp.sendEmail(recipient, subject, body);

      return "您的帳號已成功更新！一封確認信已寄至新的信箱。";
    } else {
      return "操作無效，請重試。";
    }

  } catch (error) {
    Logger.log(error.toString());
    return "發生錯誤：" + error.toString();
  }
}
