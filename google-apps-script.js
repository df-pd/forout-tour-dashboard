/**
 * ============================================================
 * Google Apps Script — 網頁表單 POST 接收器
 * ============================================================
 *
 * 【部署步驟】
 *
 * 1. 建立 Google Sheets 試算表
 *    - 前往 https://sheets.google.com
 *    - 點選「空白試算表」建立新檔案
 *    - 在第一列（A1 ~ F1）填入欄位標題：
 *      公司名稱 | 聯絡人 | 電話 | Email | 需求描述 | 時間戳記
 *
 * 2. 開啟 Apps Script 編輯器
 *    - 在試算表上方選單點選「擴充功能」→「Apps Script」
 *    - 會開啟一個新的 Apps Script 專案頁面
 *
 * 3. 貼上程式碼
 *    - 將編輯器中預設的 myFunction 程式碼全部刪除
 *    - 把本檔案的所有程式碼複製貼上
 *    - 按下 Ctrl+S（或 Cmd+S）儲存
 *
 * 4. 部署為 Web App
 *    - 點選右上角「部署」→「新增部署作業」
 *    - 左側齒輪圖示選擇「網頁應用程式」
 *    - 「執行身分」選擇「我」（你自己的帳號）
 *    - 「誰可以存取」選擇「任何人」
 *    - 點選「部署」
 *    - 首次部署需要授權，按下「授權存取」並完成 Google 帳號驗證
 *
 * 5. 取得 Web App URL
 *    - 部署完成後會顯示一組 Web App URL
 *    - 複製該 URL，貼回你的 HTML 檔案中的 GOOGLE_SCRIPT_URL 變數
 *    - 範例：const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/xxxxx/exec';
 *
 * 【注意事項】
 *    - 每次修改程式碼後，需要建立「新版本」的部署才會生效
 *    - 若要更新部署，點選「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」
 *    - 試算表與 Apps Script 必須在同一個 Google 帳號下
 *
 * ============================================================
 */

/**
 * doGet — 處理 GET 請求（用於測試部署是否成功）
 * 瀏覽器直接開啟 Web App URL 時會觸發此函式
 */
function doGet(e) {
  var output = {
    status: 'ok',
    message: 'Google Apps Script 表單接收器運作正常',
    timestamp: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
  };

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost — 處理 POST 請求（接收網頁表單資料並寫入試算表）
 * 支援兩種格式：HTML 表單（e.parameter）和 JSON（e.postData.contents）
 * @param {Object} e - 事件物件，包含 parameter 或 postData 等資訊
 */
function doPost(e) {
  try {
    // 嘗試從 HTML 表單參數取得資料（優先），否則從 JSON body 解析
    var data;
    if (e.parameter && e.parameter.company) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      throw new Error('未收到任何資料');
    }

    // 取得目前的試算表，指定「諮詢表單」工作表
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('諮詢表單');
    if (!sheet) {
      sheet = ss.insertSheet('諮詢表單');
      sheet.appendRow(['公司名稱', '聯絡人', '電話', 'Email', '需求描述', '時間戳記']);
    }

    // 產生台灣時區的時間戳記
    var timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    // 將資料追加到工作表的最後一行
    sheet.appendRow([
      data.company  || '',
      data.name     || '',
      data.phone    || '',
      data.email    || '',
      data.needs    || '',
      timestamp
    ]);

    // 回傳成功的 HTML 頁面（供 iframe 顯示）
    return HtmlService.createHtmlOutput('<html><body><script>parent.postMessage("form-success","*");</script></body></html>');

  } catch (error) {
    return HtmlService.createHtmlOutput('<html><body><script>parent.postMessage("form-error","*");</script></body></html>');
  }
}
