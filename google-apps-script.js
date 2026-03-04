/**
 * ============================================================
 * Google Apps Script — 統一 API 端點
 * ============================================================
 *
 * 功能：
 *   1. 諮詢表單接收（原有）
 *   2. 使用者登入驗證
 *   3. 預約參訪管理（新增/列表）
 *
 * 【需要的工作表】
 *
 *   1.「諮詢表單」— 接收推廣區諮詢（自動建立）
 *      A: 公司名稱 | B: 聯絡人 | C: 電話 | D: Email | E: 需求描述 | F: 時間戳記
 *
 *   2.「帳號」— 內部登入帳號（請手動建立）
 *      A: 帳號 | B: 密碼 | C: 姓名 | D: 角色
 *      例：admin | changeme123 | 管理員 | admin
 *
 *   3.「預約參訪」— 預約記錄（自動建立）
 *      A: 參訪日期 | B: 時段 | C: 公司名稱 | D: 聯絡人 | E: 電話
 *      F: Email | G: 人數 | H: 備註 | I: 預約人 | J: 建立時間
 *
 * 【部署步驟】
 *   1. 在試算表「擴充功能」→「Apps Script」
 *   2. 貼上此程式碼
 *   3.「部署」→「新增部署作業」→ 網頁應用程式
 *      - 執行身分：我
 *      - 誰可以存取：所有已登入 Google 帳戶的使用者
 *   4. 複製 Web App URL 貼回 HTML
 *
 * 【更新部署】
 *   「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」→ 部署
 *
 * ============================================================
 */

// ============================================================
// 工具函式
// ============================================================

/**
 * 簡易 hash 函式（用於產生 token）
 * 非加密等級，但足夠做前端驗證 token
 */
function simpleHash(str) {
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    var char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // 轉為 32bit 整數
  }
  // 轉為正數的 hex 字串
  return (hash >>> 0).toString(16) + '-' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    str
  ).map(function(b) { return (b + 128).toString(16).slice(-2); }).join('').slice(0, 16);
}

/**
 * 取得台灣時區時間戳記
 */
function twTimestamp() {
  return new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
}

/**
 * 回傳 JSON 回應
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 取得或建立工作表
 */
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * 驗證 token 是否有效
 */
function verifyToken(token) {
  if (!token) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('帳號');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  for (var i = 1; i < data.length; i++) {
    var username = String(data[i][0]).trim();
    var password = String(data[i][1]).trim();
    var expectedToken = simpleHash(username + ':' + password + ':' + today);
    if (token === expectedToken) {
      return { username: username, name: String(data[i][2]).trim(), role: String(data[i][3]).trim() };
    }
  }
  return null;
}

// ============================================================
// GET 請求處理
// ============================================================

function doGet(e) {
  return jsonResponse({
    status: 'ok',
    message: 'Google Apps Script API 運作正常',
    timestamp: twTimestamp()
  });
}

// ============================================================
// POST 請求處理（路由式）
// ============================================================

function doPost(e) {
  try {
    // 若是 HTML 表單提交（推廣區 iframe 方式），走舊路徑
    if (e.parameter && e.parameter.company) {
      return handleContactForm(e.parameter, true);
    }

    // JSON 請求 → 依 action 路由
    if (!e.postData || !e.postData.contents) {
      throw new Error('未收到任何資料');
    }

    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'contact';

    switch (action) {
      case 'login':
        return handleLogin(data);
      case 'booking':
        return handleBooking(data);
      case 'listBookings':
        return handleListBookings(data);
      case 'contact':
        return handleContactForm(data, false);
      default:
        // 向下相容：無 action 欄位視為諮詢表單
        return handleContactForm(data, false);
    }

  } catch (error) {
    // 若是 iframe 模式，回傳 HTML
    if (e.parameter && e.parameter.company) {
      return HtmlService.createHtmlOutput(
        '<html><body><script>parent.postMessage("form-error","*");</script></body></html>'
      );
    }
    return jsonResponse({ success: false, error: error.message || '未知錯誤' });
  }
}

// ============================================================
// 功能：登入驗證
// ============================================================

function handleLogin(data) {
  var username = String(data.username || '').trim();
  var password = String(data.password || '').trim();

  if (!username || !password) {
    return jsonResponse({ success: false, error: '請輸入帳號和密碼' });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('帳號');

  if (!sheet) {
    return jsonResponse({ success: false, error: '系統尚未設定帳號，請聯繫管理員' });
  }

  var rows = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  for (var i = 1; i < rows.length; i++) {
    var rowUser = String(rows[i][0]).trim();
    var rowPass = String(rows[i][1]).trim();

    if (rowUser === username && rowPass === password) {
      // 產生當日有效的 token
      var token = simpleHash(username + ':' + password + ':' + today);

      return jsonResponse({
        success: true,
        token: token,
        user: {
          name: String(rows[i][2]).trim() || username,
          role: String(rows[i][3]).trim() || 'user'
        }
      });
    }
  }

  return jsonResponse({ success: false, error: '帳號或密碼錯誤' });
}

// ============================================================
// 功能：新增預約參訪
// ============================================================

function handleBooking(data) {
  // 驗證登入
  var user = verifyToken(data.token);
  if (!user) {
    return jsonResponse({ success: false, error: '請先登入' });
  }

  // 驗證必填欄位
  if (!data.visitDate || !data.period || !data.company || !data.contact || !data.phone || !data.people) {
    return jsonResponse({ success: false, error: '請填寫所有必填欄位' });
  }

  var sheet = getOrCreateSheet('預約參訪', [
    '參訪日期', '時段', '公司名稱', '聯絡人', '電話',
    'Email', '人數', '備註', '預約人', '建立時間'
  ]);

  sheet.appendRow([
    data.visitDate,
    data.period,
    data.company,
    data.contact,
    data.phone,
    data.email || '',
    parseInt(data.people) || 0,
    data.note || '',
    user.name,
    twTimestamp()
  ]);

  return jsonResponse({ success: true, message: '預約成功' });
}

// ============================================================
// 功能：取得預約清單
// ============================================================

function handleListBookings(data) {
  // 驗證登入
  var user = verifyToken(data.token);
  if (!user) {
    return jsonResponse({ success: false, error: '請先登入' });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('預約參訪');

  if (!sheet) {
    return jsonResponse({ success: true, bookings: [] });
  }

  var rows = sheet.getDataRange().getValues();
  var bookings = [];

  for (var i = 1; i < rows.length; i++) {
    bookings.push({
      visitDate: String(rows[i][0]),
      period: String(rows[i][1]),
      company: String(rows[i][2]),
      contact: String(rows[i][3]),
      phone: String(rows[i][4]),
      email: String(rows[i][5]),
      people: parseInt(rows[i][6]) || 0,
      note: String(rows[i][7]),
      bookedBy: String(rows[i][8]),
      createdAt: String(rows[i][9])
    });
  }

  return jsonResponse({ success: true, bookings: bookings });
}

// ============================================================
// 功能：諮詢表單（原有，保持相容）
// ============================================================

function handleContactForm(data, isIframe) {
  var sheet = getOrCreateSheet('諮詢表單', [
    '公司名稱', '聯絡人', '電話', 'Email', '需求描述', '時間戳記'
  ]);

  sheet.appendRow([
    data.company  || '',
    data.name     || '',
    data.phone    || '',
    data.email    || '',
    data.needs    || '',
    twTimestamp()
  ]);

  if (isIframe) {
    return HtmlService.createHtmlOutput(
      '<html><body><script>parent.postMessage("form-success","*");</script></body></html>'
    );
  }

  return jsonResponse({ success: true, message: '諮詢已送出' });
}
