/**
 * ============================================================
 * Google Apps Script — 統一 API 端點
 * ============================================================
 *
 * 功能：
 *   1. 諮詢表單接收（原有）
 *   2. 使用者登入驗證
 *   3. 預約參訪管理（新增/列表）
 *   4. 一次式預約連結（產生/驗證/公開預約）
 *   5. AI Chat（Gemini API + QA 知識庫）
 *
 * 【需要的工作表】
 *
 *   1.「諮詢表單」— 接收推廣區諮詢（自動建立）
 *      A: 公司名稱 | B: 聯絡人 | C: 電話 | D: Email | E: 需求描述 | F: 時間戳記
 *
 *   2.「帳號」— 內部登入帳號（請手動建立）
 *      A: 帳號 | B: 密碼 | C: 姓名 | D: 角色 | E: 可見功能
 *      例：admin | changeme123 | 管理員 | admin | （空白＝全部）
 *      例：user1 | pass123     | 王小明 | user  | booking,list
 *
 *      可見功能欄位值（以逗號分隔）：
 *        booking — 預約參訪
 *        list    — 預約清單
 *        stats   — 內部統計
 *        qa      — 內部 QA
 *      若角色為 admin 或可見功能為空，則顯示全部功能
 *
 *   3.「預約參訪」— 預約記錄（自動建立）
 *      A: 參訪日期 | B: 時段 | C: 公司名稱 | D: 聯絡人 | E: 電話
 *      F: Email | G: 人數 | H: 備註 | I: 預約人 | J: 建立時間
 *
 *   4.「預約連結」— 一次式預約連結（自動建立）
 *      A: 連結代碼 | B: 狀態(active/used/expired) | C: 建立人
 *      D: 建立時間 | E: 到期日 | F: 使用時間
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
// Gemini API 設定（備援，優先使用指令碼屬性）
// ============================================================
var GEMINI_API_KEY_ = 'AIzaSyAdyBtmMqj0BRosoGxdzfKxWBf_CBAGAzg';

// ============================================================
// Google 日曆設定
// ============================================================

// 自有日曆（可讀寫，預約會寫入此日曆）
const CALENDAR_ID = '1ed9d1eb22c38ff479bed67eec478366a350ae750f6ce63a3ae937601d904aea@group.calendar.google.com';

// 官方日曆（唯讀，僅讀取事件顯示在系統日曆）
const OFFICIAL_CALENDAR_ID = '71dc281c478614bd6b32fa271fb5427666009350524d8d7aeb2a3dcd00511094@group.calendar.google.com';

// ============================================================
// 試算表設定
// ============================================================
var SPREADSHEET_ID = '1SlGXwWgjjqoywFYx3nkE-7YolFfdLkK-Q1JuaT4Kt5Y';

/**
 * 取得試算表（優先用綁定的，備援用 ID 開啟）
 */
function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return ss;
}

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
  var ss = getSpreadsheet();
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
  var ss = getSpreadsheet();
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
// Google 日曆：建立參訪事件
// ============================================================

/**
 * 預約成功後，自動在 Google 日曆建立事件
 * @param {Object} data - 預約資料
 * @param {string} bookedBy - 預約人（內部帳號名稱或「客戶自行預約」）
 */
function createCalendarEvent(data, bookedBy) {
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) return; // 日曆不存在就跳過

    // 解析日期與時段
    var dateStr = String(data.visitDate); // 格式：yyyy-MM-dd
    var period = String(data.period || '');

    // 根據時段決定起訖時間
    var startHour = 9, endHour = 12;
    if (period.indexOf('下午') !== -1 || period.indexOf('13:00') !== -1 || period.indexOf('Afternoon') !== -1 || period.indexOf('午後') !== -1 || period.indexOf('오후') !== -1) {
      startHour = 13;
      endHour = 17;
    }

    var parts = dateStr.split('-');
    var startTime = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]), startHour, 0, 0);
    var endTime = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]), endHour, 0, 0);

    // 事件標題
    var title = '【導覽】' + (data.company || '未填公司') + '（' + (parseInt(data.people) || 0) + '人）';

    // 事件描述
    var desc = '公司/機構：' + (data.company || '') + '\n'
      + '聯絡人：' + (data.contact || '') + '\n'
      + '電話：' + (data.phone || '') + '\n'
      + 'Email：' + (data.email || '') + '\n'
      + '人數：' + (parseInt(data.people) || 0) + '\n'
      + '備註：' + (data.note || '') + '\n'
      + '預約人：' + bookedBy + '\n'
      + '建立時間：' + twTimestamp();

    cal.createEvent(title, startTime, endTime, {
      description: desc,
      location: '新北市資源教育基地'
    });
  } catch (e) {
    // 日曆建立失敗不影響預約流程，僅記錄錯誤
    Logger.log('日曆事件建立失敗：' + e.message);
  }
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
      case 'generateLink':
        return handleGenerateLink(data);
      case 'validateLink':
        return handleValidateLink(data);
      case 'publicBooking':
        return handlePublicBooking(data);
      case 'listCalendarEvents':
        return handleListCalendarEvents(data);
      case 'listLinks':
        return handleListLinks(data);
      case 'updateCredentials':
        return handleUpdateCredentials(data);
      case 'aiChat':
        return handleAiChat(data);
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

  var ss = getSpreadsheet();
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
      var role = String(rows[i][3]).trim() || 'user';
      var permStr = String(rows[i][4] || '').trim();

      // 解析可見功能：admin 角色或空白 → 全部
      var allTabs = ['booking', 'list', 'stats', 'qa'];
      var permissions;
      if (role === 'admin' || !permStr) {
        permissions = allTabs;
      } else {
        permissions = permStr.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return allTabs.indexOf(s) !== -1; });
        if (permissions.length === 0) permissions = allTabs;
      }

      return jsonResponse({
        success: true,
        token: token,
        user: {
          name: String(rows[i][2]).trim() || username,
          role: role
        },
        permissions: permissions
      });
    }
  }

  return jsonResponse({ success: false, error: '帳號或密碼錯誤' });
}

// ============================================================
// 功能：修改帳號密碼
// ============================================================

function handleUpdateCredentials(data) {
  // 驗證登入
  var user = verifyToken(data.token);
  if (!user) {
    return jsonResponse({ success: false, error: '請重新登入' });
  }

  var currentPassword = String(data.currentPassword || '').trim();
  var newUsername = String(data.newUsername || '').trim();
  var newPassword = String(data.newPassword || '').trim();

  if (!currentPassword) {
    return jsonResponse({ success: false, error: '請輸入目前密碼' });
  }
  if (!newUsername && !newPassword) {
    return jsonResponse({ success: false, error: '請輸入新帳號或新密碼' });
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('帳號');
  if (!sheet) {
    return jsonResponse({ success: false, error: '系統錯誤' });
  }

  var rows = sheet.getDataRange().getValues();
  var targetRow = -1;

  // 找到目前使用者並驗證密碼
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === user.username) {
      if (String(rows[i][1]).trim() !== currentPassword) {
        return jsonResponse({ success: false, error: '目前密碼不正確' });
      }
      targetRow = i + 1; // Sheet 列號（1-based）
      break;
    }
  }

  if (targetRow === -1) {
    return jsonResponse({ success: false, error: '找不到帳號' });
  }

  // 檢查新帳號是否已被使用
  if (newUsername && newUsername !== user.username) {
    for (var j = 1; j < rows.length; j++) {
      if (String(rows[j][0]).trim() === newUsername) {
        return jsonResponse({ success: false, error: '此帳號已被使用' });
      }
    }
    sheet.getRange(targetRow, 1).setValue(newUsername);
  }

  // 更新密碼
  if (newPassword) {
    sheet.getRange(targetRow, 2).setValue(newPassword);
  }

  return jsonResponse({ success: true });
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

  // 同步建立 Google 日曆事件
  createCalendarEvent(data, user.name);

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

  var ss = getSpreadsheet();
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
// 功能：取得日曆事件（自有 + 官方）
// ============================================================

function handleListCalendarEvents(data) {
  var user = verifyToken(data.token);
  if (!user) {
    return jsonResponse({ success: false, error: '請先登入' });
  }

  var year = parseInt(data.year) || new Date().getFullYear();
  var month = parseInt(data.month); // 1-based
  if (isNaN(month)) month = new Date().getMonth() + 1;

  var startDate = new Date(year, month - 1, 1, 0, 0, 0);
  var endDate = new Date(year, month, 0, 23, 59, 59); // 該月最後一天

  var events = [];

  // 讀取自有日曆事件
  try {
    var ownCal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (ownCal) {
      var ownEvents = ownCal.getEvents(startDate, endDate);
      ownEvents.forEach(function(ev) {
        events.push({
          title: ev.getTitle(),
          description: ev.getDescription() || '',
          location: ev.getLocation() || '',
          start: Utilities.formatDate(ev.getStartTime(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          end: Utilities.formatDate(ev.getEndTime(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          date: Utilities.formatDate(ev.getStartTime(), 'Asia/Taipei', 'yyyy-MM-dd'),
          source: 'own'
        });
      });
    }
  } catch (e) {
    Logger.log('讀取自有日曆失敗：' + e.message);
  }

  // 讀取官方日曆事件
  try {
    var officialCal = CalendarApp.getCalendarById(OFFICIAL_CALENDAR_ID);
    if (officialCal) {
      var officialEvents = officialCal.getEvents(startDate, endDate);
      officialEvents.forEach(function(ev) {
        events.push({
          title: ev.getTitle(),
          description: ev.getDescription() || '',
          location: ev.getLocation() || '',
          start: Utilities.formatDate(ev.getStartTime(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          end: Utilities.formatDate(ev.getEndTime(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          date: Utilities.formatDate(ev.getStartTime(), 'Asia/Taipei', 'yyyy-MM-dd'),
          source: 'official'
        });
      });
    }
  } catch (e) {
    Logger.log('讀取官方日曆失敗：' + e.message);
  }

  return jsonResponse({ success: true, events: events });
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

// ============================================================
// 功能：產生一次式預約連結（需登入）
// ============================================================

function handleGenerateLink(data) {
  var user = verifyToken(data.token);
  if (!user) {
    return jsonResponse({ success: false, error: '請先登入' });
  }

  var sheet = getOrCreateSheet('預約連結', [
    '連結代碼', '狀態', '建立人', '建立時間', '到期日', '使用時間'
  ]);

  // 產生 8 碼隨機代碼
  var code = generateRandomCode(8);

  // 到期日：7 天後
  var now = new Date();
  var expiry = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  var expiryStr = Utilities.formatDate(expiry, 'Asia/Taipei', 'yyyy-MM-dd');

  sheet.appendRow([
    code,
    'active',
    user.name,
    twTimestamp(),
    expiryStr,
    ''
  ]);

  return jsonResponse({
    success: true,
    code: code,
    expiry: expiryStr,
    message: '預約連結已產生'
  });
}

/**
 * 產生隨機英數代碼
 */
function generateRandomCode(length) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 去除易混淆字元
  var code = '';
  for (var i = 0; i < length; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

// ============================================================
// 功能：驗證預約連結（公開，無需登入）
// ============================================================

function handleValidateLink(data) {
  var code = String(data.code || '').trim().toUpperCase();
  if (!code) {
    return jsonResponse({ success: false, error: '請提供預約代碼' });
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('預約連結');
  if (!sheet) {
    return jsonResponse({ success: false, error: '預約代碼無效' });
  }

  var rows = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === code) {
      var status = String(rows[i][1]).trim();
      var expiry = String(rows[i][4]).trim();

      if (status === 'used') {
        return jsonResponse({ success: false, error: '此預約連結已使用過' });
      }
      if (status === 'expired' || today > expiry) {
        // 自動標記過期
        if (status !== 'expired') {
          sheet.getRange(i + 1, 2).setValue('expired');
        }
        return jsonResponse({ success: false, error: '此預約連結已過期' });
      }
      if (status === 'active') {
        return jsonResponse({
          success: true,
          valid: true,
          expiry: expiry,
          createdBy: String(rows[i][2]).trim()
        });
      }
    }
  }

  return jsonResponse({ success: false, error: '預約代碼無效' });
}

// ============================================================
// 功能：公開預約提交（透過一次式連結，無需登入）
// ============================================================

function handlePublicBooking(data) {
  var code = String(data.code || '').trim().toUpperCase();
  if (!code) {
    return jsonResponse({ success: false, error: '缺少預約代碼' });
  }

  // 驗證必填欄位
  if (!data.visitDate || !data.period || !data.company || !data.contact || !data.phone || !data.people) {
    return jsonResponse({ success: false, error: '請填寫所有必填欄位' });
  }

  // 驗證連結有效性
  var ss = getSpreadsheet();
  var linkSheet = ss.getSheetByName('預約連結');
  if (!linkSheet) {
    return jsonResponse({ success: false, error: '預約代碼無效' });
  }

  var rows = linkSheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  var linkRowIndex = -1;

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === code) {
      var status = String(rows[i][1]).trim();
      var expiry = String(rows[i][4]).trim();

      if (status !== 'active' || today > expiry) {
        return jsonResponse({ success: false, error: '此預約連結已失效' });
      }
      linkRowIndex = i + 1; // Sheet 行號（1-based）
      break;
    }
  }

  if (linkRowIndex === -1) {
    return jsonResponse({ success: false, error: '預約代碼無效' });
  }

  // 寫入預約資料
  var bookingSheet = getOrCreateSheet('預約參訪', [
    '參訪日期', '時段', '公司名稱', '聯絡人', '電話',
    'Email', '人數', '備註', '預約人', '建立時間'
  ]);

  bookingSheet.appendRow([
    data.visitDate,
    data.period,
    data.company,
    data.contact,
    data.phone,
    data.email || '',
    parseInt(data.people) || 0,
    data.note || '',
    '客戶自行預約 (' + code + ')',
    twTimestamp()
  ]);

  // 同步建立 Google 日曆事件
  createCalendarEvent(data, '客戶自行預約 (' + code + ')');

  // 將連結標記為已使用
  linkSheet.getRange(linkRowIndex, 2).setValue('used');
  linkSheet.getRange(linkRowIndex, 6).setValue(twTimestamp());

  return jsonResponse({ success: true, message: '預約成功！我們將盡快與您聯繫確認。' });
}

// ============================================================
// 功能：取得預約連結清單（需登入）
// ============================================================

function handleListLinks(data) {
  var user = verifyToken(data.token);
  if (!user) {
    return jsonResponse({ success: false, error: '請先登入' });
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('預約連結');

  if (!sheet) {
    return jsonResponse({ success: true, links: [] });
  }

  var rows = sheet.getDataRange().getValues();
  var links = [];
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  for (var i = 1; i < rows.length; i++) {
    var status = String(rows[i][1]).trim();
    var expiry = String(rows[i][4]).trim();

    // 自動更新過期狀態
    if (status === 'active' && today > expiry) {
      status = 'expired';
      sheet.getRange(i + 1, 2).setValue('expired');
    }

    links.push({
      code: String(rows[i][0]).trim(),
      status: status,
      createdBy: String(rows[i][2]).trim(),
      createdAt: String(rows[i][3]).trim(),
      expiry: expiry,
      usedAt: String(rows[i][5]).trim()
    });
  }

  return jsonResponse({ success: true, links: links });
}

// ============================================================
// 功能：AI Chat（Gemini API + QA 知識庫）
// ============================================================

/**
 * AI Chat — 使用 Gemini API 搭配 QA 知識庫回答問題
 *
 * 請求格式：
 *   { action: 'aiChat', message: '使用者問題', history: [{role, text}] }
 *
 * 設定步驟：
 *   1. 取得 Gemini API Key：https://aistudio.google.com/apikey
 *   2. 在 GAS 編輯器 → 專案設定 → 指令碼屬性 → 新增：
 *      屬性名稱: GEMINI_API_KEY
 *      值: 你的 API Key
 */

// QA 知識庫快取（同一次執行內共用）
var _qaKnowledgeCache = null;

/**
 * 從 QA 工作表載入知識庫
 * GID 45268038 對應的工作表
 */
function loadQAKnowledge() {
  if (_qaKnowledgeCache) return _qaKnowledgeCache;

  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var qaSheet = null;

  // 透過 GID 找到 QA 工作表
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === 45268038) {
      qaSheet = sheets[i];
      break;
    }
  }

  if (!qaSheet) {
    // 備援：嘗試找名稱含 QA 的工作表
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName().toLowerCase();
      if (name.indexOf('qa') !== -1 || name.indexOf('問答') !== -1) {
        qaSheet = sheets[i];
        break;
      }
    }
  }

  if (!qaSheet) {
    _qaKnowledgeCache = '';
    return '';
  }

  var data = qaSheet.getDataRange().getValues();
  if (data.length < 2) {
    _qaKnowledgeCache = '';
    return '';
  }

  // 欄位：題號(0), 日期(1), 單位(2), 人數(3), 問題類別(4), 問題(5), 參考題號(6), 回覆(7), 備註(8)
  var qaEntries = [];
  for (var i = 1; i < data.length; i++) {
    var question = String(data[i][5] || '').trim();
    var answer = String(data[i][7] || '').trim();
    var category = String(data[i][4] || '').trim();

    if (question && answer) {
      qaEntries.push('[' + category + '] Q: ' + question + '\nA: ' + answer);
    }
  }

  _qaKnowledgeCache = qaEntries.join('\n\n');
  return _qaKnowledgeCache;
}

function handleAiChat(data) {
  var userMessage = String(data.message || '').trim();
  if (!userMessage) {
    return jsonResponse({ success: false, error: '請輸入問題' });
  }

  // 取得 Gemini API Key（優先從指令碼屬性讀取，備援用常數）
  var apiKey = '';
  try {
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
  } catch (e) {
    Logger.log('讀取 Script Properties 失敗：' + e.message);
  }
  if (!apiKey) {
    apiKey = GEMINI_API_KEY_ || '';
  }
  if (!apiKey) {
    return jsonResponse({
      success: false,
      error: 'AI 功能尚未設定，請聯繫管理員設定 GEMINI_API_KEY'
    });
  }

  // 載入 QA 知識庫
  var knowledge = loadQAKnowledge();

  // 組裝系統提示詞
  var systemPrompt = '你是「新北市五股資源循環教育基地」的智慧導覽助手。\n' +
    '本基地由大豐環保科技股份有限公司營運，負責五股、三重、蘆洲三區的資源回收物細分選作業。\n\n' +
    '回答規則：\n' +
    '1. 優先根據下方知識庫中的 QA 資料回答\n' +
    '2. 若知識庫沒有直接答案，可根據知識庫內容推理，但須註明「根據現有資料推測」\n' +
    '3. 若完全無法回答，請建議使用者撥打導覽專線或現場詢問\n' +
    '4. 使用繁體中文回答，語氣親切專業\n' +
    '5. 回答簡潔扼要，不超過 300 字\n\n' +
    '===== 知識庫 =====\n' +
    (knowledge || '（知識庫目前為空）');

  // 組裝對話歷史
  var contents = [];

  // 加入歷史訊息（最多保留最近 10 輪）
  var history = data.history || [];
  var startIdx = Math.max(0, history.length - 20);  // 最多 20 則（10 輪來回）
  for (var i = startIdx; i < history.length; i++) {
    var h = history[i];
    contents.push({
      role: h.role === 'user' ? 'user' : 'model',
      parts: [{ text: String(h.text || '') }]
    });
  }

  // 加入當前使用者訊息
  contents.push({
    role: 'user',
    parts: [{ text: userMessage }]
  });

  // 呼叫 Gemini API
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=' + apiKey;

  var payload = {
    system_instruction: {
      parts: [{ text: systemPrompt }]
    },
    contents: contents,
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 512,
      topP: 0.8,
    },
    safetySettings: [
      { category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_ONLY_HIGH' },
      { category: 'HARM_CATEGORY_HATE_SPEECH', threshold: 'BLOCK_ONLY_HIGH' },
      { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_ONLY_HIGH' },
      { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_ONLY_HIGH' }
    ]
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode !== 200) {
      var errMsg = (result.error && result.error.message) || '呼叫 AI 失敗';
      Logger.log('Gemini API 錯誤：' + statusCode + ' - ' + errMsg);
      return jsonResponse({ success: false, error: 'AI 錯誤（' + statusCode + '）：' + errMsg });
    }

    // 解析回應
    var aiText = '';
    if (result.candidates && result.candidates.length > 0) {
      var candidate = result.candidates[0];
      if (candidate.content && candidate.content.parts) {
        for (var j = 0; j < candidate.content.parts.length; j++) {
          aiText += candidate.content.parts[j].text || '';
        }
      }
    }

    if (!aiText) {
      return jsonResponse({ success: false, error: 'AI 無法產生回應' });
    }

    return jsonResponse({ success: true, reply: aiText });

  } catch (err) {
    Logger.log('AI Chat 錯誤：' + err.message);
    return jsonResponse({ success: false, error: 'AI 服務連線失敗，請稍後再試' });
  }
}

// ============================================================
// 測試：手動執行以觸發日曆授權
// ============================================================

/**
 * 在 GAS 編輯器中手動執行此函式，會跳出授權對話框。
 * 授權完成後，日曆功能才能正常運作。
 * 只需執行一次。
 */
function testCalendarAccess() {
  // 測試自有日曆
  var ownCal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (ownCal) {
    Logger.log('✅ 自有日曆連線成功：' + ownCal.getName());
  } else {
    Logger.log('❌ 自有日曆無法存取，請確認 ID 和權限');
  }

  // 測試官方日曆
  var officialCal = CalendarApp.getCalendarById(OFFICIAL_CALENDAR_ID);
  if (officialCal) {
    Logger.log('✅ 官方日曆連線成功：' + officialCal.getName());
  } else {
    Logger.log('❌ 官方日曆無法存取，請確認 ID 和權限');
  }
}
