/**
 * ============================================================
 * 新北市資源教育基地 - LINE@ Webhook → Google Sheets
 * ============================================================
 *
 * 功能：接收 LINE 訊息，解析導覽資訊，寫入 Google Sheets
 *
 * 使用方式：
 * 1. 開啟 https://script.google.com/
 * 2. 建立新專案，貼上此程式碼
 * 3. 修改下方 CONFIG 設定
 * 4. 部署為「網頁應用程式」→ 取得 URL
 * 5. 將 URL 貼到 LINE Developers 後台的 Webhook URL
 *
 * 訊息格式（自由文字，需包含以下資訊）：
 *   日期 / 名稱 / 單位 / 人數
 *   例如：「3/5 王老師 新北市OO國小 35人」
 */

// ============================================================
// 設定區 - 請填入你的資訊
// ============================================================
const CONFIG = {
  // LINE Messaging API（從 LINE Developers 後台取得）
  LINE_CHANNEL_ACCESS_TOKEN: '在此填入你的 Channel Access Token',
  LINE_CHANNEL_SECRET: '在此填入你的 Channel Secret',

  // Google Sheets（你的試算表 ID）
  SHEET_ID: '1SlGXwWgjjqoywFYx3nkE-7YolFfdLkK-Q1JuaT4Kt5Y',

  // 導覽記錄工作表名稱（需與你的 Google Sheets 工作表名稱一致）
  TOUR_SHEET_NAME: '工作表1',
};

// ============================================================
// LINE Webhook 端點
// ============================================================

/**
 * POST 端點 - 接收 LINE Webhook 事件
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.events && data.events.length > 0) {
      data.events.forEach(function(event) {
        if (event.type === 'message' && event.message.type === 'text') {
          handleTextMessage(event);
        }
      });
    }

    return ContentService.createTextOutput('OK');
  } catch (error) {
    logError('doPost', error);
    return ContentService.createTextOutput('ERROR');
  }
}

/**
 * GET 端點 - Webhook URL 驗證用
 */
function doGet(e) {
  return ContentService.createTextOutput('Webhook is active');
}

// ============================================================
// 訊息處理
// ============================================================

/**
 * 處理文字訊息
 * 嘗試解析導覽資訊（日期、名稱、單位、人數）
 */
function handleTextMessage(event) {
  const text = event.message.text.trim();
  const userId = event.source.userId;
  const userName = getUserDisplayName(userId);

  // 嘗試解析導覽資訊
  const parsed = parseTourInfo(text);

  if (parsed) {
    // 解析成功 → 新增導覽記錄
    const sessionNum = getNextSessionNumber();
    appendTourRecord({
      session: sessionNum,
      date: parsed.date,
      name: parsed.name,
      unit: parsed.unit,
      people: parsed.people,
    });

    // 回覆確認訊息
    const msg = `✓ 已新增導覽記錄\n`
      + `場次：${sessionNum}\n`
      + `日期：${parsed.date}\n`
      + `名稱：${parsed.name}\n`
      + `單位：${parsed.unit}\n`
      + `人數：${parsed.people}`;

    replyMessage(event.replyToken, msg);
  } else {
    // 解析失敗 → 提示格式
    replyMessage(event.replyToken,
      '無法解析導覽資訊，請確認訊息包含：\n'
      + '日期、名稱、單位、人數\n\n'
      + '範例：\n'
      + '3/5 王老師 新北市OO國小 35人\n'
      + '或：2026/3/5, 王老師, OO國小, 35'
    );
  }
}

// ============================================================
// 導覽資訊解析器
// ============================================================

/**
 * 解析自由格式文字中的導覽資訊
 * 支援多種格式：
 *   - 「3/5 王老師 OO國小 35人」
 *   - 「2026/3/5, 王老師, OO國小, 35」
 *   - 「日期：3/5 名稱：王老師 單位：OO國小 人數：35」
 */
function parseTourInfo(text) {
  // 移除多餘空白
  text = text.replace(/\s+/g, ' ').trim();

  let date = null, name = null, unit = null, people = null;

  // 策略 1：帶標籤格式（日期：xxx 名稱：xxx ...）
  const labelPatterns = {
    date: /(?:日期|date)[：:]\s*([^\s,，]+)/i,
    name: /(?:名稱|姓名|name|聯絡人|導覽員)[：:]\s*([^\s,，]+)/i,
    unit: /(?:單位|學校|機關|機構|unit|org)[：:]\s*([^\s,，]+)/i,
    people: /(?:人數|人次|實際人數|people|人員)[：:]\s*(\d+)/i,
  };

  const labelDate = text.match(labelPatterns.date);
  const labelName = text.match(labelPatterns.name);
  const labelUnit = text.match(labelPatterns.unit);
  const labelPeople = text.match(labelPatterns.people);

  if (labelDate && labelName && labelUnit && labelPeople) {
    return {
      date: normalizeDate(labelDate[1]),
      name: labelName[1],
      unit: labelUnit[1],
      people: parseInt(labelPeople[1]),
    };
  }

  // 策略 2：逗號/斜線分隔（日期, 名稱, 單位, 數字）
  const parts = text.split(/[,，\/\s]+/).map(s => s.trim()).filter(Boolean);

  // 從 parts 中識別各欄位
  date = extractDate(parts);
  people = extractPeople(parts);

  if (date !== null && people !== null) {
    // 移除已識別的日期和人數，剩下的是名稱和單位
    const remaining = parts.filter(p => {
      if (p === date.raw) return false;
      if (p === people.raw) return false;
      return true;
    });

    if (remaining.length >= 2) {
      // 判斷哪個是名稱、哪個是單位
      // 包含「校」「園」「局」「處」「中心」「公司」「機構」等的是單位
      const unitKeywords = /[校園局處所中心公司機構會部院署站廠場館]/;
      let nameStr = remaining[0];
      let unitStr = remaining[1];

      // 如果相反了就互換
      if (unitKeywords.test(remaining[0]) && !unitKeywords.test(remaining[1])) {
        unitStr = remaining[0];
        nameStr = remaining[1];
      } else if (!unitKeywords.test(remaining[0]) && unitKeywords.test(remaining[1])) {
        nameStr = remaining[0];
        unitStr = remaining[1];
      }
      // 如果兩邊都有或都沒有關鍵字，保持原順序（名稱在前）

      return {
        date: date.normalized,
        name: nameStr,
        unit: unitStr,
        people: people.value,
      };
    } else if (remaining.length === 1) {
      // 只剩一個，當作單位
      return {
        date: date.normalized,
        name: '（未提供）',
        unit: remaining[0],
        people: people.value,
      };
    }
  }

  return null; // 無法解析
}

/**
 * 從字串片段中提取日期
 */
function extractDate(parts) {
  for (const p of parts) {
    // yyyy/MM/dd 或 yyyy-MM-dd
    let m = p.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) return { raw: p, normalized: `${m[1]}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}` };

    // MM/dd 或 M/d
    m = p.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
      const year = new Date().getFullYear();
      return { raw: p, normalized: `${year}/${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}` };
    }

    // X月X日
    m = p.match(/^(\d{1,2})月(\d{1,2})日?$/);
    if (m) {
      const year = new Date().getFullYear();
      return { raw: p, normalized: `${year}/${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}` };
    }
  }
  return null;
}

/**
 * 從字串片段中提取人數
 */
function extractPeople(parts) {
  for (const p of parts) {
    // 「35人」或「35」
    const m = p.match(/^(\d+)\s*人?$/);
    if (m) {
      const val = parseInt(m[1]);
      // 人數通常在 1-500 之間，排除可能是日期的數字
      if (val >= 1 && val <= 2000) {
        return { raw: p, value: val };
      }
    }
  }
  return null;
}

/**
 * 標準化日期格式
 */
function normalizeDate(str) {
  // yyyy/MM/dd
  let m = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return `${m[1]}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}`;

  // MM/dd
  m = str.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const year = new Date().getFullYear();
    return `${year}/${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}`;
  }

  // X月X日
  m = str.match(/(\d{1,2})月(\d{1,2})日?/);
  if (m) {
    const year = new Date().getFullYear();
    return `${year}/${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}`;
  }

  return str;
}

// ============================================================
// Google Sheets 操作
// ============================================================

/**
 * 取得下一個場次編號
 */
function getNextSessionNumber() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1; // 只有標題列

  // 讀取最後一列的場次欄
  const lastSession = sheet.getRange(lastRow, 1).getValue();
  const num = parseInt(lastSession);
  return isNaN(num) ? lastRow : num + 1;
}

/**
 * 新增導覽記錄到 Google Sheets
 */
function appendTourRecord(data) {
  try {
    const sheet = getOrCreateSheet();
    sheet.appendRow([
      data.session,
      data.date,
      data.name,
      data.unit,
      data.people,
    ]);
  } catch (error) {
    logError('appendTourRecord', error);
  }
}

/**
 * 取得或建立導覽記錄工作表
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.TOUR_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TOUR_SHEET_NAME);

    // 建立標題列
    sheet.appendRow(['場次', '日期', '名稱', '單位', '實際人數']);

    // 凍結標題列
    sheet.setFrozenRows(1);

    // 設定標題樣式
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground('#1E40AF');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // 設定欄寬
    sheet.setColumnWidth(1, 60);   // 場次
    sheet.setColumnWidth(2, 120);  // 日期
    sheet.setColumnWidth(3, 120);  // 名稱
    sheet.setColumnWidth(4, 180);  // 單位
    sheet.setColumnWidth(5, 100);  // 實際人數
  }

  return sheet;
}

// ============================================================
// LINE Messaging API
// ============================================================

/**
 * 回覆訊息
 */
function replyMessage(replyToken, text) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = {
    replyToken: replyToken,
    messages: [{
      type: 'text',
      text: text
    }]
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

/**
 * 取得使用者名稱（帶快取）
 */
const USER_NAME_CACHE = {};
function getUserDisplayName(userId) {
  if (USER_NAME_CACHE[userId]) return USER_NAME_CACHE[userId];

  try {
    const url = `https://api.line.me/v2/bot/profile/${userId}`;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() === 200) {
      const profile = JSON.parse(res.getContentText());
      USER_NAME_CACHE[userId] = profile.displayName;
      return profile.displayName;
    }
  } catch (e) {
    // 取得名稱失敗，使用預設值
  }

  return '未知使用者';
}

// ============================================================
// 錯誤日誌
// ============================================================
function logError(funcName, error) {
  console.error(`[${funcName}] ${error.message || error}`);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let logSheet = ss.getSheetByName('錯誤日誌');
    if (!logSheet) {
      logSheet = ss.insertSheet('錯誤日誌');
      logSheet.appendRow(['時間', '函式', '錯誤訊息']);
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      funcName,
      error.message || String(error),
    ]);
  } catch (e) {
    // 寫入日誌也失敗了，只能忽略
  }
}

// ============================================================
// 初始化（手動執行一次，建立工作表結構）
// ============================================================
function initialize() {
  getOrCreateSheet();

  // 同時建立 QA 工作表
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let qaSheet = ss.getSheetByName('導覽QA');
  if (!qaSheet) {
    qaSheet = ss.insertSheet('導覽QA');
    qaSheet.appendRow(['編號', '問題', '答案', '分類']);
    qaSheet.setFrozenRows(1);

    const headerRange = qaSheet.getRange(1, 1, 1, 4);
    headerRange.setBackground('#1E40AF');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    qaSheet.setColumnWidth(1, 60);
    qaSheet.setColumnWidth(2, 300);
    qaSheet.setColumnWidth(3, 400);
    qaSheet.setColumnWidth(4, 100);
  }

  Logger.log('初始化完成！');
  Logger.log('導覽記錄工作表：已建立');
  Logger.log('導覽QA工作表：已建立');
}
