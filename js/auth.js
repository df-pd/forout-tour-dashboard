/**
 * 前端認證模組
 * 登入驗證：透過 Google Apps Script（帳號存在 Google Sheet「帳號」工作表）
 * 預約/諮詢 API：透過 Google Apps Script
 */
const Auth = {
  // Google Apps Script Web App URL
  GAS_URL: 'https://script.google.com/macros/s/AKfycbwXgJQeyY7fUOU4gpWSGK31LVB7HbuYnN86enmeAp3dI6RAK8nPb4Cmkp0uoxEIXtpKmQ/exec',

  /**
   * 登入（透過 GAS 比對 Google Sheet 帳號）
   */
  async login(username, password) {
    try {
      const resp = await fetch(this.GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'login', username, password }),
      });
      const result = await resp.json();

      if (result.success) {
        const today = new Date();
        today.setHours(23, 59, 59, 999);

        localStorage.setItem('auth_token', result.token);
        localStorage.setItem('auth_user', JSON.stringify(result.user));
        localStorage.setItem('auth_expiry', today.getTime().toString());
        // 儲存可見功能權限（由 GAS 根據帳號工作表 E 欄回傳）
        if (result.permissions) {
          localStorage.setItem('auth_permissions', JSON.stringify(result.permissions));
        }
      }

      return result;
    } catch (err) {
      return { success: false, error: '連線失敗，請檢查網路' };
    }
  },

  /**
   * 登出
   */
  logout() {
    localStorage.removeItem('auth_token');
    localStorage.removeItem('auth_user');
    localStorage.removeItem('auth_expiry');
    localStorage.removeItem('auth_permissions');
    window.location.href = 'index.html';
  },

  /**
   * 是否已登入
   */
  isLoggedIn() {
    const token = localStorage.getItem('auth_token');
    const expiry = localStorage.getItem('auth_expiry');
    if (!token || !expiry) return false;
    return Date.now() < parseInt(expiry);
  },

  /**
   * 取得目前使用者資訊
   */
  getUser() {
    try {
      return JSON.parse(localStorage.getItem('auth_user'));
    } catch {
      return null;
    }
  },

  /**
   * 取得可見功能權限
   * 回傳陣列如 ['booking', 'list', 'stats', 'qa']
   * 若無設定或 admin 角色，回傳所有功能
   */
  getPermissions() {
    try {
      const perms = JSON.parse(localStorage.getItem('auth_permissions'));
      if (Array.isArray(perms) && perms.length > 0) return perms;
    } catch { /* 忽略 */ }
    // 預設全部可見
    return ['booking', 'list', 'stats', 'qa'];
  },

  /**
   * 檢查是否有特定功能的權限
   */
  hasPermission(tabName) {
    return this.getPermissions().includes(tabName);
  },

  /**
   * 取得 token
   */
  getToken() {
    return localStorage.getItem('auth_token');
  },

  /**
   * 認證守衛：未登入跳轉首頁
   */
  guard() {
    if (!this.isLoggedIn()) {
      window.location.href = 'index.html';
      return false;
    }
    return true;
  },

  /**
   * 送出需認證的 API 請求到 Google Apps Script
   */
  /**
   * 修改帳號密碼
   */
  async updateCredentials(currentPassword, newUsername, newPassword) {
    const result = await this.apiCall('updateCredentials', { currentPassword, newUsername, newPassword });
    return result;
  },

  /**
   * 記住帳密：儲存
   */
  saveCredentials(username, password) {
    localStorage.setItem('saved_username', username);
    localStorage.setItem('saved_password', btoa(unescape(encodeURIComponent(password))));
  },

  /**
   * 記住帳密：清除
   */
  clearSavedCredentials() {
    localStorage.removeItem('saved_username');
    localStorage.removeItem('saved_password');
  },

  /**
   * 記住帳密：讀取
   */
  getSavedCredentials() {
    const username = localStorage.getItem('saved_username');
    const encoded = localStorage.getItem('saved_password');
    if (!username || !encoded) return null;
    try {
      return { username, password: decodeURIComponent(escape(atob(encoded))) };
    } catch {
      return null;
    }
  },

  async apiCall(action, data = {}) {
    try {
      const resp = await fetch(this.GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, token: this.getToken(), ...data }),
      });
      return await resp.json();
    } catch (err) {
      return { success: false, error: '連線失敗，請檢查網路' };
    }
  }
};
