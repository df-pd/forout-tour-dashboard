/**
 * 前端認證模組
 * 登入驗證：透過 Google Apps Script（帳號存在 Google Sheet「帳號」工作表）
 * 預約/諮詢 API：透過 Google Apps Script
 */
const Auth = {
  // Google Apps Script Web App URL
  GAS_URL: 'https://script.google.com/macros/s/AKfycbxJnBjVuPweU9SsqwxNF8YuPAyQJNtkp3M0G_6udXijr9vUkUAR2adYQkr3mR54EalwSw/exec',

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
