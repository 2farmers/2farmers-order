/* Apps Script RPC-compatible callbacks over HTTPS; no iframe, JSONP or credential storage. */
(function () {
  'use strict';
  const methods = ['ping', 'lookupMember', 'verifyExistingMember', 'vegetablePlanOptions',
    'fruitApplicationOptions', 'submitFruitApplication', 'submitVegetableApplication',
    'reportFruitPayment', 'reportVegetablePayment', 'submitMemberMessage',
    'submitPollAnswer', 'setFarmVisitRegistration'];
  const paymentTokens = new Map();
  const writes = new Set(['submitFruitApplication', 'submitVegetableApplication',
    'reportFruitPayment', 'reportVegetablePayment', 'submitMemberMessage',
    'submitPollAnswer', 'setFarmVisitRegistration']);
  function showError(error) {
    const node = document.getElementById('apiStatus');
    if (node) { node.className = 'notice error'; node.textContent = error.message; }
  }
  async function request(method, args) {
    const endpoint = window.COFARM_CONFIG && window.COFARM_CONFIG.endpoint;
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(endpoint || ''))
      throw new Error('共耕資料連線尚未設定，請聯絡倆口田。');
    const tokenKind = method === 'reportFruitPayment' ? 'fruit' :
      method === 'reportVegetablePayment' ? 'vegetable' : '';
    const token = tokenKind ? paymentTokens.get(tokenKind + ':' + args[0]) : '';
    if (tokenKind && !token) throw new Error('付款回報驗證已失效，請聯絡倆口田並提供申請編號，請勿重複申請。');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 45000);
    let body;
    try {
      const response = await fetch(endpoint, {
        method: 'POST', mode: 'cors', credentials: 'omit', redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ version: 1, method, args, paymentToken: token || '' }),
        signal: controller.signal
      });
      if (!response.ok) throw new Error('HTTP ' + response.status);
      body = await response.json();
    } catch (_) {
      throw new Error(writes.has(method)
        ? '未能取得送出結果，資料可能已送達。請先聯絡倆口田確認，避免重複送出。'
        : '暫時無法連接共耕資料。請稍後再試；若持續發生，請聯絡倆口田。');
    } finally { clearTimeout(timer); }
    if (!body || body.version !== 1 || typeof body.ok !== 'boolean')
      throw new Error('共耕後台版本尚未更新，請聯絡倆口田。');
    if (!body.ok) throw new Error(body.error || '操作未完成，請稍後再試。');
    const kind = method === 'submitFruitApplication' ? 'fruit' :
      method === 'submitVegetableApplication' ? 'vegetable' : '';
    if (kind && body.data && body.data.id && body.paymentToken)
      paymentTokens.set(kind + ':' + body.data.id, body.paymentToken);
    return body.data;
  }
  function runner(success, failure) {
    const api = {
      withSuccessHandler(callback) { return runner(callback, failure); },
      withFailureHandler(callback) { return runner(success, callback); }
    };
    methods.forEach(method => {
      api[method] = function (...args) {
        request(method, args).then(data => {
          if (success) success(data);
        }, error => { if (failure) failure(error); else showError(error); });
      };
    });
    return Object.freeze(api);
  }
  window.cofarmRun = runner(null, null);
})();
