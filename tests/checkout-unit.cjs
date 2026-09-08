// Run with Node.js: node tests/checkout-unit.cjs
// Minimal DOM adapters exercise application logic without a browser or network.
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const source = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
class Element {
  constructor() { this.value = ''; this.style = {}; this.dataset = {}; this.children = []; this.listeners = {}; this.textContent = ''; }
  addEventListener(type, callback) { this.listeners[type] = callback; }
  append(...children) { this.children.push(...children); }
  replaceChildren() { this.children = []; }
  closest() { return this; }
  scrollIntoView() {}
}
function setup() {
  const elements = new Map();
  const get = id => { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); };
  const radios = ['宅配', '面交'].map(value => Object.assign(new Element(), { value, dataset: { fee: value === '宅配' ? '65' : '0' } }));
  let selected = radios[0];
  radios.forEach(radio => Object.defineProperty(radio, 'checked', { get: () => selected === radio, set: value => { if (value) selected = radio; } }));
  const document = {
    getElementById: get,
    querySelector: selector => selector.includes('name="shipping"') ? selector.includes(':checked') ? selected : radios.find(r => selector.includes(`value="${r.value}"`)) : get(selector),
    querySelectorAll: selector => selector === 'input[name="shipping"]' ? radios : ['customerName', 'customerPhone', 'customerAddress', 'customerNote'].map(get),
    createElement: () => new Element()
  };
  const requests = [];
  let mode = 'success';
  let release;
  const context = vm.createContext({ document, console, navigator: { userAgent: 'unit-test' }, window: { addEventListener() {} }, fetch: async (url, options = {}) => {
    if (options.method !== 'POST') return { json: async () => ({ status: 'success', partner: { isPartner: true, partnerCode: 'TEST', partnerName: '測試', discountRate: 0.8 } }) };
    requests.push(JSON.parse(options.body));
    if (mode === 'network') throw Error('simulated disconnect');
    if (mode === 'pending') await new Promise(resolve => { release = resolve; });
    return { text: async () => JSON.stringify(mode === 'reject' ? { status: 'error', message: '庫存不足' } : mode === 'duplicate' ? { status: 'duplicate' } : mode === 'no-id' ? { status: 'success' } : { status: 'success', orderId: 'TEST-' + requests.length }) };
  } });
  vm.runInContext(source, context);
  const run = code => vm.runInContext(code, context);
  run(`products = [{id:'veg',name:'蔬菜',unit:'份',price:100,shippingType:'normal',remainingQty:20},{id:'dumpling',name:'水餃',unit:'包',price:250,shippingType:'frozen',remainingQty:20},{id:'pickle',name:'酸菜',unit:'罐',price:150,shippingType:'chilled',remainingQty:20}]; sortedProducts = products; updateTotals();`);
  const qty = (id, count) => run(`setQty(${JSON.stringify(id)}, ${count})`);
  const customer = () => { get('customerName').value = '測試姓名'; get('customerPhone').value = '0912345678'; get('customerAddress').value = '測試地址'; };
  return { get, run, qty, customer, requests, radios, mode: value => { mode = value; }, release: () => release() };
}
(async () => {
  const t = setup();
  assert.equal(t.get('submitBtn').disabled, true);
  t.qty('veg', 4); assert.equal(t.get('grandTotal').textContent, '465 元');
  t.qty('veg', 5); assert.equal(t.get('shippingFee').textContent, '0 元');
  t.qty('dumpling', 3); assert.equal(t.get('stickyGrandTotal').textContent, '1500 元');
  assert.equal(t.get('cartNotice').hidden, false);
  t.run("checkoutGroup = 'lowTemp'; updateTotals();");
  assert.equal(t.get('grandTotal').textContent, '1000 元');
  assert.equal(t.radios[1].disabled, true);
  t.qty('dumpling', 4); assert.equal(t.get('shippingFee').textContent, '0 元');
  t.qty('pickle', 1); t.customer();
  t.mode('pending'); const saving = t.run('prepareSavedOrder()');
  await t.run('prepareSavedOrder()');
  assert.equal(t.requests.length, 1);
  assert.equal(t.get('cart').inert, true);
  t.qty('veg', 9); assert.equal(t.run('state.quantities.veg'), 5);
  t.mode('success'); t.release(); await saving;
  assert.equal(t.requests[0].shippingMethod, '低溫宅配');
  assert.equal(t.requests[0].total, 1150);
  assert.deepEqual(t.requests[0].items.map(i => i.id), ['dumpling', 'pickle']);
  assert.equal(t.run('state.quantities.veg'), 5);
  assert.equal(t.run('state.quantities.dumpling'), 0);
  assert.equal(t.get('customerName').value, '測試姓名');
  const receipt = t.run('savedOrders[0].text');
  assert.match(receipt, /TEST-1/);
  assert.match(t.run('getLineUrl(savedOrders[0].text)'), /TEST-1/);
  t.qty('veg', 4); t.get('customerNote').value = '第二筆';
  t.run('clearPreparedOrder()'); assert.equal(t.run('savedOrders[0].text'), receipt);
  t.radios[1].checked = true; t.radios[1].listeners.change({ target: t.radios[1] });
  t.qty('dumpling', 1);
  t.run("checkoutGroup='lowTemp'; updateTotals(); checkoutGroup='normal'; updateTotals();");
  assert.equal(t.radios[1].checked, true);
  t.qty('dumpling', 0); t.get('customerAddress').value = '';
  await t.run('prepareSavedOrder()');
  assert.equal(t.requests[1].total, 400);
  assert.equal(t.requests[1].shippingMethod, '面交');
  assert.equal(t.requests[1].receiverAddress, '面交自取');
  assert.equal(t.run('savedOrders.length'), 2);
  assert.equal(t.get('submitBtn').disabled, true);
  console.log('PASS shipping thresholds, mixed totals, split payloads, retained cart/form, pickup, receipts/LINE IDs, submission lock');
  for (const mode of ['network', 'duplicate', 'reject', 'no-id']) {
    const f = setup(); f.mode(mode); f.qty('veg', 1); f.customer();
    await f.run('prepareSavedOrder()');
    assert.equal(f.requests.length, 1);
    if (mode === 'no-id') assert.equal(f.run('savedOrders[0].orderId'), '');
    else {
      assert.equal(f.run('state.quantities.veg'), 1);
      assert.equal(f.get('submitBtn').disabled, mode !== 'reject');
      if (mode !== 'reject') {
        assert.match(f.get('orderOutput').textContent, /送出結果待確認/);
        assert.equal(f.get('lineBtn').disabled, false);
        await f.run('prepareSavedOrder()'); assert.equal(f.requests.length, 1);
      }
    }
  }
  console.log('PASS network uncertainty, duplicate response, rejection recovery, absent server ID');
  const p = setup(); p.qty('veg', 1); await p.run('prepareSavedOrder()');
  assert.equal(p.requests.length, 0); assert.equal(p.get('errorMsg').textContent, '請填寫收件人姓名。');
  p.get('partnerCodeInput').value = 'TEST'; await p.run('applyPartnerCode()');
  assert.equal(p.get('grandTotal').textContent, '145 元');
  p.qty('veg', 100); assert.equal(p.run('state.quantities.veg'), 20);
  p.qty('veg', 0); assert.equal(p.get('submitBtn').disabled, true);
  console.log('PASS required fields, partner price, stock cap, removal; no network used.');
})().catch(error => { console.error(error); process.exitCode = 1; });
