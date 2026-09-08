// Run: NODE_PATH="$CODEX_PRIMARY_RUNTIME_NODE_MODULES" node tests/checkout.cjs
// All network requests are intercepted; no production orders are sent.
const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
const products = [
  { id: 'veg', name: '測試蔬菜', unit: '份', price: 100, shippingType: 'normal', remainingQty: 20 },
  { id: 'dumpling', name: '測試水餃', unit: '包', price: 250, shippingType: 'frozen', remainingQty: 20 },
  { id: 'pickle', name: '測試酸菜', unit: '罐', price: 150, shippingType: 'chilled', remainingQty: 20 }
];
(async () => {
  const browser = await chromium.launch({ headless: true });
  const errors = [];
  async function setup(mode = 'success', width = 390) {
    const page = await browser.newPage({ viewport: { width, height: 844 } });
    page.on('pageerror', error => errors.push(error.message));
    const posts = [];
    await page.route('**/*', async route => {
      const request = route.request();
      if (request.url() === 'https://checkout.test/') return route.fulfill({ contentType: 'text/html', body: html });
      if (request.url().startsWith('https://script.google.com/')) {
        if (request.method() !== 'POST') {
          const body = request.url().includes('validatePartner')
            ? { status: 'success', partner: { isPartner: true, partnerCode: 'TEST', partnerName: '測試夥伴', discountRate: 0.8 } }
            : { status: 'success', products };
          return route.fulfill({ json: body });
        }
        posts.push(JSON.parse(request.postData()));
        if (mode === 'network') return route.abort('failed');
        const body = mode === 'reject' ? { status: 'error', retrySafe: true, message: '庫存不足' }
          : mode === 'duplicate' ? { status: 'duplicate' }
          : mode === 'no-id' ? { status: 'success' }
          : { status: 'success', orderId: 'TEST-' + posts.length };
        return route.fulfill({ json: body });
      }
      return route.abort();
    });
    await page.goto('https://checkout.test/');
    await page.waitForSelector('#qty-veg');
    return { page, posts };
  }
  async function qty(page, id, count) { await page.locator('#qty-' + id).fill(String(count)); }
  async function customer(page) {
    await page.locator('#customerName').fill('測試收件人');
    await page.locator('#customerPhone').fill('0912345678');
    await page.locator('#customerAddress').fill('測試地址（不出貨）');
  }
  async function text(page, id, expected) { assert.equal(await page.locator('#' + id).textContent(), expected); }

  const { page, posts } = await setup();
  assert.equal(await page.locator('#submitBtn').isDisabled(), true);
  await qty(page, 'veg', 4);
  await text(page, 'grandTotal', '465 元');
  await qty(page, 'veg', 5);
  await text(page, 'shippingFee', '0 元');
  await qty(page, 'dumpling', 3);
  await text(page, 'stickyGrandTotal', '1500 元');
  await text(page, 'grandTotal', '500 元');
  assert.equal(await page.locator('#cartNotice').isVisible(), true);
  await page.locator('[data-checkout-group="lowTemp"]').click();
  await text(page, 'grandTotal', '1000 元');
  assert.equal(await page.locator('input[value="面交"]').isDisabled(), true);
  await qty(page, 'dumpling', 4);
  await text(page, 'shippingFee', '0 元');
  await qty(page, 'pickle', 1);
  await customer(page);
  await page.locator('#submitBtn').click();
  await page.waitForSelector('.receipt');
  assert.equal(posts.length, 1);
  assert.deepEqual(posts[0].items.map(item => item.id), ['dumpling', 'pickle']);
  assert.equal(posts[0].shippingMethod, '低溫宅配');
  assert.equal(posts[0].total, 1150);
  assert.equal(await page.locator('#qty-veg').inputValue(), '5');
  assert.equal(await page.locator('#qty-dumpling').inputValue(), '0');
  assert.equal(await page.locator('#customerName').inputValue(), '測試收件人');
  const receiptText = await page.locator('.receipt pre').textContent();
  assert.match(receiptText, /TEST-1/);
  assert.match(await page.locator('.receipt a').getAttribute('href'), /TEST-1/);
  await qty(page, 'veg', 4);
  await page.locator('#customerNote').fill('第二筆備註');
  assert.equal(await page.locator('.receipt pre').textContent(), receiptText);
  await page.locator('input[value="面交"]').check();
  await text(page, 'grandTotal', '400 元');
  await page.locator('#customerAddress').fill('');
  await page.locator('#submitBtn').click();
  await page.waitForFunction(() => document.querySelectorAll('.receipt').length === 2);
  assert.equal(posts.length, 2);
  assert.equal(posts[1].shippingMethod, '面交');
  assert.equal(posts[1].total, 400);
  assert.equal(posts[1].receiverAddress, '面交自取');
  await text(page, 'stickyGrandTotal', '0 元');
  assert.equal(await page.locator('#submitBtn').isDisabled(), true);
  console.log('PASS: mixed cart, both shipping thresholds, retained second group, pickup, immutable receipts, LINE IDs');

  for (const mode of ['network', 'duplicate', 'reject', 'no-id']) {
    const { page: p, posts: requests } = await setup(mode);
    await qty(p, 'veg', 1);
    await customer(p);
    await p.locator('#submitBtn').click();
    await p.waitForFunction(() => !isSubmittingOrder);
    assert.equal(requests.length, 1);
    if (mode === 'no-id') {
      assert.match(await p.locator('.receipt').textContent(), /系統未提供訂單編號/);
    } else {
      assert.equal(await p.locator('#qty-veg').inputValue(), '1');
      assert.equal(await p.locator('#submitBtn').isDisabled(), mode !== 'reject');
      if (mode !== 'reject') {
        assert.match(await p.locator('#orderOutput').textContent(), /送出結果待確認/);
        assert.equal(await p.locator('#lineBtn').isEnabled(), true);
      }
    }
    await p.close();
  }
  console.log('PASS: uncertain network result, duplicate protection, explicit rejection, absent order number');

  const { page: p } = await setup('success', 360);
  await qty(p, 'veg', 1);
  await p.locator('#submitBtn').click();
  await text(p, 'errorMsg', '請填寫收件人姓名。');
  await p.locator('#partnerCodeInput').fill('TEST');
  await p.locator('#applyPartnerCodeBtn').click();
  await p.waitForFunction(() => !isValidatingPartner && state.partner);
  await text(p, 'grandTotal', '145 元');
  await p.locator('[data-cart-id="veg"][data-diff="1"]').click();
  await text(p, 'grandTotal', '225 元');
  await p.locator('[data-remove-id="veg"]').click();
  assert.equal(await p.locator('#submitBtn').isDisabled(), true);
  for (const width of [360, 390, 760, 1280]) {
    await p.setViewportSize({ width, height: 900 });
    await qty(p, 'veg', 2);
    await qty(p, 'dumpling', 2);
    assert.equal(await p.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true, 'overflow at ' + width);
  }
  console.log('PASS: validation, partner pricing, cart controls, 360/390/760/1280 layouts');
  assert.deepEqual(errors, []);
  await browser.close();
  console.log('All checkout tests passed; zero production requests.');
})().catch(error => { console.error(error); process.exit(1); });
