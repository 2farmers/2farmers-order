// ===============================
// 倆口田訂購系統｜Apps Script 完整版
// 修正版：不使用 appendRow，改成找「訂單編號」最後一列後寫入
// ===============================

const SHEET_PRODUCTS = '商品主檔';
const SHEET_ORDERS = '訂單總表';
const SHEET_DETAILS = '出貨明細';
const SHEET_ACCOUNTING = '記帳總表';
const SHEET_SETTINGS = '網站設定';
const DUPLICATE_MINUTES = 10;
const DEFAULT_WEBSITE_SETTINGS = {
  normalShippingFee: 65,
  normalFreeShippingThreshold: 500,
  lowTempShippingFee: 250,
  lowTempFreeShippingThreshold: 1000,
  shippingDays: '週一、週二、週三、週四'
};

function doGet(e) {
  try {
    return jsonOutput_({
      status: 'success',
      products: getProducts_(),
      settings: getWebsiteSettings_()
    });
  } catch (err) {
    return jsonOutput_({
      status: 'error',
      message: err.message
    });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let writeStarted = false;
  try {
    if (!lock.tryLock(20000)) throw new Error('系統忙碌，請稍後再試');
    const rawPayload = JSON.parse(e && e.postData && e.postData.contents || '{}');
    if (!Array.isArray(rawPayload.items) || !rawPayload.items.length) throw new Error('沒有訂購商品');
    const payload = applyPricingAndTotals_(rawPayload);
    if (!String(payload.receiverName).trim()) throw new Error('請填寫收件人姓名');
    if (!String(payload.receiverPhone).trim()) throw new Error('請填寫電話');
    if (payload.shippingMethod !== '面交' && !String(payload.receiverAddress).trim()) throw new Error('請填寫地址');
    if (isDuplicateOrder_(payload)) {
      return jsonOutput_({ status: 'duplicate', message: '10分鐘內已有相同訂單，請聯絡倆口田確認，勿重複送出' });
    }
    checkStock_(payload.items);
    const orderId = createOrderId_();
    writeStarted = true;
    writeOrder_(orderId, new Date(), payload);
    SpreadsheetApp.flush();
    return jsonOutput_({
      status: 'success', orderId,
      priceType: payload.priceType, partnerName: payload.partnerName || '',
      partnerCode: payload.partnerCode || '', discountRate: payload.discountRate,
      subtotal: payload.subtotal, shipping: payload.shipping, total: payload.total,
      shippingMethod: payload.shippingMethod, items: payload.items
    });
  } catch (err) {
    return jsonOutput_({
      status: writeStarted ? 'uncertain' : 'error', retrySafe: !writeStarted,
      message: writeStarted ? '訂單寫入結果待確認，請聯絡倆口田，勿重複下單' : err.message
    });
  } finally {
    if (lock.hasLock()) {
      try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
    }
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_ORDERS) return;
  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf('狀態') + 1;
  if (!statusCol || e.range.getColumn() > statusCol || e.range.getColumn() + e.range.getNumColumns() <= statusCol) return;
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) throw new Error('系統忙碌，請稍後重新設定訂單狀態');
    for (let row = Math.max(2, e.range.getRow()); row < e.range.getRow() + e.range.getNumRows(); row++) {
      const cell = sheet.getRange(row, statusCol);
      const status = String(cell.getValue() || '').trim();
      if (status === '已確認') {
        try {
          confirmOrder_(sheet, row);
          cell.setNote('');
        } catch (err) {
          cell.setNote('確認未完成：' + err.message + '。排除問題後，再選一次「已確認」可接續處理。');
          cell.setValue('新訂單');
          SpreadsheetApp.getActiveSpreadsheet().toast('訂單第 ' + row + ' 列確認未完成：' + err.message);
          continue;
        }
      }
      if (status === '取消') {
        try {
          cancelOrder_(sheet, row);
          cell.setNote('');
        } catch (err) {
          cell.setNote('取消未完成：' + err.message + '。排除問題後，再選一次「取消」可接續處理。');
          cell.setValue(e.oldValue || '已確認');
          SpreadsheetApp.getActiveSpreadsheet().toast('訂單第 ' + row + ' 列取消未完成：' + err.message);
          continue;
        }
      }
      if (['已確認', '備貨中', '已出貨', '已完成', '取消'].includes(status)) syncOrderStatus_(sheet, row, status);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    e.range.setNote('處理未完成：' + err.message + '。請重新設定狀態。');
    SpreadsheetApp.getActiveSpreadsheet().toast('狀態同步失敗：' + err.message);
  } finally {
    if (lock.hasLock()) {
      try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
    }
  }
}

// Called while holding the script lock. New orders carry a resumable confirmation journal.
function confirmOrder_(orderSheet, rowNumber) {
  ensureHeaders_(orderSheet, ['確認處理']);
  const order = getOrderFromRow_(orderSheet, rowNumber);
  if (!order.orderId) throw new Error('訂單缺少訂單編號');
  const journalCol = getHeaders_(orderSheet).indexOf('確認處理') + 1;
  const journalCell = orderSheet.getRange(rowNumber, journalCol);
  let journal = readConfirmationJournal_(journalCell.getValue());
  if (journal && journal.cancellation && journal.cancellation.phase === 'done') throw new Error('此訂單已取消並回補庫存，不可再次確認；請建立新訂單');
  if (journal && journal.phase === 'done') return;
  const fingerprint = JSON.stringify([order.items, order.subtotal, order.shipping, order.total, order.receiverName, order.receiverPhone, order.receiverAddress]);
  if (journal && journal.fingerprint !== fingerprint) throw new Error('確認處理中訂單內容已被修改，請先核對原始訂單');

  // Do not interleave inventory changes with an interrupted confirmation.
  const journals = orderSheet.getRange(2, journalCol, Math.max(1, orderSheet.getLastRow() - 1), 1).getValues();
  journals.forEach((row, index) => {
    if (index + 2 === rowNumber || !row[0]) return;
    const other = readConfirmationJournal_(row[0]);
    if (other && other.phase !== 'done') throw new Error('請先完成第 ' + (index + 2) + ' 列的確認處理');
  });
  if (!order.items || !order.items.length) throw new Error('訂單缺少商品JSON');
  const productsMap = buildProductsMap_(getProducts_());
  if (!journal) {
    // Existing legacy rows cannot prove whether stock/accounting was completed.
    if (detailsAlreadyExists_(order.orderId) || accountingAlreadyExists_(order.orderId)) {
      throw new Error('此舊訂單已有明細或記帳，但沒有確認紀錄；請先人工核對，避免重複扣庫存');
    }
    checkStock_(order.items);
    journal = { version: 1, phase: 'prepared', fingerprint, inventory: aggregateItems_(order.items).map(item => {
      const product = productsMap[String(item.id)];
      return { id: product.id, before: product.soldQty, after: product.soldQty + item.qty };
    }) };
    journalCell.setValue(JSON.stringify(journal));
    SpreadsheetApp.flush();
  }
  writeDetails_(order.orderId, order.time || new Date(), order, productsMap);
  writeAccounting_(order.orderId, order.time || new Date(), order, productsMap);
  SpreadsheetApp.flush();
  // Absolute targets make a retry safe even if a previous setValue succeeded but its reply failed.
  applyInventoryJournal_(journal.inventory);
  SpreadsheetApp.flush();
  journal.phase = 'done';
  journalCell.setValue(JSON.stringify(journal));
  SpreadsheetApp.flush();
}

function cancelOrder_(orderSheet, rowNumber) {
  ensureHeaders_(orderSheet, ['確認處理']);
  const order = getOrderFromRow_(orderSheet, rowNumber);
  if (!order.orderId) throw new Error('訂單缺少訂單編號');
  const journalCol = getHeaders_(orderSheet).indexOf('確認處理') + 1;
  const journalCell = orderSheet.getRange(rowNumber, journalCol);
  const journal = readConfirmationJournal_(journalCell.getValue());
  if (!journal) return;
  if (journal.phase !== 'done') throw new Error('請先完成此訂單的確認處理');
  if (journal.cancellation && journal.cancellation.phase === 'done') return;

  if (!journal.cancellation) {
    const currentMap = buildProductsMap_(getProducts_());
    journal.cancellation = {
      phase: 'prepared',
      inventory: aggregateItems_(order.items).map(item => {
        const product = currentMap[String(item.id)];
        if (!product) throw new Error('找不到商品：' + item.id);
        const after = product.soldQty - item.qty;
        if (after < 0) throw new Error('商品 ' + product.name + ' 的已售數量不足以回補，請先核對庫存');
        return { id: product.id, before: product.soldQty, after };
      })
    };
    journalCell.setValue(JSON.stringify(journal));
    SpreadsheetApp.flush();
  }
  applyInventoryJournal_(journal.cancellation.inventory);
  SpreadsheetApp.flush();
  journal.cancellation.phase = 'done';
  journalCell.setValue(JSON.stringify(journal));
  SpreadsheetApp.flush();
}

function readConfirmationJournal_(value) {
  if (!value) return null;
  let journal;
  try { journal = JSON.parse(String(value)); } catch (_) { throw new Error('確認處理紀錄格式錯誤，請勿清空紀錄'); }
  if (journal.version !== 1 || !['prepared', 'done'].includes(journal.phase) || !Array.isArray(journal.inventory)) throw new Error('確認處理紀錄格式錯誤');
  return journal;
}

function aggregateItems_(items) {
  const map = Object.create(null);
  items.forEach(item => {
    const id = String(item.id || '').trim();
    const qty = Number(item.qty);
    if (!id || !Number.isFinite(qty) || qty <= 0) throw new Error('商品數量或商品ID錯誤');
    if (!map[id]) map[id] = { ...item, id, qty: 0 };
    map[id].qty += qty;
    if (!Number.isFinite(map[id].qty)) throw new Error('商品數量過大');
  });
  return Object.keys(map).sort().map(id => map[id]);
}

function applyInventoryJournal_(entries) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value).trim());
  const idCol = headers.indexOf('id'), soldCol = headers.indexOf('soldQty');
  if (idCol < 0 || soldCol < 0) throw new Error('商品主檔缺少 id 或 soldQty');
  entries.forEach(entry => {
    if (!Number.isFinite(entry.before) || !Number.isFinite(entry.after) || entry.before < 0 || entry.after < 0) throw new Error('庫存處理紀錄格式錯誤');
    const row = values.findIndex((value, index) => index > 0 && String(value[idCol]).trim() === entry.id);
    if (row < 0) throw new Error('找不到商品：' + entry.id);
    const current = toNumber_(values[row][soldCol], 0);
    if (current !== entry.before && current !== entry.after) throw new Error('商品 ' + entry.id + ' 的已售數量已被另外修改，請先核對庫存');
    if (current === entry.before) sheet.getRange(row + 1, soldCol + 1).setValue(entry.after);
  });
}

function accountingAlreadyExists_(orderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ACCOUNTING);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const values = sheet.getDataRange().getValues();
  const col = values[0].map(String).indexOf('訂單編號');
  return col >= 0 && values.slice(1).some(row => String(row[col]) === String(orderId));
}

// Confirmation output is updated by stable row key, not appended again on retry.
function writeConfirmationRow_(sheet, orderId, key, rowObject) {
  ensureHeaders_(sheet, ['明細鍵']);
  const headers = getHeaders_(sheet);
  const values = sheet.getDataRange().getValues();
  const idCol = headers.indexOf('訂單編號'), keyCol = headers.indexOf('明細鍵');
  const index = values.findIndex((row, i) => i > 0 && String(row[idCol]) === String(orderId) && String(row[keyCol]) === key);
  const targetRow = index < 0 ? findNextRowByColumn_(sheet, idCol + 1) : index + 1;
  const obj = { ...rowObject, '明細鍵': key };
  const rowValues = headers.map((header, col) => Object.prototype.hasOwnProperty.call(obj, header) ? sheetValue_(obj[header]) : (index < 0 ? '' : values[index][col]));
  sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
}

// ===============================
// 後端重新計價
// ===============================

function applyPricingAndTotals_(rawPayload) {
  const products = getProducts_();
  const productsMap = buildProductsMap_(products);
  const settings = getWebsiteSettings_();

  const pricedItems = aggregateItems_(rawPayload.items).map(item => {
    const product = productsMap[String(item.id)];
    if (!product) throw new Error(`找不到商品：${item.name || item.id}`);

    const qty = Number(item.qty || 0);
    if (!Number.isFinite(qty) || qty <= 0) {
      throw new Error(`商品數量錯誤：${product.name}`);
    }

    const originalPrice = Number(product.price);
    if (!Number.isFinite(originalPrice) || originalPrice < 0) throw new Error('商品價格設定錯誤：' + product.name);
    const dealPrice = roundPrice_(originalPrice);
    const amount = dealPrice * qty;
    if (!Number.isFinite(amount)) throw new Error('商品金額過大：' + product.name);
    const shippingType = normalizeShippingType_(product.shippingType);

    return {
      id: product.id,
      name: product.name,
      qty,
      unit: product.unit,
      originalPrice,
      price: dealPrice,
      dealPrice,
      amount,
      total: amount,
      shippingType,
      accountCategory: product.accountCategory || '其他收入',
      priceType: '一般售價',
      partnerName: '',
      partnerCode: '',
      discountRate: 1
    };
  });

  const normalSubtotal = pricedItems
    .filter(item => item.shippingType === 'normal')
    .reduce((sum, item) => sum + item.amount, 0);

  const lowTempSubtotal = pricedItems
    .filter(item => item.shippingType === 'chilled' || item.shippingType === 'frozen')
    .reduce((sum, item) => sum + item.amount, 0);

  const hasNormal = pricedItems.some(item => item.shippingType === 'normal');
  const hasLowTemp = pricedItems.some(item => item.shippingType !== 'normal');

  if (hasNormal && hasLowTemp) {
    throw new Error('常溫商品與低溫商品請分開下單。冷藏與冷凍可以合併低溫宅配。');
  }

  let shippingMethod = rawPayload.shippingMethod || rawPayload.deliveryMethod || '宅配';
  let shipping = 0;

  if (hasLowTemp) {
    shippingMethod = '低溫宅配';
    shipping = lowTempSubtotal >= settings.lowTempFreeShippingThreshold ? 0 : settings.lowTempShippingFee;
  } else if (shippingMethod === '面交') {
    shipping = 0;
  } else {
    shippingMethod = '宅配';
    shipping = normalSubtotal >= settings.normalFreeShippingThreshold ? 0 : settings.normalShippingFee;
  }

  const subtotal = normalSubtotal + lowTempSubtotal;
  const total = subtotal + shipping;

  return {
    ...rawPayload,
    receiverName: rawPayload.receiverName || rawPayload.customerName || '',
    receiverPhone: rawPayload.receiverPhone || rawPayload.phone || '',
    receiverAddress: shippingMethod === '面交' ? '面交自取' : (rawPayload.receiverAddress || rawPayload.customerAddress || ''),
    shippingMethod,
    deliveryMethod: shippingMethod,
    shipping,
    shippingFee: shipping,
    subtotal,
    total,
    grandTotal: total,
    items: pricedItems,
    itemsText: buildItemsText_(pricedItems),
    priceType: '一般售價',
    partnerName: '',
    partnerCode: '',
    discountRate: 1
  };
}

function roundPrice_(value) {
  return Math.round(Number(value || 0));
}

// ===============================
// 商品主檔
// ===============================

function getProducts_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  if (!sheet) throw new Error('找不到商品主檔');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  return rows
    .map(row => rowToProduct_(headers, row))
    .filter(product => product.id && product.name && product.visible !== false)
    .sort((a, b) => {
      const aSold = isSoldOut_(a) ? 1 : 0;
      const bSold = isSoldOut_(b) ? 1 : 0;
      if (aSold !== bSold) return aSold - bSold;
      return Number(a.sortOrder || 999) - Number(b.sortOrder || 999);
    });
}

function rowToProduct_(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index];
  });

  const limitQty = toPositiveNumberOrNull_(obj.limitQty);
  const soldQty = toNumber_(obj.soldQty, 0);
  const remainingQty = limitQty === null ? null : Math.max(0, limitQty - soldQty);

  return {
    id: String(obj.id || '').trim(),
    name: String(obj.name || '').trim(),
    price: toNumber_(obj.price, 0),
    unit: String(obj.unit || '').trim() || '份',
    tag: String(obj.tag || '').trim(),
    desc: String(obj.desc || '').trim(),
    imageUrl: normalizeProductImageUrl_(obj.imageUrl || obj.imageURL || obj.image || ''),
    shippingType: normalizeShippingType_(obj.shippingType || obj.category || obj.tag),
    accountCategory: String(obj.accountCategory || '').trim() || '其他收入',
    soldOut: toBoolean_(obj.soldOut),
    soldOutText: String(obj.soldOutText || '').trim(),
    visible: obj.visible === '' || obj.visible === undefined ? true : toBoolean_(obj.visible),
    sortOrder: toNumber_(obj.sortOrder, 999),
    limitQty,
    soldQty,
    remainingQty
  };
}

function buildProductsMap_(products) {
  const map = Object.create(null);
  products.forEach(product => {
    map[String(product.id)] = product;
  });
  return map;
}

// ===============================
// 寫入訂單總表
// ===============================

function writeOrder_(orderId, now, payload) {
  const sheet = getOrCreateSheet_(SHEET_ORDERS);

  const headers = [
    '訂單編號',
    '時間',
    '收件人',
    '電話',
    '地址',
    '配送方式',
    '商品總金額',
    '運費',
    '總金額',
    '價格類型',
    '合作夥伴名稱',
    '合作代碼',
    '折扣率',
    '商品明細',
    '商品JSON',
    '備註',
    '狀態',
    '付款狀態',
    '匯款末五碼',
    '確認處理'
  ];

  ensureHeaders_(sheet, headers);

  safeWriteObjectByOrderId_(sheet, {
    '訂單編號': orderId,
    '時間': now,
    '收件人': payload.receiverName || payload.customerName || '',
    '電話': payload.receiverPhone || payload.phone || '',
    '地址': payload.receiverAddress || '',
    '配送方式': payload.shippingMethod || payload.deliveryMethod || '',
    '商品總金額': payload.subtotal || 0,
    '運費': payload.shippingFee || payload.shipping || 0,
    '總金額': payload.total || 0,
    '價格類型': payload.priceType || '一般售價',
    '合作夥伴名稱': payload.partnerName || '',
    '合作代碼': payload.partnerCode || '',
    '折扣率': payload.discountRate || 1,
    '商品明細': payload.itemsText || buildItemsText_(payload.items),
    '商品JSON': JSON.stringify(payload.items || []),
    '備註': payload.note || '',
    '狀態': '新訂單',
    '付款狀態': '未付款',
    '匯款末五碼': '',
    '確認處理': ''
  });
}

function getOrderFromRow_(sheet, rowNumber) {
  const headers = getHeaders_(sheet);
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index];
  });

  let items = [];

  try {
    items = JSON.parse(obj['商品JSON'] || '[]');
  } catch (err) {
    items = [];
  }

  return {
    orderId: obj['訂單編號'],
    time: obj['時間'],
    customerName: obj['收件人'],
    customerPhone: obj['電話'],
    receiverName: obj['收件人'],
    receiverPhone: obj['電話'],
    receiverAddress: obj['地址'],
    shippingMethod: obj['配送方式'],
    subtotal: Number(obj['商品總金額'] || 0),
    shippingFee: Number(obj['運費'] || 0),
    shipping: Number(obj['運費'] || 0),
    total: Number(obj['總金額'] || 0),
    priceType: obj['價格類型'] || '一般售價',
    partnerName: obj['合作夥伴名稱'] || '',
    partnerCode: obj['合作代碼'] || '',
    discountRate: Number(obj['折扣率'] || 1),
    itemsText: obj['商品明細'],
    items,
    note: obj['備註'],
    status: obj['狀態']
  };
}

// ===============================
// 出貨明細
// ===============================

function writeDetails_(orderId, now, payload, productsMap) {
  const sheet = getOrCreateSheet_(SHEET_DETAILS);

  const headers = [
    '訂單編號',
    '時間',
    '商品ID',
    '商品名稱',
    '配送類型',
    '記帳分類',
    '原價',
    '成交單價',
    '單價',
    '數量',
    '單位',
    '小計',
    '價格類型',
    '合作夥伴名稱',
    '合作代碼',
    '折扣率',
    '收件人',
    '電話',
    '配送方式',
    '狀態'
  ];

  ensureHeaders_(sheet, headers);

  payload.items.forEach(item => {
    const product = productsMap[String(item.id)] || {};
    const shippingType = normalizeShippingType_(item.shippingType || product.shippingType);
    const originalPrice = Number(item.originalPrice || product.price || item.price || 0);
    const dealPrice = Number(item.price || item.dealPrice || 0);

    writeConfirmationRow_(sheet, orderId, '商品:' + item.id, {
      '訂單編號': orderId,
      '時間': now,
      '商品ID': item.id || '',
      '商品名稱': item.name || '',
      '配送類型': getShippingTypeLabel_(shippingType),
      '記帳分類': product.accountCategory || item.accountCategory || '其他收入',
      '原價': originalPrice,
      '成交單價': dealPrice,
      '單價': dealPrice,
      '數量': Number(item.qty || 0),
      '單位': item.unit || '',
      '小計': Number(item.amount || item.total || 0),
      '價格類型': item.priceType || payload.priceType || '一般售價',
      '合作夥伴名稱': item.partnerName || payload.partnerName || '',
      '合作代碼': item.partnerCode || payload.partnerCode || '',
      '折扣率': item.discountRate || payload.discountRate || 1,
      '收件人': payload.receiverName || payload.customerName || '',
      '電話': payload.receiverPhone || payload.customerPhone || payload.phone || '',
      '配送方式': payload.shippingMethod || payload.deliveryMethod || '',
      '狀態': '已確認'
    });
  });
}

function detailsAlreadyExists_(orderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return false;

  const headers = getHeaders_(sheet);
  const orderIdCol = headers.indexOf('訂單編號');
  if (orderIdCol === -1) return false;

  const values = sheet.getDataRange().getValues();

  return values.slice(1).some(row => String(row[orderIdCol]).trim() === String(orderId).trim());
}

// ===============================
// 記帳總表
// ===============================

function writeAccounting_(orderId, now, payload, productsMap) {
  const sheet = getOrCreateSheet_(SHEET_ACCOUNTING);

  const headers = [
    '日期',
    '收支',
    '大分類',
    '品項',
    '金額',
    '來源',
    '價格類型',
    '合作夥伴名稱',
    '合作代碼',
    '折扣率',
    '訂單編號',
    '備註',
    '狀態'
  ];

  ensureHeaders_(sheet, headers);

  payload.items.forEach(item => {
    const product = productsMap[String(item.id)] || {};

    writeConfirmationRow_(sheet, orderId, '商品:' + item.id, {
      '日期': now,
      '收支': '收入',
      '大分類': product.accountCategory || item.accountCategory || '其他收入',
      '品項': item.name || '',
      '金額': Number(item.amount || item.total || 0),
      '來源': payload.partnerName ? `線上下單｜${payload.partnerName}` : '線上下單',
      '價格類型': item.priceType || payload.priceType || '一般售價',
      '合作夥伴名稱': item.partnerName || payload.partnerName || '',
      '合作代碼': item.partnerCode || payload.partnerCode || '',
      '折扣率': item.discountRate || payload.discountRate || 1,
      '訂單編號': orderId,
      '備註': '商品收入自動寫入',
      '狀態': '已確認'
    });
  });

  const shippingFee = Number(payload.shippingFee || payload.shipping || 0);

  if (shippingFee > 0) {
    writeConfirmationRow_(sheet, orderId, '運費', {
      '日期': now,
      '收支': '收入',
      '大分類': '運費收入',
      '品項': '運費',
      '金額': shippingFee,
      '來源': payload.partnerName ? `線上下單｜${payload.partnerName}` : '線上下單',
      '價格類型': payload.priceType || '一般售價',
      '合作夥伴名稱': payload.partnerName || '',
      '合作代碼': payload.partnerCode || '',
      '折扣率': payload.discountRate || 1,
      '訂單編號': orderId,
      '備註': '運費自動寫入',
      '狀態': '已確認'
    });
  }
}

// ===============================
// 狀態同步
// ===============================

function syncOrderStatus_(orderSheet, rowNumber, newStatus) {
  const order = getOrderFromRow_(orderSheet, rowNumber);
  if (!order.orderId) return;

  syncStatusToSheet_(SHEET_DETAILS, order.orderId, newStatus);
  syncStatusToSheet_(SHEET_ACCOUNTING, order.orderId, newStatus);
}

function syncStatusToSheet_(sheetName, orderId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = getHeaders_(sheet);
  const orderIdCol = headers.indexOf('訂單編號') + 1;
  const statusCol = headers.indexOf('狀態') + 1;

  if (orderIdCol <= 0 || statusCol <= 0) return;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  values.forEach((row, index) => {
    if (String(row[orderIdCol - 1]).trim() === String(orderId).trim()) {
      sheet.getRange(index + 2, statusCol).setValue(newStatus);
    }
  });
}

// ===============================
// 庫存檢查與更新
// ===============================

function checkStock_(items) {
  const products = getProducts_();

  aggregateItems_(items).forEach(item => {
    const product = products.find(p => String(p.id) === String(item.id));
    if (!product) throw new Error(`找不到商品：${item.name || item.id}`);

    if (isSoldOut_(product)) {
      throw new Error(`商品已售完：${product.name}`);
    }

    if (product.remainingQty !== null && Number(item.qty || 0) > product.remainingQty) {
      throw new Error(`${product.name} 庫存不足，目前剩 ${product.remainingQty}${product.unit}`);
    }
  });
}

function updateProductSoldQty_(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODUCTS);
  if (!sheet) throw new Error('找不到商品主檔');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const idCol = headers.indexOf('id');
  const soldQtyCol = headers.indexOf('soldQty');

  if (idCol === -1) throw new Error('商品主檔缺少 id 欄位');
  if (soldQtyCol === -1) throw new Error('商品主檔缺少 soldQty 欄位');

  items.forEach(item => {
    const rowIndex = values.findIndex((row, index) => {
      if (index === 0) return false;
      return String(row[idCol]).trim() === String(item.id).trim();
    });

    if (rowIndex === -1) return;

    const currentSoldQty = toNumber_(values[rowIndex][soldQtyCol], 0);
    const newSoldQty = currentSoldQty + Number(item.qty || 0);

    sheet.getRange(rowIndex + 1, soldQtyCol + 1).setValue(newSoldQty);
  });
}

// ===============================
// 防重複訂單
// ===============================

function isDuplicateOrder_(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return false;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const timeCol = headers.indexOf('時間');
  const phoneCol = headers.indexOf('電話');
  const totalCol = headers.indexOf('總金額');
  const itemsCol = headers.indexOf('商品明細');

  if (timeCol === -1 || phoneCol === -1 || totalCol === -1 || itemsCol === -1) return false;

  const now = new Date();
  const phone = String(payload.receiverPhone || payload.phone || '').trim();
  const total = Number(payload.total || 0);
  const itemsText = payload.itemsText || buildItemsText_(payload.items);

  for (let i = values.length - 1; i >= 1; i--) {
    const orderTime = new Date(values[i][timeCol]);
    const diffMinutes = (now - orderTime) / 1000 / 60;

    if (!Number.isFinite(diffMinutes) || diffMinutes < 0 || diffMinutes > DUPLICATE_MINUTES) continue;

    const samePhone = String(values[i][phoneCol]).trim() === phone;
    const sameTotal = Number(values[i][totalCol]) === total;
    const sameItems = String(values[i][itemsCol]).trim() === String(itemsText).trim();

    if (samePhone && sameTotal && sameItems) return true;
  }

  return false;
}

// ===============================
// 安全寫入：找「訂單編號」最後有資料的位置後寫入
// ===============================

function safeWriteRowByOrderId_(sheet, rowValues) {
  const headers = getHeaders_(sheet);
  let orderIdColIndex = headers.indexOf('訂單編號');

  // 記帳總表也有訂單編號，但不一定在第1欄
  if (orderIdColIndex === -1) {
    orderIdColIndex = 0;
  }

  const targetRow = findNextRowByColumn_(sheet, orderIdColIndex + 1);

  sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
}

function safeWriteObjectByOrderId_(sheet, rowObject) {
  const headers = getHeaders_(sheet);
  let orderIdColIndex = headers.indexOf('訂單編號');

  if (orderIdColIndex === -1) {
    orderIdColIndex = 0;
  }

  const targetRow = findNextRowByColumn_(sheet, orderIdColIndex + 1);
  const rowValues = headers.map(header => {
    if (Object.prototype.hasOwnProperty.call(rowObject, header)) {
      return sheetValue_(rowObject[header]);
    }
    return '';
  });

  sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
}

function findNextRowByColumn_(sheet, colNumber) {
  const lastRow = Math.max(sheet.getLastRow(), 1);

  if (lastRow < 2) return 2;

  const values = sheet.getRange(2, colNumber, lastRow - 1, 1).getValues();

  let lastDataRow = 1;

  values.forEach((row, index) => {
    const value = String(row[0] || '').trim();
    if (value !== '') {
      lastDataRow = index + 2;
    }
  });

  return lastDataRow + 1;
}

// ===============================
// 工具函式
// ===============================


function normalizeProductImageUrl_(value) {
  const url = String(value || '').trim();
  if (!url) return '';

  // Google Drive 圖片嵌入前台時，thumbnail 格式比 uc?export=view 穩定。
  // 支援：
  // 1. https://drive.google.com/file/d/FILE_ID/view?usp=drive_link
  // 2. https://drive.google.com/open?id=FILE_ID
  // 3. https://drive.google.com/uc?export=view&id=FILE_ID
  let match = url.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  if (match && match[1]) {
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1000`;
  }

  match = url.match(/[?&]id=([^&]+)/);
  if (url.includes('drive.google.com') && match && match[1]) {
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1000`;
  }

  return url;
}

function normalizeShippingType_(value) {
  const raw = String(value || '').trim().toLowerCase();

  if (['frozen', 'freeze', '冷凍'].includes(raw)) return 'frozen';
  if (['chilled', 'cold', 'refrigerated', '冷藏', '低溫'].includes(raw)) return 'chilled';

  return 'normal';
}

function getShippingTypeLabel_(type) {
  const normalized = normalizeShippingType_(type);
  if (normalized === 'frozen') return '冷凍';
  if (normalized === 'chilled') return '冷藏';
  return '常溫';
}

function isSoldOut_(product) {
  return product.soldOut === true || product.remainingQty === 0;
}

function buildItemsText_(items) {
  return (items || [])
    .map(item => {
      const qty = Number(item.qty || 0);
      const unit = item.unit || '';
      const amount = Number(item.amount || item.total || 0);
      const price = Number(item.price || item.dealPrice || 0);
      return `${item.name} x ${qty}${unit}（單價${price}）= ${amount}`;
    })
    .join('\n');
}

function createOrderId_() {
  const now = new Date();
  const prefix = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMdd-HHmmss');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  const existing = sheet && sheet.getLastRow() > 1 ? new Set(sheet.getDataRange().getValues().slice(1).map(row => String(row[0]))) : new Set();
  for (let attempt = 0; attempt < 5; attempt++) {
    const suffix = String(Utilities.getUuid()).replace(/[^a-z0-9]/gi, '').slice(0, 4).toUpperCase().padEnd(4, '0');
    const id = prefix + '-' + suffix;
    if (!existing.has(id)) return id;
  }
  throw new Error('無法產生唯一訂單編號，請重新送出');
}

function getWebsiteSettings_() {
  const settings = { ...DEFAULT_WEBSITE_SETTINGS };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return settings;
  const values = sheet.getDataRange().getValues();
  values.slice(1).forEach(row => {
    const key = String(row[0] || '').trim();
    if (!Object.prototype.hasOwnProperty.call(settings, key)) return;
    if (key === 'shippingDays') {
      const value = String(row[1] || '').trim();
      if (value) settings[key] = value;
      return;
    }
    const value = Number(row[1]);
    if (Number.isFinite(value) && value >= 0) settings[key] = value;
  });
  return settings;
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const currentHeaders = getHeaders_(sheet);

  if (currentHeaders.join('') === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  headers.forEach(header => {
    if (!currentHeaders.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });
}

function toNumber_(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function toPositiveNumberOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;

  const num = Number(value);
  if (!Number.isFinite(num) || num <= 0) return null;

  return num;
}

function toBoolean_(value) {
  if (value === true) return true;
  if (value === false) return false;

  const raw = String(value || '').trim().toLowerCase();

  return ['true', 'yes', 'y', '1', '是', '售完'].includes(raw);
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
// Keep user text literal in spreadsheet cells (including leading-zero phone numbers).
function sheetValue_(value) {
  return typeof value === 'string' && /^[=+@\-]|^0[0-9]/.test(value) ? "'" + value : value;
}
