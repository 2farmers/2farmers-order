// node tests/backend.cjs — in-memory Sheets/Lock mocks; never connects to Google.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../backend/Code.gs'), 'utf8');
function setup() {
  let held = false, busy = false, nextId = 0;
  let failure = null;
  const events = [];
  const sheets = new Map();
  const fail = (sheet, row, col, after) => {
    if (failure && failure.sheet === sheet && (!failure.col || failure.col === col) && row > 1 && failure.after === after) {
      failure = null;
      throw new Error('simulated write failure');
    }
  };
  class Range {
    constructor(sheet,r,c,nr=1,nc=1) { Object.assign(this,{sheet,r,c,nr,nc}); }
    getValues() { return Array.from({length:this.nr},(_,r)=>Array.from({length:this.nc},(_,c)=>this.sheet.rows[this.r+r-1]?.[this.c+c-1] ?? '')); }
    getValue() { return this.getValues()[0][0]; }
    setValues(values) {
      assert.equal(held,true,'writes must hold script lock');
      fail(this.sheet.name,this.r,this.c,false);
      values.forEach((row,r)=>row.forEach((value,c)=>{
        this.sheet.rows[this.r+r-1] ||= [];
        this.sheet.rows[this.r+r-1][this.c+c-1] = typeof value === 'string' && value.startsWith("'") ? value.slice(1) : value;
      }));
      events.push(['write',this.sheet.name,this.r,this.c]);
      fail(this.sheet.name,this.r,this.c,true);
      return this;
    }
    setValue(value) { return this.setValues([[value]]); }
    copyTo(destination) { destination.setValues(this.getValues()); return this; }
    setNote(value) { this.note=value; return this; }
    getSheet() { return this.sheet; }
    getRow() { return this.r; }
    getColumn() { return this.c; }
    getNumRows() { return this.nr; }
    getNumColumns() { return this.nc; }
  }
  class Sheet {
    constructor(name,rows=[]) {this.name=name;this.rows=rows;}
    getName(){return this.name;}
    getLastRow(){return this.rows.length;}
    getLastColumn(){return Math.max(0,...this.rows.map(r=>r.length));}
    getRange(...args){return new Range(this,...args);}
    getDataRange(){return this.getRange(1,1,Math.max(1,this.getLastRow()),Math.max(1,this.getLastColumn()));}
    deleteRow(row){assert.equal(held,true,'deletes must hold script lock');this.rows.splice(row-1,1);events.push(['delete',this.name,row]);}
  }
  const book = {getSheetByName:name=>sheets.get(name),insertSheet:name=>{const s=new Sheet(name);sheets.set(name,s);return s;},toast:message=>events.push(['toast',message])};
  sheets.set('商品主檔',new Sheet('商品主檔',[
    ['id','name','price','unit','shippingType','limitQty','soldQty','visible','accountCategory'],
    ['veg','蔬菜',100,'份','normal',10,0,true,'蔬果收入'],
    ['dumpling','水餃',250,'包','frozen',10,0,true,'料理收入']
  ]));
  const context=vm.createContext({console,Date,SpreadsheetApp:{getActiveSpreadsheet:()=>book,flush:()=>{assert.equal(held,true);events.push(['flush']);}},
    LockService:{getScriptLock:()=>({tryLock:()=>{if(busy||held)return false;held=true;events.push(['lock']);return true;},hasLock:()=>held,releaseLock:()=>{held=false;events.push(['release']);}})},
    Utilities:{formatDate:(_date,_tz,pattern)=> pattern === 'yyMMdd-HHmmss' ? '260908-123456' : '20260908',getUuid:()=>String(++nextId)},
    ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({setMimeType:()=>JSON.parse(text)})}
  });
  vm.runInContext(source,context);
  const post=(items=[{id:'veg',qty:2}],more={})=>context.doPost({postData:{contents:JSON.stringify({items,receiverName:'測試',receiverPhone:'0900000000',receiverAddress:'測試地址',...more})}});
  const confirm=row=>{
    const orders=sheets.get('訂單總表'), col=orders.rows[0].indexOf('狀態')+1;
    orders.rows[row-1][col-1]='已確認';
    context.onEdit({range:orders.getRange(row,col),value:'已確認'});
  };
  const status=(row,value,oldValue,sheetName='訂單總表')=>{
    const orders=sheets.get(sheetName), col=orders.rows[0].indexOf('狀態')+1;
    orders.rows[row-1][col-1]=value;
    context.onEdit({range:orders.getRange(row,col),value,oldValue});
  };
  const objects=name=>{const s=sheets.get(name);return s ? s.rows.slice(1).map(row=>Object.fromEntries(s.rows[0].map((h,i)=>[h,row[i]]))):[];};
  return {context,post,confirm,status,objects,sheets,events,setBusy:v=>{busy=v;},fail:(sheet,after=false,col=null)=>{failure={sheet,after,col};}};
}
{
  const t=setup();
  const order=t.post([{id:'dumpling',qty:1,shippingType:'normal',price:1}],{partnerCode:'TEST',discountRate:0.1});
  assert.equal(order.status,'success'); assert.equal(order.shipping,250);assert.equal(order.total,500);assert.equal(order.items[0].price,250);assert.equal(order.items[0].shippingType,'frozen');
  assert.equal(order.priceType,'一般售價');assert.equal(order.partnerCode,'');assert.equal(order.discountRate,1);
  assert.match(order.orderId,/^260908-123456-[A-Z0-9]{4}$/);
  assert.equal(t.objects('訂單總表').length,1);
  assert.equal(t.objects('商品主檔')[1].soldQty,0,'new order does not reserve stock');
  assert.equal(t.post([{id:'dumpling',qty:1}]).status,'duplicate');
  assert.equal(t.objects('訂單總表').length,1);
  assert.equal(t.post([{id:'veg',qty:6},{id:'veg',qty:6}]).status,'error');
  const grouped=t.post([{id:'veg',qty:2},{id:'veg',qty:2}]);
  assert.equal(grouped.items.length,1);assert.equal(grouped.items[0].qty,4);assert.equal(grouped.total,465);
  assert.equal(t.post([{id:'veg',qty:1},{id:'dumpling',qty:1}]).retrySafe,true);
  assert.equal(t.post([{id:'veg',qty:1}],{receiverName:''}).status,'error');
  assert.equal(t.events.at(-1)[0],'release');
  t.setBusy(true); const count=t.objects('訂單總表').length;
  assert.equal(t.post().retrySafe,true);assert.equal(t.objects('訂單總表').length,count);
  console.log('PASS authoritative prices/temperature, aggregated stock, duplicate protection, validation, lock contention, no reservation on receipt');
}
{
  const t=setup();t.fail('訂單總表',true);
  const result=t.post();assert.equal(result.status,'uncertain');assert.equal(result.retrySafe,false);
  assert.equal(t.objects('訂單總表').length,1);
  assert.equal(t.post().status,'duplicate');
  console.log('PASS lost acknowledgement after order write does not invite duplicate retry');
}
for(const [sheet,after,col] of [['出貨明細',false,null],['出貨明細',true,null],['記帳總表',false,null],['記帳總表',true,null],['商品主檔',true,7]]) {
  const t=setup(); t.post([{id:'veg',qty:2}]);t.fail(sheet,after,col);t.confirm(2);
  assert.equal(t.objects('訂單總表')[0]['狀態'],'新訂單');
  assert.equal(JSON.parse(t.objects('訂單總表')[0]['確認處理']).phase,'prepared');
  t.confirm(2);t.confirm(2);
  assert.equal(JSON.parse(t.objects('訂單總表')[0]['確認處理']).phase,'done');
  assert.equal(t.objects('出貨明細').length,1);
  assert.equal(t.objects('記帳總表').length,2);
  assert.equal(t.objects('商品主檔')[0].soldQty,2);
  assert.equal(t.objects('訂單總表')[0]['狀態'],'已確認');
  assert.equal(t.objects('記帳總表').reduce((n,r)=>n+r['金額'],0),265);
}
console.log('PASS resume after details/accounting/inventory failures, including writes that succeeded before error; no double deductions');
{
  const t=setup();t.post();t.fail('商品主檔',true,7);t.confirm(2);
  t.post([{id:'veg',qty:1}]);t.confirm(3);
  assert.equal(t.objects('訂單總表')[1]['狀態'],'新訂單');
  assert.equal(t.objects('商品主檔')[0].soldQty,2);
  t.confirm(2);t.confirm(3);
  assert.equal(t.objects('商品主檔')[0].soldQty,3);
  console.log('PASS unresolved confirmation blocks other inventory confirmations until resumed');
}
{
  const t=setup();t.post();t.fail('記帳總表');t.confirm(2);
  t.sheets.get('商品主檔').rows[1][6]=7;
  t.confirm(2);assert.equal(t.objects('訂單總表')[0]['狀態'],'新訂單');assert.equal(t.objects('商品主檔')[0].soldQty,7);
  console.log('PASS manual inventory conflict is flagged without overwriting stock');
}
{
  const t=setup(); t.post();t.confirm(2);
  const s=t.sheets.get('訂單總表');s.rows[1][s.rows[0].indexOf('確認處理')]='';
  t.confirm(2);assert.equal(t.objects('商品主檔')[0].soldQty,2);
  assert.equal(t.objects('訂單總表')[0]['狀態'],'新訂單');
  console.log('PASS legacy records with unproven completion require manual review');
}
console.log('All backend tests passed. No live Google calls.');

{
  const t=setup();
  const settings=t.context.SpreadsheetApp.getActiveSpreadsheet().insertSheet('網站設定');
  settings.rows=[['設定鍵','設定值'],['normalShippingFee',80],['normalFreeShippingThreshold',300],['lowTempShippingFee',300],['lowTempFreeShippingThreshold',700],['shippingDays','週二、週四']];
  const result=t.post([{id:'veg',qty:2}]);
  assert.equal(result.shipping,80);assert.equal(result.total,280);
  assert.equal(t.context.doGet({}).settings.shippingDays,'週二、週四');
  console.log('PASS website settings control authoritative shipping and shipping-day text');
}

{
  const t=setup();t.post([{id:'veg',qty:2}]);t.confirm(2);
  assert.equal(t.objects('商品主檔')[0].soldQty,2);
  t.status(2,'取消','已確認');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  t.status(2,'取消','取消','已完成訂單');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  assert.equal(JSON.parse(t.objects('已完成訂單')[0]['確認處理']).cancellation.phase,'done');
  t.status(2,'已確認','取消','已完成訂單');
  assert.equal(t.objects('已完成訂單')[0]['狀態'],'取消');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  console.log('PASS cancellation restores inventory once and canceled order cannot be reconfirmed');
}

{
  const t=setup();t.post([{id:'veg',qty:2}]);t.confirm(2);
  t.fail('商品主檔',true,7);t.status(2,'取消','已確認');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  t.status(2,'取消','已確認');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  console.log('PASS interrupted cancellation resumes without double restock');
}

{
  const t=setup();t.post([{id:'veg',qty:2}]);t.status(2,'取消','新訂單');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  assert.equal(t.objects('訂單總表').length,0);
  assert.equal(t.objects('已完成訂單')[0]['狀態'],'取消');
  console.log('PASS canceling an unconfirmed order does not change inventory');
}

{
  const t=setup();t.post([{id:'veg',qty:2}]);t.confirm(2);
  t.status(2,'已完成','已確認');
  assert.equal(t.objects('訂單總表').length,0);
  assert.equal(t.objects('已完成訂單')[0]['狀態'],'已完成');
  assert.equal(t.objects('出貨明細')[0]['狀態'],'已完成');
  t.status(2,'取消','已完成','已完成訂單');
  assert.equal(t.objects('商品主檔')[0].soldQty,0);
  assert.equal(t.objects('已完成訂單')[0]['狀態'],'取消');
  assert.equal(t.objects('記帳總表')[0]['狀態'],'取消');
  console.log('PASS completed orders move to archive and can later be canceled safely');
}
