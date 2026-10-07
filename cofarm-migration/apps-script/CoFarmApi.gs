/** Add this file to the existing bound Apps Script project. Keep Code.gs and Index.html. */
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > 30000) throw new Error('請求內容不完整或過長。');
    const request = JSON.parse(raw);
    if (request.version !== 1 || !Array.isArray(request.args)) throw new Error('請更新共耕網頁後重試。');
    const routes = {
      ping: {call: function(){return {ready:true};}, count: 0},
      lookupMember: {call: lookupMember, count: 2},
      verifyExistingMember: {call: verifyExistingMember, count: 2},
      vegetablePlanOptions: {call: vegetablePlanOptions, count: 0},
      fruitApplicationOptions: {call: fruitApplicationOptions, count: 0},
      submitFruitApplication: {call: submitFruitApplication, count: 1},
      submitVegetableApplication: {call: submitVegetableApplication, count: 1},
      reportFruitPayment: {call: reportFruitPayment, count: 3},
      reportVegetablePayment: {call: reportVegetablePayment, count: 3},
      submitMemberMessage: {call: submitMemberMessage, count: 4},
      submitPollAnswer: {call: submitPollAnswer, count: 4},
      setFarmVisitRegistration: {call: setFarmVisitRegistration, count: 4}
    };
    if (!Object.prototype.hasOwnProperty.call(routes,request.method)) throw new Error('不支援的操作。');
    const route=routes[request.method];
    if (request.args.length !== route.count) throw new Error('請求資料不完整。');
    const paymentKind=request.method==='reportFruitPayment'?'fruit':
      request.method==='reportVegetablePayment'?'vegetable':'';
    if (paymentKind) cofarmCheckPaymentToken_(paymentKind,request.args[0],request.paymentToken);
    if(request.method==='submitFruitApplication'||request.method==='submitVegetableApplication') cofarmSecret_();
    const result=route.call.apply(null,request.args);
    const output={version:1,ok:true,data:result};
    const newKind=request.method==='submitFruitApplication'?'fruit':
      request.method==='submitVegetableApplication'?'vegetable':'';
    if(newKind && result && result.id) output.paymentToken=cofarmPaymentToken_(newKind,result.id);
    return cofarmJson_(output);
  } catch(error) {
    // No request payloads, credentials, stack traces or private records are logged.
    return cofarmJson_({version:1,ok:false,error:error instanceof SyntaxError?'請求格式不正確。':String(error.message||'操作未完成。')});
  }
}
function cofarmJson_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
function cofarmSecret_() {
  const properties=PropertiesService.getScriptProperties();
  let secret=properties.getProperty('COFARM_PAYMENT_SECRET');
  if(secret)return secret;
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    secret=properties.getProperty('COFARM_PAYMENT_SECRET');
    if(!secret){secret=Utilities.getUuid()+Utilities.getUuid();properties.setProperty('COFARM_PAYMENT_SECRET',secret);}
    return secret;
  }finally{lock.releaseLock();}
}
function cofarmSign_(kind,id,expiry) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(
    kind+'|'+String(id)+'|'+expiry,cofarmSecret_())).replace(/=+$/,'');
}
function cofarmPaymentToken_(kind,id) {
  const expiry=String(Date.now()+7*24*60*60*1000);
  return expiry+'.'+cofarmSign_(kind,id,expiry);
}
function cofarmCheckPaymentToken_(kind,id,token) {
  const parts=String(token||'').split('.');
  if(parts.length!==2||!/^\d+$/.test(parts[0])||Number(parts[0])<Date.now())
    throw new Error('付款回報驗證已失效，請聯絡倆口田並提供申請編號。');
  const expected=cofarmSign_(kind,id,parts[0]);
  let difference=expected.length^parts[1].length;
  for(let i=0;i<expected.length;i++)difference|=expected.charCodeAt(i)^(parts[1].charCodeAt(i)||0);
  if(difference)throw new Error('付款回報驗證不正確。');
}
