const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const path = require('node:path');
const root = path.join(__dirname, '..');
const calls=[];const props=new Map();
const context = {
  ContentService: {MimeType:{JSON:'json'},createTextOutput:text=>({text,setMimeType(){return this;}})},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k),setProperty:(k,v)=>props.set(k,v)})},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
  Utilities:{getUuid:()=>crypto.randomUUID(),computeHmacSha256Signature:(s,key)=>crypto.createHmac('sha256',key).update(s).digest(),base64EncodeWebSafe:x=>Buffer.from(x).toString('base64url')},
};
const names=['lookupMember','verifyExistingMember','vegetablePlanOptions','fruitApplicationOptions','submitFruitApplication','submitVegetableApplication','reportFruitPayment','reportVegetablePayment','submitMemberMessage','submitPollAnswer','setFarmVisitRegistration'];
for(const name of names) context[name]=(...args)=>{calls.push({name,args});if(name==='lookupMember'&&args[1]!=='valid')throw Error('驗證失敗');return name.startsWith('submit')?{id:'TEST-1'}:{name};};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(root,'apps-script/CoFarmApi.gs'),'utf8'),context);
const post=(method,args=[],extra={})=>JSON.parse(context.doPost({postData:{contents:JSON.stringify({version:1,method,args,...extra})}}).text);
assert.equal(post('ping').ok,true);
for(const method of ['onOpen','issueMemberCode','createVegetableBoxBatch','constructor','__proto__'])assert.equal(post(method).ok,false);
assert.equal(post('lookupMember',['M','invalid']).ok,false);
assert.equal(post('lookupMember',['M','valid']).ok,true);
assert.equal(post('lookupMember',['M']).ok,false);
const app=post('submitFruitApplication',[{}]);assert.ok(app.paymentToken);
const before=calls.length;
assert.equal(post('reportFruitPayment',['TEST-1','郵局','12345']).ok,false);
assert.equal(post('reportFruitPayment',['OTHER','郵局','12345'],{paymentToken:app.paymentToken}).ok,false);
assert.equal(post('reportVegetablePayment',['TEST-1','郵局','12345'],{paymentToken:app.paymentToken}).ok,false);
assert.equal(calls.length,before);
assert.equal(post('reportFruitPayment',['TEST-1','郵局','12345'],{paymentToken:app.paymentToken}).ok,true);
assert.equal(post('reportFruitPayment',['TEST-1','郵局','12345'],{paymentToken:'1.fake'}).ok,false);
assert.equal(JSON.parse(context.doPost({postData:{contents:'{oops'}}).text).ok,false);
const html=fs.readFileSync(path.join(root,'cofarm/index.html'),'utf8');
for(const m of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))new vm.Script(m[1]);
assert.ok(!html.includes('google.script.run'));
assert.ok(!html.includes('<iframe'));
async function clientTests(){
  const sent=[];let mode='ok';
  const browser={window:{COFARM_CONFIG:{endpoint:'https://script.google.com/macros/s/test/exec'}},document:{getElementById:()=>null},AbortController,setTimeout,clearTimeout,Map,Set,
    fetch:async(url,options)=>{sent.push(options);const req=JSON.parse(options.body);if(mode==='network')throw Error('network');return{ok:true,json:async()=>mode==='wrong-version'?{ok:true}:{version:1,ok:true,data:{id:'TEST-1'},paymentToken:'test-token'}};}};
  vm.createContext(browser);vm.runInContext(fs.readFileSync(path.join(root,'cofarm/api.js'),'utf8'),browser);
  const invoke=(name,...args)=>new Promise((resolve,reject)=>browser.window.cofarmRun.withSuccessHandler(resolve).withFailureHandler(reject)[name](...args));
  await Promise.all([invoke('lookupMember','M','valid'),invoke('vegetablePlanOptions')]);
  assert.equal(sent[0].credentials,'omit');assert.equal(sent[0].headers['Content-Type'],'text/plain;charset=utf-8');
  assert.equal(JSON.parse(sent[0].body).args[1],'valid');
  await assert.rejects(invoke('reportFruitPayment','TEST-1','郵局','12345'),/驗證已失效/);
  await invoke('submitFruitApplication',{});await invoke('reportFruitPayment','TEST-1','郵局','12345');
  assert.equal(JSON.parse(sent.at(-1).body).paymentToken,'test-token');
  mode='network';const count=sent.length;await assert.rejects(invoke('submitVegetableApplication',{}),/避免重複/);assert.equal(sent.length,count+1);
  mode='wrong-version';await assert.rejects(invoke('ping'),/後台版本/);
}
clientTests().then(()=>console.log('PASS: routes, credential delegation, payment token binding, callback isolation, transport errors, no retries, HTML syntax.')).catch(error=>{console.error(error);process.exitCode=1;});
