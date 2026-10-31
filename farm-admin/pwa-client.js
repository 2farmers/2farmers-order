/* Session credentials live in memory only. No member records or photos are cached. */
(() => {
 'use strict';
 const methods=['farmAdminGetProjects','farmAdminGetTrees','farmAdminGetHistory','farmAdminGetHistoryBySubject','farmAdminGetNewCropOptions','farmAdminGetNewTreePreview','farmAdminNewTreePreview','farmAdminCreateCropProject','farmAdminCreateTree','farmAdminSaveCrop','farmAdminSaveTree','farmAdminSaveEcology','farmAdminGetMemberUsage'];
 const writes=new Set(['farmAdminCreateCropProject','farmAdminCreateTree','farmAdminSaveCrop','farmAdminSaveTree','farmAdminSaveEcology']);
 const loginUrl='https://script.google.com/macros/s/AKfycbyseJZzdU4i9w8v6mDyge1Vfc0-V6TlVaLYexk7Av9HSg3aPiyON_D5gdYdKiSYYoLsBg/exec?pwaLogin=1';
 const key='liangkou-farm-pwa-api-v1';let token='',endpoint='',popup=null,expiry=0,started=false,pending=0,loggingOut=false;
 const $=id=>document.getElementById(id);
 try{endpoint=localStorage.getItem(key)||'';}catch(ignore){}
 function validEndpoint(v){return /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(v);}
 function show(text){$('pwaAuthStatus').textContent=text;}
 function connection(){const online=navigator.onLine;$('pwaConnection').textContent=online?'已連線':'目前離線，請恢復連線後儲存';$('pwaConnection').classList.toggle('offline',!online);}
 async function request(method,args,credential=token){
  if(!validEndpoint(endpoint))throw Error('請先設定 PWA 資料連線網址。');
  if(!navigator.onLine)throw Error('目前離線，內容尚未送出，請恢復連線後再試。');
  if(method!=='exchange'&&(!credential||Date.now()>=expiry)){show('登入已逾時，請重新登入；尚未送出的內容仍留在畫面。');$('pwaAuth').hidden=false;throw Error('請重新登入後再儲存。');}
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),90000);let result;
  pending++;
  try{
   const response=await fetch(endpoint,{method:'POST',mode:'cors',credentials:'omit',redirect:'follow',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({v:1,method,args,token:credential}),signal:controller.signal});
   if(!response.ok)throw Error('HTTP');result=await response.json();
  }catch(ignore){throw Error(writes.has(method)?'未能取得儲存結果，資料可能已送達。請保留目前內容與識別碼，確認紀錄後再重試。':'無法連接管理後台。請確認部署與網路後再試。');}
  finally{clearTimeout(timer);pending--;}
  if(!result||result.v!==1||typeof result.ok!=='boolean')throw Error('管理後台版本不符，請確認使用新版 PWA API 部署網址。');
  if(!result.ok){if(result.code==='AUTH_REQUIRED'||result.code==='AUTH_EXPIRED'){token='';$('pwaAuth').hidden=false;show(result.error);}throw Error(result.error||'操作未完成。');}
  return result.data;
 }
 function runner(success,failure){const api={withSuccessHandler:cb=>runner(cb,failure),withFailureHandler:cb=>runner(success,cb)};methods.forEach(method=>{api[method]=(...args)=>{request(method,args).then(data=>{if(success)success(data);},err=>{if(failure)failure(err);else show(err.message);});};});return Object.freeze(api);}
 window.farmPwaRun=runner(null,null);
 async function acceptTicket(ticket){
  if(loggingOut)return;show('確認登入中……');
  try{const session=await request('exchange',[ticket],'');if(loggingOut)return;token=session.token;expiry=session.expiresAt;$('pwaAuth').hidden=true;$('pwaWorkspace').hidden=false;$('pwaLogout').hidden=false;show('');$('pwaSession').textContent='已登入・'+new Intl.DateTimeFormat('zh-TW',{timeZone:'Asia/Taipei',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).format(new Date(expiry))+' 前有效';
   if(!started){started=true;loadProjects();loadTrees();loadNewCropOptions();loadNewTreePreview();loadMemberUsage();}
  }catch(e){show(e.message+' 若登入連結已使用或逾時，請重新登入。');}
 }
 document.addEventListener('DOMContentLoaded',()=>{
  $('pwaEndpoint').value=endpoint;connection();window.addEventListener('online',connection);window.addEventListener('offline',connection);
  $('pwaSaveEndpoint').addEventListener('click',()=>{const value=$('pwaEndpoint').value.trim();if(!validEndpoint(value)){show('請貼上 Apps Script 正式 /exec 網址。');return;}if(value!==endpoint){if(started){show('更換資料連線前，請先儲存內容並登出。');return;}token='';expiry=0;}endpoint=value;try{localStorage.setItem(key,endpoint);}catch(ignore){}show('連線網址已儲存，現在可以登入。');});
  $('pwaLogin').addEventListener('click',()=>{if(!validEndpoint(endpoint)){show('請先展開「首次設定」，貼上新版 PWA API 網址並儲存。');$('pwaSetup').open=true;return;}popup=window.open(loginUrl,'liangkou-farm-login');show('請在 Google 帳號頁面點「進入田間管理 PWA」。若瀏覽器擋住新視窗，請允許此網站開啟視窗。');});
  window.addEventListener('message',e=>{if(e.origin!==location.origin||!popup||e.source!==popup||!e.data||e.data.type!=='liangkou-farm-ticket'||typeof e.data.ticket!=='string')return;acceptTicket(e.data.ticket);});
  const params=new URLSearchParams(location.hash.slice(1)),ticket=params.get('ticket');if(ticket){history.replaceState(null,'',location.pathname+location.search);acceptTicket(ticket);}
  $('pwaLogout').addEventListener('click',async()=>{
   if(pending||window.farmPwaHasPending()){show('目前有資料傳送中，請等待結果後再登出。');$('pwaAuth').hidden=false;return;}
   if(!confirm('登出會清除這次畫面中的資料與尚未送出的內容，確定登出？'))return;
   loggingOut=true;$('pwaLogout').disabled=true;
   if(!token||Date.now()>=expiry){token='';expiry=0;window.farmPwaAllowLeave=true;location.replace(location.pathname);return;}
   try{await request('logout',[]);}catch(e){loggingOut=false;$('pwaLogout').disabled=false;show('尚未確認登出：'+e.message+' 請保持此頁，恢復連線後重試。');$('pwaAuth').hidden=false;return;}
   token='';expiry=0;window.farmPwaAllowLeave=true;location.replace(location.pathname);
  });
 let installEvent=null,updating=false;window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();installEvent=e;$('pwaInstall').hidden=false;});
  $('pwaInstall').addEventListener('click',async()=>{if(!installEvent)return;await installEvent.prompt();await installEvent.userChoice;installEvent=null;$('pwaInstall').hidden=true;});
  if('serviceWorker'in navigator){navigator.serviceWorker.register('./sw.js',{scope:'./'}).then(reg=>{
   const offer=()=>{if(reg.waiting&&navigator.serviceWorker.controller)$('pwaUpdate').hidden=false;};offer();reg.addEventListener('updatefound',()=>{const worker=reg.installing;if(worker)worker.addEventListener('statechange',()=>{if(worker.state==='installed')offer();});});
   $('pwaReload').addEventListener('click',()=>{if(pending||window.farmPwaHasPending()){alert('資料傳送中，請等待結果後再更新。');return;}if(!confirm('更新會重新載入，請先確認田間紀錄都已儲存。'))return;updating=true;if(reg.waiting){reg.waiting.postMessage({type:'ACTIVATE'});}else location.reload();});
  }).catch(()=>{$('pwaInstallHint').textContent='目前無法啟用安裝與離線開啟，可先使用瀏覽器操作。';});
   navigator.serviceWorker.addEventListener('controllerchange',()=>{if(updating){window.farmPwaAllowLeave=true;location.reload();}});
  }
 });
})();
