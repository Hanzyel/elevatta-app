/* Elevatta REV36 — Microsoft 365 / SharePoint + offline-first resilient outbox.
   Auth: server-side delegated OAuth. Refresh token stays in an HttpOnly cookie;
   browser receives only a short-lived access token. */
window.ElevattaM365=(function(){
  'use strict';
  const M={};
  const VERSION='40.1.0';
  const K={
    tenant:'elevatta_m365_tenant_v1',client:'elevatta_m365_client_v1',workbook:'elevatta_m365_workbook_v1',
    locator:'elevatta_m365_locator_v1',account:'elevatta_m365_account_v1',access:'elevatta_m365_access_v2',expires:'elevatta_m365_expires_v2',
    lastCatalog:'elevatta_m365_catalog_sync_v1',catalogDirty:'elevatta_m365_catalog_dirty_v1'
  };
  const TABLE={ap:"tbApontamentos",fvs:"tbFvsRegistros",items:"tbFvsItens",med:"tbMedicoes",photos:"tbFotos",packs:"tbPacotes",funcs:"tbFuncionarios",config:"tbConfig",obras:"tbObras",equipes:"tbEquipes",frentes:"tbFrentesServico",tools:"tbFerramentas",toolTerms:"tbTermosFerramentas",toolMoves:"tbMovimentosFerramentas",imports:"tbImportacoes",importIssues:"tbImportacaoOcorrencias",audit:"tbAuditoria",users:"tbUsuarios"};
  const HEAD={ap:["ID","DATA","RESPONSAVEL","MATRICULA","NOME","FUNCAO","REGIME","OBRA","ETAPA","PACOTE_COD","PACOTE_SERVICO","PAVIMENTO","EQUIPE","FRACAO_DIA","CLASSIFICACAO","SITUACAO","OBSERVACAO","MOTIVO_FALTA","SAIU_CEDO","UPDATED_AT","FRENTE_ID","LOCAL_FRENTE","DEVICE_ID","SYNC_VERSION"],fvs:["ID","DATA","OBRA","RESPONSAVEL","PACOTE_COD","PACOTE_SERVICO","ETAPA","TIPO","LOCAL","SITUACAO","CONFORMES","NAO_CONFORMES","NA","PENDENTES","QTD_FOTOS","OBSERVACAO","UPDATED_AT","FRENTE_ID","PAVIMENTO","DEVICE_ID","SYNC_VERSION"],items:["ID","REGISTRO_ID","ITEM","VERIFICACAO","CRITERIO","STATUS","OBSERVACAO","QTD_FOTOS","UPDATED_AT","FRENTE_ID"],med:["ID","REGISTRO_ID","DATA","OBRA","PACOTE_COD","PACOTE_SERVICO","LOCAL","ANTERIOR_PCT","EXECUTADO_PCT","APROVADO_PCT","MEDIDO_PCT","ACUMULADO_PCT","SALDO_PCT","PENDENTE_QUALIDADE_PCT","APROVADO_NAO_MEDIDO_PCT","OBSERVACAO","UPDATED_AT","FRENTE_ID","PAVIMENTO"],photos:["ID","REGISTRO_ID","DATA","OBRA","PACOTE_COD","PACOTE_SERVICO","ITEM","STATUS","LEGENDA","MARCACOES_QTD","ARQUIVO_URL","ARQUIVO_NOME","UPDATED_AT","FRENTE_ID","ENTIDADE_TIPO"],packs:["CODIGO","PACOTE_SERVICO","ETAPA","OBRA_ESPECIFICA","STATUS","GLOBAL","UPDATED_AT"],funcs:["MATRICULA","NOME","FUNCAO","REGIME","ATIVO","UPDATED_AT","FUNCIONARIO_ID"]};
  const requiredTables=["tbApontamentos","tbFvsRegistros","tbFvsItens","tbMedicoes","tbFotos","tbPacotes","tbFuncionarios","tbConfig","tbObras","tbEquipes","tbFrentesServico","tbFerramentas","tbTermosFerramentas","tbMovimentosFerramentas","tbImportacoes","tbImportacaoOcorrencias","tbAuditoria","tbUsuarios"];
  const $=id=>document.getElementById(id);
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const ls=(k,d='')=>{try{return localStorage.getItem(k)??d}catch{return d}};
  const lset=(k,v)=>{try{localStorage.setItem(k,v);return true}catch{return false}};
  const ss=(k,d='')=>{try{return sessionStorage.getItem(k)??d}catch{return d}};
  const sset=(k,v)=>{try{sessionStorage.setItem(k,v);return true}catch{return false}};
  const sdel=k=>{try{sessionStorage.removeItem(k)}catch{}};
  const now=()=>new Date().toISOString();
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  const toast=(t,ms=2200)=>{try{if(window.RF?.toast)return RF.toast(t,ms)}catch{};console.log('[Elevatta]',t)};
  const sanitize=s=>String(s||'arquivo').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9._-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,100)||'arquivo';
  const upper=s=>String(s??'').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');
  function fnv1a(str,seed=0x811c9dc5){let h=seed>>>0;for(let i=0;i<str.length;i++){h^=str.charCodeAt(i);h=Math.imul(h,0x01000193)>>>0}return h>>>0}
  function hash12(str){const z=String(str);return fnv1a(z,0x811c9dc5).toString(16).padStart(8,'0')+fnv1a(z,0x9e3779b9).toString(16).padStart(8,'0').slice(0,4)}
  function frontId(x={}){if(window.ElevattaERP?.frontId)return ElevattaERP.frontId(x);const o=upper(x.obra),p=upper(x.pacote||x.pacoteCod),pv=upper(x.pavimento),lf=upper(x.local||x.localFrente);if(!o||!p||(!pv&&!lf))return '';const k=[o,p,pv,lf].join('|');return 'FRT-'+hash12(k).toUpperCase()}
  function employeeId(m){if(window.ElevattaERP?.employeeId)return ElevattaERP.employeeId(m);const k=upper(m);return k?'FUN-'+hash12(k).toUpperCase():''}
  function deviceId(){return window.ElevattaERP?.deviceId?.()||(()=>{let x=ls('elevatta_device_id_v1','');if(!x){x='DEV-'+Date.now().toString(36).toUpperCase()+Math.random().toString(36).slice(2,8).toUpperCase();lset('elevatta_device_id_v1',x)}return x})()}
  let serverCfg={loaded:false,authConfigured:false,tenant:'',clientId:'',workbookUrl:'',authMode:''};
  let workbookSessionId='',workbookSessionAt=0,workbookSessionPromise=null;
  let flushing=null;
  const syncState={busy:false,lastError:'',lastOkAt:0,storagePersistent:null,usage:0,quota:0};
  let lastConfigError='';

  function cfg(){return {tenant:ls(K.tenant,'organizations').trim()||'organizations',client:ls(K.client,'').trim(),workbook:ls(K.workbook,'').trim()}}
  M.isConfigured=()=>Boolean(serverCfg.authConfigured&&cfg().workbook);
  function locator(){try{return JSON.parse(ls(K.locator,''))||null}catch{return null}}
  function saveLocator(x){lset(K.locator,JSON.stringify(x||null));workbookSessionId='';workbookSessionAt=0;workbookSessionPromise=null}
  function tokenInfo(){return {access:ss(K.access),expires:Number(ss(K.expires,'0'))||0}}
  function clearAccess(){sdel(K.access);sdel(K.expires)}
  function setToken(j){if(j?.access_token)sset(K.access,j.access_token);sset(K.expires,String(Date.now()+Math.max(60,Number(j?.expires_in)||3600)*1000))}
  function account(){try{return JSON.parse(ls(K.account,'{}'))||{}}catch{return {}}}

  async function fetchTimed(url,init={},ms=12000){
    const ctl=new AbortController(),timer=setTimeout(()=>ctl.abort(),ms);
    try{return await fetch(url,Object.assign({},init,{signal:ctl.signal}));}
    catch(e){if(e?.name==='AbortError')throw new Error('Tempo de conexão esgotado. Os dados continuam protegidos no aparelho.');throw e}
    finally{clearTimeout(timer)}
  }
  function retryMs(r,attempt){const h=Number(r?.headers?.get?.('retry-after'));return Number.isFinite(h)&&h>0?Math.min(15000,h*1000):Math.min(5000,600*Math.pow(2,attempt));}

  async function loadServerConfig(force=false){
    if(serverCfg.loaded&&!force)return serverCfg;
    try{
      const r=await fetchTimed('/api/m365/config',{cache:'no-store',credentials:'same-origin'},10000);
      const j=await r.json().catch(()=>({}));
      serverCfg={loaded:true,authConfigured:!!j.authConfigured,tenant:String(j.tenant||''),clientId:String(j.clientId||''),workbookUrl:String(j.workbookUrl||''),authMode:String(j.authMode||'')};
      lastConfigError='';
      if(serverCfg.tenant)lset(K.tenant,serverCfg.tenant);
      if(serverCfg.clientId)lset(K.client,serverCfg.clientId);
      if(serverCfg.workbookUrl&&cfg().workbook!==serverCfg.workbookUrl){lset(K.workbook,serverCfg.workbookUrl);saveLocator(null)}
    }catch(e){lastConfigError=String(e?.message||e);serverCfg=Object.assign(serverCfg,{loaded:false});console.warn('Config M365 indisponível; tentarei novamente quando a conexão voltar:',e)}
    return serverCfg;
  }

  async function login(){
    if(!navigator.onLine){open('connect');setMsg('Sem internet. Continue trabalhando offline e faça o login quando a conexão voltar.',false);return false}
    await loadServerConfig(true);
    if(!serverCfg.authConfigured){open('connect');setMsg('Configure M365_CLIENT_ID, M365_CLIENT_SECRET e SESSION_SECRET no Render.',false);return false}
    const ret=location.pathname+location.search.replace(/([?&])m365=[^&]*/g,'$1').replace(/[?&]$/,'')+location.hash;
    location.assign('/auth/login?returnTo='+encodeURIComponent(ret||'/'));
    return true;
  }
  async function refreshToken(){
    const r=await fetchTimed('/api/m365/token',{method:'GET',cache:'no-store',credentials:'same-origin'},12000);
    const j=await r.json().catch(()=>({}));
    if(!r.ok||!j.access_token){clearAccess();const e=new Error(j.error||'Entre com a conta Microsoft 365.');e.reauthRequired=!!j.reauthRequired;throw e}
    setToken(j);return j.access_token;
  }
  async function getToken(interactive=false){
    const t=tokenInfo();if(t.access&&t.expires>Date.now()+90000)return t.access;
    try{return await refreshToken()}catch(e){if(interactive){await login();return null}throw e}
  }
  M.login=login;
  M.logout=async()=>{try{await fetchTimed('/auth/logout?json=1',{cache:'no-store',credentials:'same-origin'},8000)}catch{}clearAccess();lset(K.account,'');updateStatus();toast('Sessão Microsoft encerrada')};

  async function handleRedirect(){
    const u=new URL(location.href),mode=u.searchParams.get('m365');if(!mode)return;
    const msg=u.searchParams.get('message')||'';u.searchParams.delete('m365');u.searchParams.delete('message');history.replaceState(history.state||{},'',u.pathname+(u.search?'?'+u.searchParams.toString():'')+u.hash);
    if(mode==='connected'){
      clearAccess();try{await getToken(false);await loadProfile();updateStatus();toast('Microsoft 365 conectado ✓');setTimeout(()=>flushOutbox({force:true}).catch(()=>{}),120)}catch(e){open('connect');setMsg(e.message||String(e),false)}
    }else if(mode==='error'){open('connect');setMsg(msg||'Não foi possível concluir o login Microsoft.',false)}
    else if(mode==='logout'){updateStatus()}
  }

  async function graph(path,opt={}){
    const token=await getToken(!!opt.interactive);if(!token)throw new Error('Login Microsoft iniciado.');
    const init={method:opt.method||'GET',headers:Object.assign({Authorization:'Bearer '+token},opt.headers||{})};
    const workbookCall=/\/workbook\//i.test(String(path))&&!/\/workbook\/(createSession|closeSession)/i.test(String(path));
    if(workbookCall&&!opt.noWorkbookSession){try{const sid=await ensureWorkbookSession();if(sid)init.headers['workbook-session-id']=sid}catch(e){console.warn('Sessão Excel indisponível; seguindo sem sessão:',e)}}
    if(opt.body!==undefined){if(opt.raw)init.body=opt.body;else{init.headers['Content-Type']='application/json';init.body=JSON.stringify(opt.body)}}
    const url=path.startsWith('http')?path:'https://graph.microsoft.com/v1.0'+path;
    let r=null;
    for(let attempt=0;attempt<3;attempt++){
      r=await fetchTimed(url,init,opt.raw?45000:22000);
      if(![429,502,503,504].includes(r.status)||attempt===2)break;
      await sleep(retryMs(r,attempt));
    }
    if(r.status===401&&!opt._retry){clearAccess();await refreshToken();return graph(path,Object.assign({},opt,{_retry:true,interactive:false}))}
    if(r.status===204)return null;
    const ct=r.headers.get('content-type')||'';const j=ct.includes('json')?await r.json().catch(()=>({})):await r.blob();
    if(!r.ok){
      const code=String(j?.error?.code||''),msg=j?.error?.message||j?.error_description||('Microsoft Graph: '+r.status);
      if(workbookCall&&!opt._sessionRetry&&/session/i.test(code+' '+msg)){workbookSessionId='';workbookSessionAt=0;workbookSessionPromise=null;return graph(path,Object.assign({},opt,{_sessionRetry:true}))}
      throw new Error(msg)
    }
    return j;
  }
  async function loadProfile(){try{const me=await graph('/me?$select=displayName,mail,userPrincipalName');const a={name:me.displayName||'',email:me.mail||me.userPrincipalName||''};lset(K.account,JSON.stringify(a));return a}catch{return null}}

  function b64urlBytes(bytes){let bin='';bytes.forEach(b=>bin+=String.fromCharCode(b));return btoa(bin).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'')}
  function encodeShareUrl(url){return 'u!'+b64urlBytes(new TextEncoder().encode(url))}
  async function resolveWorkbook(force=false){
    const c=cfg();if(!c.workbook)throw new Error('A URL da planilha BASE_ELEVATTA_SHAREPOINT.xlsx não está configurada no Render.');
    if(!force){const x=locator();if(x&&x.share===c.workbook&&x.driveId&&x.itemId)return x}
    const item=await graph('/shares/'+encodeURIComponent(encodeShareUrl(c.workbook))+'/driveItem?$select=id,name,webUrl,parentReference');
    const x={share:c.workbook,driveId:item.parentReference?.driveId,itemId:item.id,parentId:item.parentReference?.id,name:item.name,webUrl:item.webUrl};
    if(!x.driveId||!x.itemId)throw new Error('Não consegui identificar a planilha no SharePoint. Confira M365_WORKBOOK_URL.');saveLocator(x);return x
  }
  async function ensureWorkbookSession(){
    if(workbookSessionId&&(Date.now()-workbookSessionAt)<240000)return workbookSessionId;
    if(workbookSessionPromise)return workbookSessionPromise;
    workbookSessionPromise=(async()=>{const l=await resolveWorkbook();const j=await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(l.itemId)}/workbook/createSession`,{method:'POST',body:{persistChanges:true},noWorkbookSession:true});workbookSessionId=String(j?.id||'');workbookSessionAt=Date.now();return workbookSessionId})().finally(()=>{workbookSessionPromise=null});
    return workbookSessionPromise;
  }
  async function tablePath(table,suffix=''){const l=await resolveWorkbook();return `/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(l.itemId)}/workbook/tables/${encodeURIComponent(table)}${suffix}`}
  async function listRows(table){let next=await tablePath(table,'/rows?$top=1000'),out=[],pages=0;while(next&&pages<100){const j=await graph(next);out.push(...((j?.value||[]).map(r=>({index:r.index,values:(r.values&&r.values[0])||[]}))));next=j?.['@odata.nextLink']||'';pages++}return out}
  function objRows(rows,headers){return rows.map(r=>{const o={__index:r.index};headers.forEach((h,i)=>o[h]=r.values[i]);return o}).filter(o=>String(o[headers[0]]??'').trim())}
  async function deleteIndexes(table,indexes){for(const i of [...indexes].sort((a,b)=>b-a))await graph(await tablePath(table,'/rows/'+i),{method:'DELETE'})}
  async function appendRows(table,rows){if(!rows?.length)return;await graph(await tablePath(table,'/rows'),{method:'POST',body:{values:rows.map(r=>Array.isArray(r)?r:[])}})}
  async function updateRow(table,index,row){await graph(await tablePath(table,'/rows/'+index),{method:'PATCH',body:{values:[Array.isArray(row)?row:[]]}})}
  /* Upsert primeiro, exclusão por último. Se a internet cair no meio, o Excel
     pode ficar com uma linha antiga extra, mas nunca perde todas as linhas do
     registro antes de receber as novas. A próxima sincronização reconcilia. */
  async function replaceWhere(table,headers,predicate,newRows){
    const scoped=objRows(await listRows(table),headers).filter(predicate),key=headers[0],byId=new Map();
    scoped.forEach(x=>{const id=String(x[key]??'');if(!byId.has(id))byId.set(id,[]);byId.get(id).push(x)});
    const keep=new Set();
    for(const row of (newRows||[])){const id=String(row?.[0]??'');keep.add(id);const group=byId.get(id)||[];if(group.length)await updateRow(table,group[0].__index,row);else await appendRows(table,[row]);}
    const stale=[];for(const x of scoped){const id=String(x[key]??''),group=byId.get(id)||[];if(!keep.has(id)||group.indexOf(x)>0)stale.push(x.__index)}
    if(stale.length)await deleteIndexes(table,stale);
  }
  async function upsertNoDelete(table,headers,newRows){const all=objRows(await listRows(table),headers),key=headers[0],byId=new Map();all.forEach(x=>{const id=String(x[key]??'');if(id&&!byId.has(id))byId.set(id,x)});for(const row of (newRows||[])){const id=String(row?.[0]??'');if(!id)continue;const old=byId.get(id);if(old)await updateRow(table,old.__index,row);else await appendRows(table,[row])}}

  async function testBase(){
    await loadServerConfig(true);setMsg('Conectando à Microsoft 365…',true);
    try{await getToken(true);if(!tokenInfo().access)return false;await loadProfile();const l=await resolveWorkbook(true);const t=await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(l.itemId)}/workbook/tables?$select=name`);const names=(t.value||[]).map(x=>x.name);const missing=requiredTables.filter(x=>!names.includes(x));if(missing.length)throw new Error('A planilha não tem as tabelas esperadas: '+missing.join(', '));setMsg('Conectado ✓ '+l.name,true);updateStatus();await flushOutbox({force:true});return true}catch(e){setMsg(e.message||String(e),false);updateStatus();return false}
  }
  M.test=testBase;

  // ---------- IndexedDB: arquivo local + outbox ----------
  let dbp=null;
  function db(){if(dbp)return dbp;dbp=new Promise((res,rej)=>{const q=indexedDB.open('elevatta-m365',2);q.onupgradeneeded=()=>{const d=q.result;if(!d.objectStoreNames.contains('outbox'))d.createObjectStore('outbox',{keyPath:'id'});if(!d.objectStoreNames.contains('records'))d.createObjectStore('records',{keyPath:'id'});if(!d.objectStoreNames.contains('photos'))d.createObjectStore('photos',{keyPath:'key'});if(!d.objectStoreNames.contains('apArchive'))d.createObjectStore('apArchive',{keyPath:'date'});if(!d.objectStoreNames.contains('meta'))d.createObjectStore('meta',{keyPath:'key'})};q.onsuccess=()=>{const d=q.result;d.onversionchange=()=>{try{d.close()}catch{}dbp=null};res(d)};q.onerror=()=>{dbp=null;rej(q.error)}});return dbp}
  async function idbGet(store,key){const d=await db();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).get(key);q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}
  async function idbAll(store){const d=await db();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).getAll();q.onsuccess=()=>res(q.result||[]);q.onerror=()=>rej(q.error)})}
  async function idbPut(store,val){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).put(val);tx.oncomplete=()=>res(val);tx.onerror=()=>rej(tx.error)})}
  async function idbDel(store,key){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).delete(key);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
  const PAGE_SYNC_OWNER='PAGE:'+((crypto&&crypto.randomUUID)?crypto.randomUUID():Math.random().toString(36).slice(2));
  async function acquireSyncLock(owner=PAGE_SYNC_OWNER,ttl=120000){const d=await db();return new Promise((res,rej)=>{let acquired=false;const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('sync-lock');q.onsuccess=()=>{const cur=q.result;if(!cur||Number(cur.expiresAt||0)<Date.now()||cur.owner===owner){acquired=true;os.put({key:'sync-lock',owner,expiresAt:Date.now()+ttl,updatedAt:now()})}};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res(acquired);tx.onerror=()=>rej(tx.error)})}
  async function releaseSyncLock(owner=PAGE_SYNC_OWNER){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('sync-lock');q.onsuccess=()=>{if(q.result?.owner===owner)os.delete('sync-lock')};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
  function dataUrlBlob(data){const m=String(data||'').match(/^data:([^;]+);base64,(.+)$/);if(!m)throw new Error('Foto inválida');const bin=atob(m[2]),a=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)a[i]=bin.charCodeAt(i);return new Blob([a],{type:m[1]})}
  async function srcBlob(src){if(src instanceof Blob)return src;if(/^data:/i.test(String(src||'')))return dataUrlBlob(src);const r=await fetch(src);if(!r.ok)throw new Error('Não consegui preservar uma foto no aparelho.');return r.blob()}
  function blobDataUrl(blob){return new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(fr.result);fr.onerror=()=>rej(fr.error);fr.readAsDataURL(blob)})}
  async function archiveFvsBundle(bundle){
    const recId=String(bundle?.rec?.id||'');if(!recId)throw new Error('Registro FVS sem ID.');
    const previous=await idbGet('records',recId).catch(()=>null),photos=[];const photoKeys=[];
    for(let i=0;i<(bundle.photos||[]).length;i++){
      const p=bundle.photos[i]||{},key=`REC:${recId}:${String(p.id||i)}`;const blob=await srcBlob(p.src);let uploadBlob=blob;try{uploadBlob=await markedBlob(p)}catch(e){console.warn('Não foi possível pré-renderizar marcações para sync em background:',e)}await idbPut('photos',{key,recordId:recId,blob,uploadBlob,createdAt:now()});photoKeys.push(key);photos.push(Object.assign({},p,{src:'',photoKey:key}))
    }
    const stored={id:recId,kind:'FVS',bundle:Object.assign({},bundle,{photos}),photoKeys,createdAt:previous?.createdAt||bundle?.rec?.createdAt||now(),updatedAt:now(),syncedAt:null};await idbPut('records',stored);
    const keep=new Set(photoKeys);for(const key of (previous?.photoKeys||[])){if(!keep.has(key))await idbDel('photos',key).catch(()=>{})}
    return stored
  }
  async function hydrateFvsStored(stored){
    if(!stored)return null;const b=Object.assign({},stored.bundle,{photos:[]});
    for(const p of (stored.bundle?.photos||[])){let src=p.src||'';if(!src&&p.photoKey){const x=await idbGet('photos',p.photoKey);if(x?.blob)src=await blobDataUrl(x.blob)}b.photos.push(Object.assign({},p,{src}))}
    return b
  }
  async function qPut(x){return idbPut('outbox',x)}
  async function qGet(id){return idbGet('outbox',id)}
  async function qAll(){return idbAll('outbox')}
  async function qDel(id){return idbDel('outbox',id)}
  async function migrateLegacyOutbox(){
    let list=[];try{list=await qAll()}catch{return 0}
    let moved=0;
    for(const x of list){
      if(x?.type!=='FVS'||x.recordId||!x.bundle?.rec?.id)continue;
      try{
        await archiveFvsBundle(x.bundle);
        await qPut({id:x.id,type:'FVS',recordId:String(x.bundle.rec.id),createdAt:x.createdAt||now(),updatedAt:now(),attempts:Number(x.attempts||0),nextAttemptAt:Number(x.nextAttemptAt||0),lastError:String(x.lastError||'')});
        moved++;
      }catch(e){console.warn('Pendência antiga mantida sem migração:',e)}
    }
    return moved;
  }
  async function queue(bundle){
    const id=String(bundle.queueId||bundle.id||('Q-'+Date.now()));const old=await qGet(id).catch(()=>null);let item;
    if(bundle.kind==='FVS'){await archiveFvsBundle(bundle);item={id,type:'FVS',recordId:String(bundle.rec.id),createdAt:old?.createdAt||now(),updatedAt:now(),attempts:old?.attempts||0,nextAttemptAt:0,lastError:''}}
    else if(bundle.kind==='AP_DAY'){const date=String(bundle.date||''),prev=await idbGet('apArchive',date).catch(()=>null),prevIds=[...(prev?.ownedIds||[]),...((prev?.bundle?.ownedIds)||[]),...((prev?.bundle?.rows||[]).map(r=>String(r?.[0]??'')))],curIds=(bundle.rows||[]).map(r=>String(r?.[0]??'')).filter(Boolean);bundle.ownedIds=[...new Set([...prevIds,...curIds].filter(Boolean))];await idbPut('apArchive',{date,bundle,ownedIds:bundle.ownedIds,updatedAt:now()});item={id,type:bundle.kind,bundle,createdAt:old?.createdAt||now(),updatedAt:now(),attempts:old?.attempts||0,nextAttemptAt:0,lastError:''}}
    else{item={id,type:bundle.kind,bundle,createdAt:old?.createdAt||now(),updatedAt:now(),attempts:old?.attempts||0,nextAttemptAt:0,lastError:''}}
    await qPut(item);await persistStorage().catch(()=>{});await requestBackgroundSync();updateStatus();return item
  }
  async function hydrateQueueItem(x){if(x.type==='FVS'){if(x.recordId)return hydrateFvsStored(await idbGet('records',x.recordId));if(x.bundle)return x.bundle}return x.bundle}
  async function markRecordSynced(id){const r=await idbGet('records',id);if(r){r.syncedAt=now();r.updatedAt=now();await idbPut('records',r);/* O blob já marcado só é necessário enquanto há pendência. Depois do envio, mantém-se o original + vetor de marcações para reedição e libera-se a cópia duplicada. */for(const key of (r.photoKeys||[])){const p=await idbGet('photos',key).catch(()=>null);if(p&&p.uploadBlob){delete p.uploadBlob;await idbPut('photos',p).catch(()=>{})}}}}
  async function requestBackgroundSync(){try{if(!('serviceWorker'in navigator))return;const reg=await navigator.serviceWorker.ready;if(reg.sync)await reg.sync.register('elevatta-outbox')}catch{}}
  async function persistStorage(){try{if(navigator.storage?.persisted)syncState.storagePersistent=await navigator.storage.persisted();if(syncState.storagePersistent!==true&&navigator.storage?.persist)syncState.storagePersistent=await navigator.storage.persist();if(navigator.storage?.estimate){const e=await navigator.storage.estimate();syncState.usage=Number(e.usage||0);syncState.quota=Number(e.quota||0)}}catch{}}
  function retryDelay(attempts){return Math.min(15*60*1000,5000*Math.pow(2,Math.min(7,Math.max(0,attempts-1))))}
  async function flushOutbox(opt={}){
    if(flushing)return flushing;
    flushing=(async()=>{
      syncState.busy=true;syncState.lastError='';await updateStatus().catch(()=>{});
      await loadServerConfig();if(!M.isConfigured()||!navigator.onLine){if(navigator.onLine&&lastConfigError)syncState.lastError=lastConfigError;return 0;}
      if(!(await acquireSyncLock()))return 0;
      const lockBeat=setInterval(()=>acquireSyncLock().catch(()=>{}),30000);
      try{
      let list=[];try{list=await qAll()}catch{return 0}list.sort((a,b)=>String(a.createdAt).localeCompare(String(b.createdAt)));
      let ok=0;
      for(const x of list){
        if(!opt.force&&Number(x.nextAttemptAt||0)>Date.now())continue;
        try{
          const b=await hydrateQueueItem(x);if(!b)throw new Error('Pendência local sem conteúdo.');
          if(x.type==='AP_DAY')await pushApBundle(b);else if(x.type==='FVS')await pushFvsBundle(b);else if(x.type==='CATALOG')await pushCatalogBundle(b);
          await qDel(x.id);if(x.type==='FVS'&&x.recordId)await markRecordSynced(x.recordId);if(x.type==='CATALOG'){lset(K.catalogDirty,'0');lset(K.lastCatalog,String(Date.now()))}syncState.lastError='';syncState.lastOkAt=Date.now();ok++;
        }catch(e){
          const attempts=Number(x.attempts||0)+1;x.attempts=attempts;x.lastError=String(e?.message||e);syncState.lastError=x.lastError;x.updatedAt=now();x.nextAttemptAt=Date.now()+retryDelay(attempts);await qPut(x).catch(()=>{});console.warn('M365 outbox:',e);break;
        }
      }
      updateStatus();if(ok&&!opt.silent)toast(ok+' pendência'+(ok>1?'s':'')+' sincronizada'+(ok>1?'s':'')+' ✓');return ok
      } finally { clearInterval(lockBeat);await releaseSyncLock().catch(()=>{}); }
    })().finally(()=>{syncState.busy=false;flushing=null;updateStatus().catch(()=>{});});
    return flushing;
  }
  M.flush=flushOutbox;
  M.pendingCount=async()=>{try{return (await qAll()).length}catch{return 0}};

  // ---------- Apontamento ----------
  function apBundle(){
    const a=window.AP?.auditSnapshot?.();if(!a)return null;const funcs=new Map((a.funcs||[]).map(x=>[String(x.m),x])),pacs=new Map((a.pacs||[]).map(x=>[String(x.c),x]));const upd=now(),dev=deviceId();
    const rows=(a.lanc||[]).filter(l=>String(l.p||'').trim()||l.st==='F').map((l,i)=>{const f=funcs.get(String(l.m))||{},p=pacs.get(String(l.p))||{},obra=l.obr||p.ob||'',fid=l.st==='F'?'':frontId({obra,pacote:l.p,pavimento:l.pv,local:l.lf});return ['AP-'+a.date+'-'+String(l.id||i),a.date,a.enc||'',l.m||'',f.n||'',f.f||'',f.r||'',obra,l.et||p.et||'',l.p||'',p.n||'',l.pv||'',l.eq||'',Number(l.fr||0),l.cl||'',l.sit||'',l.ob||'',l.st==='F'?(l.mf||'Falta'):'',l.sa?'SIM':'NAO',upd,fid,l.lf||'',dev,'40']});
    return {kind:'AP_DAY',queueId:'APDAY-'+a.date,date:a.date,rows}
  }
  /* Apontamento é reconciliado pelos IDs que este aparelho conhece. Isso evita que
     um aparelho offline apague, ao reconectar, linhas novas lançadas por outro
     aparelho no mesmo dia. Linhas já conhecidas e depois removidas continuam
     sendo excluídas normalmente. */
  async function pushApBundle(b){
    const current=(b.rows||[]),currentIds=new Set(current.map(r=>String(r?.[0]??''))),owned=new Set([...(b.ownedIds||[]).map(String),...currentIds]);
    const scoped=objRows(await listRows(TABLE.ap),HEAD.ap).filter(o=>String(o.DATA)===String(b.date)&&owned.has(String(o.ID||''))),byId=new Map();
    scoped.forEach(x=>{const id=String(x.ID||'');if(!byId.has(id))byId.set(id,[]);byId.get(id).push(x)});
    for(const row of current){const id=String(row?.[0]??''),group=byId.get(id)||[];if(group.length)await updateRow(TABLE.ap,group[0].__index,row);else await appendRows(TABLE.ap,[row])}
    const stale=[];for(const x of scoped){const id=String(x.ID||''),group=byId.get(id)||[];if(!currentIds.has(id)||group.indexOf(x)>0)stale.push(x.__index)}if(stale.length)await deleteIndexes(TABLE.ap,stale);return true
  }
  M.syncApontamento=async function(opt={}){const b=apBundle();if(!b)return false;await queue(b);if(!navigator.onLine||!M.isConfigured()){if(!opt.silent)toast('Apontamento salvo offline · sincronização automática pendente');return false}await flushOutbox({force:true,silent:true});const pending=!!(await qGet(b.queueId));if(!opt.silent)toast(pending?'Apontamento local · sincronização pendente':'Apontamento sincronizado com SharePoint ✓');return !pending};
  M.pullApontamentoDate=async function(date,opt={}){if(!M.isConfigured()||!navigator.onLine)return false;try{const rows=objRows(await listRows(TABLE.ap),HEAD.ap).filter(x=>String(x.DATA)===String(date));if(!rows.length)return false;const prefix='AP-'+date+'-',ownedIds=rows.map(x=>String(x.ID||'')).filter(Boolean);const lanc=rows.map(x=>{const rid=String(x.ID||'');return {id:rid.startsWith(prefix)?rid.slice(prefix.length):(rid||undefined),m:String(x.MATRICULA||''),p:String(x.PACOTE_COD||''),fr:Number(x.FRACAO_DIA||0),cl:String(x.CLASSIFICACAO||''),ob:String(x.OBSERVACAO||''),ex:false,st:String(x.MOTIVO_FALTA||'')?'F':'T',obr:String(x.OBRA||''),et:String(x.ETAPA||''),sit:String(x.SITUACAO||''),pv:String(x.PAVIMENTO||''),lf:String(x.LOCAL_FRENTE||''),sa:String(x.SAIU_CEDO||'').toUpperCase()==='SIM',mf:String(x.MOTIVO_FALTA||'')}});localStorage.setItem('ap_dia_'+date,JSON.stringify({enc:String(rows[0].RESPONSAVEL||''),by:'',as:null,ar:null,l:lanc}));await idbPut('apArchive',{date:String(date),ownedIds,bundle:{kind:'AP_DAY',queueId:'APDAY-'+date,date:String(date),rows:[],ownedIds},updatedAt:now()}).catch(()=>{});if(!opt.silent)toast('Dia carregado do SharePoint ✓');return true}catch(e){if(!opt.silent)toast('Não foi possível consultar o SharePoint');return false}};

  // ---------- Photos ----------
  function drawArrow(ctx,x1,y1,x2,y2,scale){const a=Math.atan2(y2-y1,x2-x1),h=Math.max(12,20*scale);const paint=(c,w)=>{ctx.strokeStyle=c;ctx.lineWidth=w;ctx.lineCap='round';ctx.lineJoin='round';ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.moveTo(x2,y2);ctx.lineTo(x2-h*Math.cos(a-Math.PI/6),y2-h*Math.sin(a-Math.PI/6));ctx.moveTo(x2,y2);ctx.lineTo(x2-h*Math.cos(a+Math.PI/6),y2-h*Math.sin(a+Math.PI/6));ctx.stroke()};paint('rgba(255,255,255,.94)',Math.max(7,10*scale));paint('#ff3b30',Math.max(3.5,5.5*scale))}
  async function markedBlob(p){const im=await new Promise((res,rej)=>{const x=new Image();x.onload=()=>res(x);x.onerror=()=>rej(new Error('Foto não pôde ser aberta'));x.src=p.src});const max=1280,r=Math.min(1,max/Math.max(im.naturalWidth,im.naturalHeight)),cv=document.createElement('canvas');cv.width=Math.max(1,Math.round(im.naturalWidth*r));cv.height=Math.max(1,Math.round(im.naturalHeight*r));const c=cv.getContext('2d',{alpha:false});c.drawImage(im,0,0,cv.width,cv.height);(p.marks||[]).forEach(m=>{if(m.t==='arrow')drawArrow(c,m.x1*cv.width,m.y1*cv.height,m.x2*cv.width,m.y2*cv.height,Math.min(cv.width,cv.height)/900);else if(m.t==='circle'){const cx=m.x*cv.width,cy=m.y*cv.height,rx=m.rx*cv.width,ry=m.ry*cv.height,sc=Math.min(cv.width,cv.height)/900;for(const [col,lw] of [['rgba(255,255,255,.94)',Math.max(7,10*sc)],['#ff3b30',Math.max(3.5,5.5*sc)]]){c.strokeStyle=col;c.lineWidth=lw;c.beginPath();c.ellipse(cx,cy,rx,ry,0,0,Math.PI*2);c.stroke()}}});return new Promise((res,rej)=>cv.toBlob(b=>b?res(b):rej(new Error('Falha ao preparar foto')),'image/jpeg',.76))}
  async function children(parentId){const l=await resolveWorkbook();const j=await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(parentId)}/children?$select=id,name,folder,webUrl&$top=200`);return j.value||[]}
  async function ensureFolder(parentId,name){const list=await children(parentId),found=list.find(x=>x.folder&&String(x.name).toLowerCase()===String(name).toLowerCase());if(found)return found;const l=await resolveWorkbook();return graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(parentId)}/children`,{method:'POST',body:{name,folder:{},'@microsoft.graph.conflictBehavior':'rename'}})}
  async function photoFolder(recordId){const l=await resolveWorkbook();const root=await ensureFolder(l.parentId,'ELEVATTA_FOTOS');return ensureFolder(root.id,sanitize(recordId))}
  async function uploadPhoto(recordId,p,i,folder){const l=await resolveWorkbook(),dest=folder||await photoFolder(recordId),blob=await markedBlob(p),name=String(i+1).padStart(2,'0')+'_'+sanitize(p.id||('foto_'+i))+'.jpg';return graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(dest.id)}:/${encodeURIComponent(name)}:/content`,{method:'PUT',body:blob,raw:true,headers:{'Content-Type':'image/jpeg'}})}
  async function cleanupPhotoFolder(folder,expected){if(!folder)return;const l=await resolveWorkbook(),keep=new Set(expected);for(const x of await children(folder.id)){if(!x.folder&&!keep.has(String(x.name||'')))await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(x.id)}`,{method:'DELETE'})}}
  async function deletePhotoFolder(recordId){try{const l=await resolveWorkbook(),base=(await children(l.parentId)).find(x=>x.folder&&x.name==='ELEVATTA_FOTOS');if(!base)return;const rec=(await children(base.id)).find(x=>x.folder&&x.name===sanitize(recordId));if(rec)await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(rec.id)}`,{method:'DELETE'})}catch(e){console.warn('photo folder delete',e)}}

  // ---------- FVS / Measurement ----------
  function fvsBundle(rec){rec=Object.assign({},rec,{deviceId:rec?.deviceId||deviceId()});const s=window.RF?.auditSnapshot?.()||{},items=window.ElevattaMeasurement?._getFvsRows?.()||[],q=s.quality||{},photos=(s.photos||[]).map(p=>({id:p.id,g:p.g,src:p.src,w:p.w,h:p.h,marks:p.marks||[],dt:p.dt||'',cap:p.cap||''}));return {kind:'FVS',queueId:'FVS-'+rec.id,rec,snapshot:{service:s.service||{},header:s.header||{},quality:q},items,photos}}
  function statusLabel(x){return x==='C'?'CONFORME':x==='NC'?'NÃO CONFORME':x==='NA'?'N/A':'PENDENTE'}
  async function pushFvsBundle(b){
    const r=b.rec,s=b.snapshot||{},q=s.quality||{},upd=now(),sit=Number(q.NC||0)>0?'NÃO CONFORME':Number(q.P||0)>0?'PENDENTE':'APROVADA',code=r.serviceCode||s.service?.code||'',name=r.serviceName||s.service?.name||'',stage=r.serviceStage||s.service?.etapa||'',fid=frontId({obra:r.obra,pacote:code,pavimento:r.pavimento,local:r.local}),dev=deviceId();
    await replaceWhere(TABLE.fvs,HEAD.fvs,o=>String(o.ID)===String(r.id),[[r.id,r.date||'',r.obra||'',r.responsavel||'',code,name,stage,r.type==='MEDICAO_FISICA'?'FVS + MEDIÇÃO FÍSICA':'SOMENTE FVS',r.local||'',sit,Number(q.C||0),Number(q.NC||0),Number(q.NA||0),Number(q.P||0),Number(r.photoCount||b.photos?.length||0),r.note||'',upd,fid,r.pavimento||'',dev,'40']]);
    const ir=(b.items||[]).map((x,i)=>[r.id+'-I-'+String(x.n||i+1),r.id,Number(x.n||i+1),x.desc||'',x.criterion||'',statusLabel(x.status),x.obs||'',Number(x.photos||0),upd,fid]);await replaceWhere(TABLE.items,HEAD.items,o=>String(o.REGISTRO_ID)===String(r.id),ir);
    if(r.type==='MEDICAO_FISICA'){const mr=[[r.id,r.id,r.date||'',r.obra||'',code,name,r.local||'',Number(r.previous||0)/100,Number(r.executed||0)/100,Number(r.approved||0)/100,Number(r.measured||0)/100,Number(r.accumulated||0)/100,Number(r.balance||0)/100,Number(r.pendingQuality||0)/100,Number(r.approvedUnmeasured||0)/100,r.note||'',upd,fid,r.pavimento||'']];await replaceWhere(TABLE.med,HEAD.med,o=>String(o.ID)===String(r.id),mr)}
    const photoRows=[],expectedNames=[],folder=(b.photos||[]).length?await photoFolder(r.id):null;
    for(let i=0;i<(b.photos||[]).length;i++){
      const p=b.photos[i];const file=await uploadPhoto(r.id,p,i,folder);if(!file?.id)throw new Error('A foto '+(i+1)+' não foi confirmada no SharePoint.');const appUrl='/api/erp/file?itemId='+encodeURIComponent(file.id)+'&name='+encodeURIComponent(file.name||('foto_'+i+'.jpg'));photoRows.push([p.id||r.id+'-P-'+i,r.id,r.date||'',r.obra||'',code,name,p.g||'','',p.cap||'',(p.marks||[]).length,appUrl,file.name||'',upd,fid,'FVS']);expectedNames.push(String(file.name||''));await sleep(15)
    }
    await replaceWhere(TABLE.photos,HEAD.photos,o=>String(o.REGISTRO_ID)===String(r.id),photoRows);if(folder)await cleanupPhotoFolder(folder,expectedNames);else await deletePhotoFolder(r.id);return true
  }
  M.queueFvsRecord=async function(rec){const b=fvsBundle(rec);const item=await queue(b);if(navigator.onLine&&M.isConfigured())setTimeout(()=>flushOutbox({force:true,silent:true}).catch(()=>{}),0);return item.id};
  M.saveFvsRecord=async function(rec){const queueId=await M.queueFvsRecord(rec);if(!navigator.onLine||!M.isConfigured())return false;await flushOutbox({force:true,silent:true});return !(await qGet(queueId))};
  M.syncCurrentMeasurementHistory=async function(opt={}){if(!M.isConfigured()||!navigator.onLine)return false;try{const s=window.RF?.auditSnapshot?.(),code=s?.service?.code||'';if(!code)return false;const obra=s?.header?.obraNome||$('r_h_obra2')?.value||$('r_h_obra')?.value||'';const remote=objRows(await listRows(TABLE.med),HEAD.med).filter(x=>String(x.PACOTE_COD)===String(code)&&String(x.OBRA)===String(obra));if(!remote.length)return false;let local=[];try{local=JSON.parse(localStorage.getItem('elevatta_medicoes_fisicas_v1')||'[]')}catch{}const ids=new Set(local.map(x=>x.id));const sk=s?.service?.key||$('r_svc')?.value||code,ok=($('r_h_obra')?.value==='OUT'?'OUT:'+String($('r_h_obra2')?.value||'').trim().toUpperCase():$('r_h_obra')?.value)||'SEM_OBRA';remote.forEach(x=>{if(ids.has(String(x.ID)))return;local.push({id:String(x.ID),type:'MEDICAO_FISICA',createdAt:String(x.UPDATED_AT||''),date:String(x.DATA||''),serviceKey:sk,serviceCode:String(x.PACOTE_COD||''),serviceName:String(x.PACOTE_SERVICO||''),obraKey:ok,obra:String(x.OBRA||''),pavimento:String(x.PAVIMENTO||''),local:String(x.LOCAL||''),previous:Number(x.ANTERIOR_PCT||0)*100,executed:Number(x.EXECUTADO_PCT||0)*100,approved:Number(x.APROVADO_PCT||0)*100,measured:Number(x.MEDIDO_PCT||0)*100,accumulated:Number(x.ACUMULADO_PCT||0)*100,balance:Number(x.SALDO_PCT||0)*100,pendingQuality:Number(x.PENDENTE_QUALIDADE_PCT||0)*100,approvedUnmeasured:Number(x.APROVADO_NAO_MEDIDO_PCT||0)*100,note:String(x.OBSERVACAO||'')});ids.add(String(x.ID))});localStorage.setItem('elevatta_medicoes_fisicas_v1',JSON.stringify(local.slice(-2000)));window.ElevattaMeasurement?._reload?.();if(!opt.silent)toast('Histórico de medição atualizado do SharePoint ✓');return true}catch(e){if(!opt.silent)toast('Não foi possível consultar as medições');return false}};

  // ---------- Catalog ----------
  async function pullCatalog(opt={}){if(!navigator.onLine){toast('Sem internet · usando os cadastros deste aparelho');return false}if(ls(K.catalogDirty,'0')==='1'&&!opt.force){if(opt.auto)return false;if(!confirm('Há alterações locais de cadastro ainda não enviadas. Atualizar pelo SharePoint pode substituí-las. Continuar?'))return false;}const pr=objRows(await listRows(TABLE.packs),HEAD.packs),fr=objRows(await listRows(TABLE.funcs),HEAD.funcs);const old=(()=>{try{return JSON.parse(localStorage.getItem('ap_cadastro')||'{}')}catch{return {}}})();const p=pr.map(x=>({c:String(x.CODIGO||''),n:String(x.PACOTE_SERVICO||''),et:String(x.ETAPA||''),ob:String(x.OBRA_ESPECIFICA||''),l:'',t:'',st:String(x.STATUS||'Em andamento'),global:String(x.GLOBAL||'').toUpperCase()==='SIM'})).filter(x=>x.c&&x.n);const f=fr.map(x=>({m:String(x.MATRICULA||''),n:String(x.NOME||''),f:String(x.FUNCAO||''),r:String(x.REGIME||''),e:'',a:String(x.ATIVO||'Sim')})).filter(x=>x.m&&x.n);localStorage.setItem('ap_cadastro',JSON.stringify({f:f.length?f:(old.f||[]),p:p.length?p:(old.p||[])}));lset(K.catalogDirty,'0');lset(K.lastCatalog,String(Date.now()));window.dispatchEvent(new CustomEvent('elevatta-packages-updated'));try{RF.refreshPackages?.()}catch{}try{AP.reloadCatalog?.()}catch{}toast('Cadastros atualizados do SharePoint ✓');return {p:p.length,f:f.length}}
  M.pullCatalog=pullCatalog;
  function catalogBundle(){let d={};try{d=JSON.parse(localStorage.getItem('ap_cadastro')||'{}')}catch{}const upd=now(),packs=(d.p||[]).map(x=>[x.c||'',x.n||'',x.et||'',x.ob||'',x.st||'Em andamento',x.global?'SIM':'NAO',upd]).filter(r=>String(r[0]).trim()),funcs=(d.f||[]).map(x=>[x.m||'',x.n||'',x.f||'',x.r||'',x.a||'Sim',upd,employeeId(x.m)]).filter(r=>String(r[0]).trim());return {kind:'CATALOG',queueId:'CATALOG-MASTER',packs,funcs,updatedAt:upd}}
  async function pushCatalogBundle(b){await upsertNoDelete(TABLE.packs,HEAD.packs,b.packs||[]);await upsertNoDelete(TABLE.funcs,HEAD.funcs,b.funcs||[]);return true}
  async function queueCatalog(){const b=catalogBundle();await queue(b);if(navigator.onLine&&M.isConfigured())setTimeout(()=>flushOutbox({force:true,silent:true}).catch(()=>{}),0);return b.queueId}
  async function pushCatalog(opt={}){try{const id=await queueCatalog();if(!navigator.onLine||!M.isConfigured()){if(!opt.silent)toast('Cadastros protegidos offline · envio automático pendente');return false}await flushOutbox({force:true,silent:true});const pending=!!(await qGet(id));if(!opt.silent)toast(pending?'Cadastros protegidos · sincronização pendente':'Cadastros sincronizados com SharePoint ✓');return !pending}catch(e){if(!opt.silent)toast('Não foi possível proteger/sincronizar os cadastros');throw e}}
  M.pushCatalog=pushCatalog;
  let catalogQueueTimer=0;
  M.markCatalogDirty=()=>{lset(K.catalogDirty,'1');clearTimeout(catalogQueueTimer);catalogQueueTimer=setTimeout(()=>queueCatalog().catch(e=>console.warn('Fila de cadastro:',e)),900);};
  async function maybePushCatalog(){if(ls(K.catalogDirty,'0')!=='1')return false;try{return await pushCatalog({silent:true})}catch(e){console.warn('Cadastro pendente:',e);return false}}
  async function maybePullCatalog(){if(!navigator.onLine||!M.isConfigured()||ls(K.catalogDirty,'0')==='1')return;const last=Number(ls(K.lastCatalog,'0'))||0;if(Date.now()-last<15*60*1000)return;try{await pullCatalog({auto:true})}catch(e){console.warn('Catálogo M365:',e)}}

  // ---------- Query / local fallback / delete ----------
  let currentRecords=[];
  async function localFvsRecords(){const rows=await idbAll('records').catch(()=>[]);return rows.sort((a,b)=>String(b.updatedAt||b.createdAt).localeCompare(String(a.updatedAt||a.createdAt))).map(x=>{const r=x.bundle?.rec||{},s=x.bundle?.snapshot||{};return{id:String(r.id||x.id),date:String(r.date||''),obra:String(r.obra||''),service:String(r.serviceCode||s.service?.code||'')+' · '+String(r.serviceName||s.service?.name||''),type:r.type==='MEDICAO_FISICA'?'FVS + MEDIÇÃO FÍSICA':'SOMENTE FVS',status:x.syncedAt?'SINCRONIZADO':'LOCAL / PENDENTE',local:String(r.local||''),localOnly:!x.syncedAt}})}
  function localApRecords(){const L=[];try{for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i);if(!k?.startsWith('ap_dia_'))continue;const d=JSON.parse(localStorage.getItem(k)||'{}'),date=k.slice(7),works=new Set();(d.l||[]).forEach(x=>{if(x.obr)works.add(x.obr)});L.push({id:date,date,count:(d.l||[]).length,works:[...works],responsible:d.enc||'',localOnly:true})}}catch{}return L.sort((a,b)=>b.date.localeCompare(a.date))}
  async function loadRecords(kind='fvs'){
    const box=$('m365_records');if(box)box.innerHTML='<div class="m365-empty">Carregando registros…</div>';
    if(!navigator.onLine||!M.isConfigured()){
      currentRecords=kind==='ap'?localApRecords():await localFvsRecords();renderRecords();return currentRecords
    }
    try{
      if(kind==='ap'){
        const rows=objRows(await listRows(TABLE.ap),HEAD.ap),map=new Map(),pending=new Set((await qAll().catch(()=>[])).filter(x=>x.type==='AP_DAY').map(x=>String(x.bundle?.date||x.id?.replace(/^APDAY-/,''))));
        rows.forEach(x=>{const d=String(x.DATA||'');if(!d)return;const o=map.get(d)||{id:d,date:d,count:0,works:new Set(),responsible:x.RESPONSAVEL||'',localOnly:false};o.count++;if(x.OBRA)o.works.add(x.OBRA);map.set(d,o)});
        for(const l of localApRecords()){if(pending.has(l.date)||!map.has(l.date))map.set(l.date,Object.assign({},l,{localOnly:pending.has(l.date)||!map.has(l.date)}))}
        currentRecords=[...map.values()].sort((a,b)=>b.date.localeCompare(a.date)).map(x=>({...x,works:Array.isArray(x.works)?x.works:[...(x.works||[])]}));
      }else{
        const rows=objRows(await listRows(TABLE.fvs),HEAD.fvs),map=new Map(rows.map(x=>[String(x.ID),{id:String(x.ID),date:String(x.DATA||''),obra:String(x.OBRA||''),service:String(x.PACOTE_COD||'')+' · '+String(x.PACOTE_SERVICO||''),type:String(x.TIPO||''),status:String(x.SITUACAO||''),local:String(x.LOCAL||''),localOnly:false}]));
        for(const l of await localFvsRecords()){if(l.localOnly||!map.has(l.id))map.set(l.id,l)}currentRecords=[...map.values()].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
      }
      renderRecords();return currentRecords
    }catch(e){console.warn('Histórico remoto:',e);currentRecords=kind==='ap'?localApRecords():await localFvsRecords();renderRecords();return currentRecords}
  }
  M.loadRecords=loadRecords;
  function renderRecords(){const box=$('m365_records');if(!box)return;const q=String($('m365_search')?.value||'').toLowerCase(),kind=$('m365_kind')?.value||'fvs';const L=currentRecords.filter(x=>!q||JSON.stringify(x).toLowerCase().includes(q)).slice(0,120);box.innerHTML=L.length?L.map(x=>kind==='ap'?`<div class="m365-record"><div><b>${esc(x.date)} · ${x.count} lançamento${x.count===1?'':'s'}</b><small>${esc((x.works||[]).join(' · ')||'Sem obra')} · ${esc(x.responsible||'')}${x.localOnly?' · local':''}</small></div><div class="m365-ra"><button onclick="ElevattaM365.openApDay('${esc(x.date)}')">Abrir</button>${(x.localOnly||!navigator.onLine)?'':`<button class="danger" onclick="ElevattaM365.deleteApDay('${esc(x.date)}')">Excluir</button>`}</div></div>`:`<div class="m365-record"><div><b>${esc(x.date)} · ${esc(x.service)}</b><small>${esc(x.obra)}${x.local?' · '+esc(x.local):''} · ${esc(x.type)} · ${esc(x.status)}</small></div><div class="m365-ra">${(x.localOnly||!navigator.onLine)?'':`<button class="danger" onclick="ElevattaM365.deleteFvs('${esc(x.id)}')">Excluir</button>`}</div></div>`).join(''):'<div class="m365-empty">Nenhum registro encontrado.</div>'}
  M.filterRecords=renderRecords;
  M.changeKind=()=>loadRecords($('m365_kind')?.value||'fvs');
  M.openApDay=async date=>{close();const input=$('a_data');if(input)input.value=date;if(navigator.onLine)await M.pullApontamentoDate(date,{silent:true});window.setMode?.('ap');window.AP?.trocarDia?.()};
  M.deleteApDay=async date=>{if(!confirm('Excluir do SharePoint todos os apontamentos de '+date+'?'))return;try{const rows=objRows(await listRows(TABLE.ap),HEAD.ap);await deleteIndexes(TABLE.ap,rows.filter(x=>String(x.DATA)===String(date)).map(x=>x.__index));toast('Apontamento excluído do SharePoint');await loadRecords('ap')}catch(e){alert(e.message||e)}};
  async function removeLocalFvs(id){const r=await idbGet('records',id).catch(()=>null);for(const key of (r?.photoKeys||[]))await idbDel('photos',key).catch(()=>{});await idbDel('records',id).catch(()=>{});await idbDel('outbox','FVS-'+id).catch(()=>{})}
  M.deleteFvs=async id=>{if(!navigator.onLine){toast('Sem internet: exclusão do SharePoint indisponível');return}if(!confirm('Excluir este registro FVS/Medição e suas evidências do SharePoint?'))return;try{for(const [t,h,key] of [[TABLE.items,HEAD.items,'REGISTRO_ID'],[TABLE.med,HEAD.med,'REGISTRO_ID'],[TABLE.photos,HEAD.photos,'REGISTRO_ID'],[TABLE.fvs,HEAD.fvs,'ID']]){const rows=objRows(await listRows(t),h);await deleteIndexes(t,rows.filter(x=>String(x[key])===String(id)).map(x=>x.__index))}await deletePhotoFolder(id);await removeLocalFvs(String(id));toast('Registro excluído do SharePoint');await loadRecords('fvs')}catch(e){alert(e.message||e)}};

  // ---------- UI ----------
  function inject(){if($('m365box'))return;const style=document.createElement('style');style.id='m365-style';style.textContent=`
#m365box{position:fixed;inset:0;z-index:1200;display:none;align-items:flex-end;justify-content:center;background:rgba(5,15,30,.73);backdrop-filter:blur(8px)}#m365box.show{display:flex}.m365-sheet{width:min(780px,100%);max-height:94dvh;background:#f4f7fb;border-radius:24px 24px 0 0;overflow:hidden;display:flex;flex-direction:column;color:#152c49;box-shadow:0 -18px 65px rgba(2,13,30,.3)}.m365-head{display:flex;gap:12px;align-items:center;padding:15px 16px;background:#10233f;color:#fff}.m365-logo{width:42px;height:42px;border-radius:12px;background:#2467c7;display:grid;place-items:center;font-weight:900;font-size:18px}.m365-head .copy{min-width:0;flex:1}.m365-head b{display:block;font-size:15px}.m365-head small{display:block;color:#a9bdd9;font-size:10px;margin-top:2px}.m365-x{border:0;width:38px;height:38px;border-radius:11px;background:rgba(255,255,255,.1);color:#fff;font-size:23px}.m365-status{display:flex;align-items:center;gap:9px;padding:9px 14px;background:#fff;border-bottom:1px solid #dde5ef;font-size:10px}.m365-status i{width:8px;height:8px;border-radius:50%;background:#9aa7b8}.m365-status.ok i{background:#17865f}.m365-status.warn i{background:#c4850c}.m365-tabs{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;padding:8px 10px;background:#fff;border-bottom:1px solid #e0e6ef}.m365-tabs button{border:0;border-radius:10px;padding:10px 5px;background:#edf2f8;color:#61718a;font:800 10px Arial}.m365-tabs button.on{background:#17365f;color:#fff}.m365-body{overflow:auto;padding:12px;-webkit-overflow-scrolling:touch}.m365-panel{display:none}.m365-panel.on{display:block}.m365-card{background:#fff;border:1px solid #dfe6ef;border-radius:16px;padding:13px;margin-bottom:10px;box-shadow:0 4px 15px rgba(24,53,88,.04)}.m365-card h4{margin:0 0 5px;font-size:13px}.m365-card p{margin:0 0 10px;font-size:10px;color:#74839a;line-height:1.45}.m365-card label{font:800 9px Arial;color:#52647e;display:block;margin:9px 0 4px;text-transform:uppercase;letter-spacing:.04em}.m365-card input,.m365-card select{box-sizing:border-box;width:100%;border:1px solid #d4deea;background:#fbfcfe;border-radius:11px;padding:11px;font:11px Arial;color:#172d4d}.m365-card input[readonly]{background:#eef3f8;color:#6a7a91}.m365-grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px}.m365-actions{display:flex;gap:7px;flex-wrap:wrap;margin-top:11px}.m365-actions button,.m365-ra button{border:0;border-radius:10px;padding:9px 12px;background:#17365f;color:#fff;font:800 10px Arial}.m365-actions .primary{background:#ff3f1a}.m365-actions .green{background:#16825d}.m365-actions .soft,.m365-ra button{background:#eaf0f8;color:#23476f}.m365-msg{margin-top:9px;border-radius:9px;padding:8px 9px;background:#eef3f9;color:#586b84;font-size:9.5px;min-height:14px}.m365-msg.ok{background:#e8f7f1;color:#176d50}.m365-msg.err{background:#fff0ee;color:#aa3028}.m365-search{display:grid;grid-template-columns:120px 1fr auto;gap:7px;margin-bottom:9px}.m365-search button{border:0;border-radius:10px;background:#17365f;color:#fff;font-weight:850}.m365-record{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;align-items:center;padding:10px 11px;background:#fff;border:1px solid #dfe6ef;border-radius:12px;margin-bottom:7px}.m365-record b{font-size:10.5px;display:block}.m365-record small{font-size:9px;color:#718099;display:block;margin-top:3px;line-height:1.35}.m365-ra{display:flex;gap:5px}.m365-ra .danger{background:#fff0ee;color:#b2342b}.m365-empty{text-align:center;padding:22px 9px;color:#718099;font-size:10.5px}.m365-empty.error{color:#b2342b}.m365-note{font-size:9px;color:#72819a;line-height:1.45;margin-top:9px}.m365-mini{position:fixed;right:12px;bottom:12px;z-index:850;border:0;border-radius:999px;background:#10233f;color:#fff;padding:9px 12px;font:850 9px Arial;box-shadow:0 8px 24px rgba(7,24,49,.25);display:none}body[data-modo="ap"] .m365-mini,body[data-modo="rf"] .m365-mini{display:block}.m365-mini i{display:inline-block;width:7px;height:7px;border-radius:50%;background:#c4850c;margin-right:5px}.m365-mini.ok i{background:#24ad7d}@media(min-width:700px){#m365box{align-items:center;padding:20px}.m365-sheet{border-radius:24px;max-height:88vh}}@media(max-width:520px){.m365-grid2{grid-template-columns:1fr}.m365-search{grid-template-columns:1fr 1fr}.m365-search input{grid-column:1/-1}.m365-ra{flex-direction:column}}`;
    document.head.appendChild(style);
    const div=document.createElement('div');div.id='m365box';div.innerHTML=`<div class="m365-sheet"><div class="m365-head"><div class="m365-logo">M</div><div class="copy"><b>Base Microsoft 365</b><small>Offline-first · SharePoint · sincronização automática</small></div><button class="m365-x" onclick="ElevattaM365.close()">×</button></div><div class="m365-status" id="m365_status"><i></i><span id="m365_status_text">Preparando armazenamento local…</span></div><div class="m365-tabs"><button id="m365_tab_records" onclick="ElevattaM365.tab('records')">Registros</button><button id="m365_tab_catalog" onclick="ElevattaM365.tab('catalog')">Cadastros</button><button id="m365_tab_connect" onclick="ElevattaM365.tab('connect')">Conexão</button></div><div class="m365-body"><section class="m365-panel" id="m365_panel_records"><div class="m365-card"><h4>Histórico local + compartilhado</h4><p>Sem internet, o aparelho mostra a cópia local. Com internet, pendências são enviadas automaticamente ao Excel do SharePoint.</p><div class="m365-search"><select id="m365_kind" onchange="ElevattaM365.changeKind()"><option value="fvs">FVS / Medições</option><option value="ap">Apontamentos</option></select><input id="m365_search" placeholder="Buscar obra, pacote, data…" oninput="ElevattaM365.filterRecords()"><button onclick="ElevattaM365.changeKind()">Atualizar</button></div><div id="m365_records"></div></div></section><section class="m365-panel" id="m365_panel_catalog"><div class="m365-card"><h4>Pacotes e funcionários</h4><p>A cópia mais recente fica no aparelho para uso offline. Quando houver conexão, você pode atualizar os cadastros do SharePoint.</p><div class="m365-actions"><button class="green" onclick="ElevattaM365.pullCatalog()">↓ Atualizar este aparelho</button><button class="soft" onclick="ElevattaM365.pushCatalog()">↑ Enviar cadastro local</button></div></div></section><section class="m365-panel" id="m365_panel_connect"><div class="m365-card"><h4>Conexão corporativa</h4><p>O login Microsoft é feito uma vez. O refresh token fica protegido em cookie HttpOnly no servidor; o PWA não armazena client secret.</p><div class="m365-grid2"><div><label>Tenant / Directory ID</label><input id="m365_tenant" readonly></div><div><label>Application / Client ID</label><input id="m365_client" readonly></div></div><label>Excel compartilhado no SharePoint</label><input id="m365_workbook" readonly placeholder="Configure M365_WORKBOOK_URL no Render"><div class="m365-actions"><button class="primary" onclick="ElevattaM365.saveAndLogin()">Entrar com Microsoft</button><button class="green" onclick="ElevattaM365.test()">Testar base</button><button class="soft" onclick="ElevattaM365.logout()">Sair</button></div><div class="m365-msg" id="m365_msg"></div><div class="m365-note" id="m365_storage"></div><div class="m365-note">No Entra, use aplicação Web/confidential com Redirect URI <b>/auth/callback</b> e permissões delegadas <b>User.Read</b> + <b>Files.ReadWrite</b>. O servidor usa o client secret apenas no Render.</div></div></section></div></div>`;document.body.appendChild(div);
    const mini=document.createElement('button');mini.className='m365-mini';mini.id='m365_mini';mini.innerHTML='<i></i><span>Offline</span>';mini.onclick=()=>open('records');document.body.appendChild(mini)
  }
  function setMsg(t,ok){const e=$('m365_msg');if(!e)return;e.textContent=t||'';e.classList.toggle('ok',ok===true);e.classList.toggle('err',ok===false)}
  function fillForm(){const c=cfg();if($('m365_tenant'))$('m365_tenant').value=c.tenant;if($('m365_client'))$('m365_client').value=c.client;if($('m365_workbook'))$('m365_workbook').value=c.workbook}
  M.saveAndLogin=async()=>{await loadServerConfig(true);fillForm();await login()};
  async function updateStatus(){
    inject();const pending=await M.pendingCount(),a=account(),configured=M.isConfigured(),online=navigator.onLine,logged=!!tokenInfo().access;
    const st=$('m365_status'),txt=$('m365_status_text'),mini=$('m365_mini'),badge=$('m365_home_badge');
    if(st)st.className='m365-status '+(online&&configured&&logged&&!pending?'ok':(pending||configured?'warn':''));
    let text='Armazenamento local ativo';
    if(!online)text=pending?`Sem internet · ${pending} pendência${pending>1?'s':''} protegida${pending>1?'s':''} no aparelho`:'Sem internet · trabalhando no aparelho';
    else if(!serverCfg.authConfigured&&!lastConfigError)text='Microsoft 365 ainda não configurado no Render';
    else if(syncState.busy&&pending)text=`Sincronizando ${pending} pendência${pending>1?'s':''}…`;
    else if(syncState.lastError&&pending)text=`Internet/SharePoint instável · ${pending} pendência${pending>1?'s':''} protegida${pending>1?'s':''}`;
    else if(pending)text=`Online · ${pending} pendência${pending>1?'s':''} aguardando sincronização`;
    else if(logged)text=a.email||a.name||'Microsoft 365 conectado';
    else if(lastConfigError)text='Online, mas o servidor do Elevatta não respondeu';
    else text='Online · faça o primeiro login Microsoft';
    if(txt)txt.textContent=text;
    if(mini){mini.classList.toggle('ok',online&&configured&&logged&&!pending);const sp=mini.querySelector('span');if(sp)sp.textContent=!online?'Offline':pending?(pending+' pend.'):(logged?'Sincronizado':'M365')}
    if(badge)badge.textContent=!online?'offline':pending?(pending+' pend.'):(logged?'conectado':configured?'entrar':'configurar');
    const su=$('m365_storage');if(su){const used=syncState.usage?Math.round(syncState.usage/1048576):0,quota=syncState.quota?Math.round(syncState.quota/1048576):0,pct=syncState.quota?Math.round(syncState.usage/syncState.quota*100):0;su.textContent='Armazenamento offline: '+(used?used+' MB'+(quota?' de '+quota+' MB ('+pct+'%)':''):'ativo')+' · '+(pct>=85?'ATENÇÃO: pouco espaço disponível':syncState.storagePersistent===true?'protegido contra limpeza automática':syncState.storagePersistent===false?'o navegador pode liberar dados se faltar espaço':'proteção do armazenamento não informada');}
  }
  function tab(t){['records','catalog','connect'].forEach(x=>{$('m365_panel_'+x)?.classList.toggle('on',x===t);$('m365_tab_'+x)?.classList.toggle('on',x===t)});if(t==='records')loadRecords($('m365_kind')?.value||'fvs');if(t==='connect')fillForm()}
  M.tab=tab;
  function open(t='records'){inject();$('m365box').classList.add('show');fillForm();updateStatus();tab(t);if(!serverCfg.authConfigured||!cfg().workbook)tab('connect')}
  function close(){$('m365box')?.classList.remove('show')}
  M.open=open;M.close=close;

  async function restoreServerSession(){
    if(!navigator.onLine)return false;
    await loadServerConfig(true);
    if(!serverCfg.authConfigured)return false;
    try{await getToken(false);if(!account().email&&!account().name)await loadProfile();return true}catch{return false}
  }
  async function onConnectivity(force=false){
    if(navigator.onLine){await restoreServerSession().catch(()=>{});fillForm()}
    await updateStatus();
    if(navigator.onLine){await migrateLegacyOutbox().catch(()=>{});await maybePushCatalog().catch(()=>{});await flushOutbox({force:!!force}).catch(()=>{});await maybePullCatalog().catch(()=>{});await updateStatus()}
  }
  window.addEventListener('online',()=>onConnectivity(true));
  window.addEventListener('offline',()=>{syncState.busy=false;updateStatus()});
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')onConnectivity(false)});
  navigator.serviceWorker?.addEventListener?.('message',e=>{if(e.data?.type==='ELEVATTA_SYNC_NOW')onConnectivity(false);else if(e.data?.type==='ELEVATTA_SYNC_DONE')updateStatus()});
  setInterval(()=>{if(navigator.onLine)flushOutbox({silent:true}).catch(()=>{})},60000);

  async function boot(){
    inject();await loadServerConfig();fillForm();await persistStorage();await migrateLegacyOutbox().catch(()=>{});await handleRedirect();
    if(navigator.onLine)await restoreServerSession().catch(()=>{});
    fillForm();await updateStatus();
    if(navigator.onLine){setTimeout(()=>maybePushCatalog().catch(()=>{}),350);setTimeout(()=>flushOutbox({force:true,silent:true}).catch(()=>{}),700);setTimeout(()=>maybePullCatalog().catch(()=>{}),1800)}
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
  M._version=VERSION;
  M._db=db;
  return M;
})();
