/* Elevatta ERP REV40 — shell offline-first + sincronização resiliente Microsoft 365. */
const CACHE='elevatta-erp-v40-r3';
const SHELL=['/','/index.html','/erp.html','/erp.css','/erp.js','/erp-core.js','/validator.js','/erp-schema.json','/consolidacao_rev38_legacy.html','/rev32.js','/m365.js','/rev33.js','/exceljs.min.js','/jspdf.umd.min.js','/manifest.webmanifest','/icon-192.png','/icon-512.png'];
const TABLE={ap:'tbApontamentos',fvs:'tbFvsRegistros',items:'tbFvsItens',med:'tbMedicoes',photos:'tbFotos',packs:'tbPacotes',funcs:'tbFuncionarios'};
const HEAD={ap:["ID","DATA","RESPONSAVEL","MATRICULA","NOME","FUNCAO","REGIME","OBRA","ETAPA","PACOTE_COD","PACOTE_SERVICO","PAVIMENTO","EQUIPE","FRACAO_DIA","CLASSIFICACAO","SITUACAO","OBSERVACAO","MOTIVO_FALTA","SAIU_CEDO","UPDATED_AT","FRENTE_ID","LOCAL_FRENTE","DEVICE_ID","SYNC_VERSION"],fvs:["ID","DATA","OBRA","RESPONSAVEL","PACOTE_COD","PACOTE_SERVICO","ETAPA","TIPO","LOCAL","SITUACAO","CONFORMES","NAO_CONFORMES","NA","PENDENTES","QTD_FOTOS","OBSERVACAO","UPDATED_AT","FRENTE_ID","PAVIMENTO","DEVICE_ID","SYNC_VERSION"],items:["ID","REGISTRO_ID","ITEM","VERIFICACAO","CRITERIO","STATUS","OBSERVACAO","QTD_FOTOS","UPDATED_AT","FRENTE_ID"],med:["ID","REGISTRO_ID","DATA","OBRA","PACOTE_COD","PACOTE_SERVICO","LOCAL","ANTERIOR_PCT","EXECUTADO_PCT","APROVADO_PCT","MEDIDO_PCT","ACUMULADO_PCT","SALDO_PCT","PENDENTE_QUALIDADE_PCT","APROVADO_NAO_MEDIDO_PCT","OBSERVACAO","UPDATED_AT","FRENTE_ID","PAVIMENTO"],photos:["ID","REGISTRO_ID","DATA","OBRA","PACOTE_COD","PACOTE_SERVICO","ITEM","STATUS","LEGENDA","MARCACOES_QTD","ARQUIVO_URL","ARQUIVO_NOME","UPDATED_AT","FRENTE_ID","ENTIDADE_TIPO"],packs:["CODIGO","PACOTE_SERVICO","ETAPA","OBRA_ESPECIFICA","STATUS","GLOBAL","UPDATED_AT"],funcs:["MATRICULA","NOME","FUNCAO","REGIME","ATIVO","UPDATED_AT","FUNCIONARIO_ID"]};

self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE).then(async cache=>{
    /* O shell só é considerado pronto quando todos os arquivos essenciais
       entram no cache. Evita instalar um PWA "meio offline". */
    await cache.addAll(SHELL);
    await self.skipWaiting();
  }));
});

self.addEventListener('activate',event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});

async function fetchTimed(req,init={},ms=10000){const ctl=new AbortController(),timer=setTimeout(()=>ctl.abort(),ms);try{return await fetch(req,Object.assign({},init,{signal:ctl.signal}))}finally{clearTimeout(timer)}}
async function networkAndUpdate(req,ms=10000){
  try{
    const resp=await fetchTimed(req,{},ms);
    if(resp&&resp.ok&&resp.type==='basic'){
      const cache=await caches.open(CACHE);
      cache.put(req,resp.clone()).catch(()=>{});
    }
    return resp;
  }catch{return null;}
}

self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin!==location.origin)return;
  if(url.pathname.startsWith('/api/')||url.pathname.startsWith('/auth/')||url.pathname==='/healthz')return;

  if(req.mode==='navigate'){
    event.respondWith((async()=>{
      /* Campo primeiro: se o shell já existe, abre imediatamente pelo cache e
         atualiza o index em paralelo. O sw.js continua sendo verificado pelo
         registro da página; uma nova revisão assume o controle e recarrega. */
      const target=(url.pathname==='/erp.html'||url.pathname.startsWith('/erp/'))?'/erp.html':'/index.html';
      const cached=await caches.match(target);
      if(cached){event.waitUntil(networkAndUpdate(new Request(target,{cache:'no-store'}),6500).catch(()=>null));return cached;}
      const net=await networkAndUpdate(req,6500);
      return net||new Response('<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Elevatta</title><body style="font-family:Arial;padding:24px;background:#10233f;color:#fff"><h2>Elevatta</h2><p>Este aparelho ainda não concluiu o primeiro cache do Elevatta. Conecte uma vez, abra o aplicativo e aguarde a instalação terminar.</p></body>',{headers:{'content-type':'text/html; charset=utf-8'}});
    })());
    return;
  }

  event.respondWith((async()=>{
    const cached=await caches.match(req);
    if(cached){networkAndUpdate(req).catch(()=>{});return cached;}
    return (await networkAndUpdate(req))||Response.error();
  })());
});

// ---------- IndexedDB compartilhado com m365.js ----------
let dbp=null;
function db(){
  if(dbp)return dbp;
  dbp=new Promise((res,rej)=>{
    const q=indexedDB.open('elevatta-m365',2);
    q.onupgradeneeded=()=>{
      const d=q.result;
      if(!d.objectStoreNames.contains('outbox'))d.createObjectStore('outbox',{keyPath:'id'});
      if(!d.objectStoreNames.contains('records'))d.createObjectStore('records',{keyPath:'id'});
      if(!d.objectStoreNames.contains('photos'))d.createObjectStore('photos',{keyPath:'key'});
      if(!d.objectStoreNames.contains('apArchive'))d.createObjectStore('apArchive',{keyPath:'date'});
      if(!d.objectStoreNames.contains('meta'))d.createObjectStore('meta',{keyPath:'key'});
    };
    q.onsuccess=()=>{const d=q.result;d.onversionchange=()=>{try{d.close()}catch{}dbp=null};res(d)};q.onerror=()=>{dbp=null;rej(q.error)};
  });
  return dbp;
}
async function idbGet(store,key){const d=await db();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).get(key);q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}
async function idbAll(store){const d=await db();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).getAll();q.onsuccess=()=>res(q.result||[]);q.onerror=()=>rej(q.error)})}
async function idbPut(store,val){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).put(val);tx.oncomplete=()=>res(val);tx.onerror=()=>rej(tx.error)})}
async function idbDel(store,key){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).delete(key);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
async function acquireSyncLock(owner,ttl=120000){const d=await db();return new Promise((res,rej)=>{let acquired=false;const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('sync-lock');q.onsuccess=()=>{const cur=q.result;if(!cur||Number(cur.expiresAt||0)<Date.now()||cur.owner===owner){acquired=true;os.put({key:'sync-lock',owner,expiresAt:Date.now()+ttl,updatedAt:new Date().toISOString()})}};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res(acquired);tx.onerror=()=>rej(tx.error)})}
async function releaseSyncLock(owner){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('sync-lock');q.onsuccess=()=>{if(q.result?.owner===owner)os.delete('sync-lock')};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}

// ---------- Microsoft Graph ----------
function b64urlBytes(bytes){let bin='';bytes.forEach(b=>bin+=String.fromCharCode(b));return btoa(bin).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'')}
function encodeShareUrl(url){return 'u!'+b64urlBytes(new TextEncoder().encode(url))}
function sanitize(s){return String(s||'arquivo').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9._-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,100)||'arquivo'}
function upper(s){return String(s??'').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')}
function fnv1a(str,seed=0x811c9dc5){let h=seed>>>0;for(let i=0;i<str.length;i++){h^=str.charCodeAt(i);h=Math.imul(h,0x01000193)>>>0}return h>>>0}
function hash12(str){const z=String(str);return fnv1a(z,0x811c9dc5).toString(16).padStart(8,'0')+fnv1a(z,0x9e3779b9).toString(16).padStart(8,'0').slice(0,4)}
function frontId(x={}){const o=upper(x.obra),p=upper(x.pacote||x.pacoteCod),pv=upper(x.pavimento),lf=upper(x.local||x.localFrente);if(!o||!p||(!pv&&!lf))return '';const k=[o,p,pv,lf].join('|');return 'FRT-'+hash12(k).toUpperCase()}
function statusLabel(x){return x==='C'?'CONFORME':x==='NC'?'NÃO CONFORME':x==='NA'?'N/A':'PENDENTE'}
function retryDelay(attempts){return Math.min(15*60*1000,5000*Math.pow(2,Math.min(7,Math.max(0,attempts-1))))}

async function config(){
  const r=await fetchTimed('/api/m365/config',{cache:'no-store',credentials:'include'},10000);
  if(!r.ok)throw new Error('Configuração M365 indisponível.');
  return r.json();
}
async function accessToken(){
  const r=await fetchTimed('/api/m365/token',{cache:'no-store',credentials:'include'},12000);
  const j=await r.json().catch(()=>({}));
  if(!r.ok||!j.access_token){const e=new Error(j.error||'Sessão Microsoft indisponível.');e.reauthRequired=!!j.reauthRequired;throw e}
  return j.access_token;
}
async function graph(path,token,opt={}){
  const init={method:opt.method||'GET',headers:Object.assign({Authorization:'Bearer '+token},opt.headers||{})};
  if(opt.body!==undefined){if(opt.raw)init.body=opt.body;else{init.headers['Content-Type']='application/json';init.body=JSON.stringify(opt.body)}}
  const url=path.startsWith('http')?path:'https://graph.microsoft.com/v1.0'+path;let r=null;
  for(let attempt=0;attempt<3;attempt++){
    r=await fetchTimed(url,init,opt.raw?45000:22000);
    if(![429,502,503,504].includes(r.status)||attempt===2)break;
    const ra=Number(r.headers.get('retry-after')),delay=Number.isFinite(ra)&&ra>0?Math.min(15000,ra*1000):Math.min(5000,600*Math.pow(2,attempt));await new Promise(x=>setTimeout(x,delay));
  }
  if(r.status===204)return null;
  const ct=r.headers.get('content-type')||'';const j=ct.includes('json')?await r.json().catch(()=>({})):await r.blob();
  if(!r.ok)throw new Error(j?.error?.message||j?.error_description||('Microsoft Graph: '+r.status));
  return j;
}
async function workbookLocator(workbookUrl,token){
  const item=await graph('/shares/'+encodeURIComponent(encodeShareUrl(workbookUrl))+'/driveItem?$select=id,name,webUrl,parentReference',token);
  const x={driveId:item.parentReference?.driveId,itemId:item.id,parentId:item.parentReference?.id};
  if(!x.driveId||!x.itemId)throw new Error('Planilha do SharePoint não localizada.');return x;
}
function tablePath(l,table,suffix=''){return `/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(l.itemId)}/workbook/tables/${encodeURIComponent(table)}${suffix}`}
async function listRows(l,table,token){let next=tablePath(l,table,'/rows?$top=1000'),out=[],pages=0;while(next&&pages<100){const j=await graph(next,token);out.push(...((j?.value||[]).map(r=>({index:r.index,values:(r.values&&r.values[0])||[]}))));next=j?.['@odata.nextLink']||'';pages++}return out}
function objRows(rows,headers){return rows.map(r=>{const o={__index:r.index};headers.forEach((h,i)=>o[h]=r.values[i]);return o}).filter(o=>String(o[headers[0]]??'').trim())}
async function deleteIndexes(l,table,indexes,token){for(const i of [...indexes].sort((a,b)=>b-a))await graph(tablePath(l,table,'/rows/'+i),token,{method:'DELETE'})}
async function appendRows(l,table,rows,token){if(rows?.length)await graph(tablePath(l,table,'/rows'),token,{method:'POST',body:{values:rows}})}
async function updateRow(l,table,index,row,token){await graph(tablePath(l,table,'/rows/'+index),token,{method:'PATCH',body:{values:[row]}})}
async function replaceWhere(l,table,headers,predicate,newRows,token){const scoped=objRows(await listRows(l,table,token),headers).filter(predicate),key=headers[0],byId=new Map();scoped.forEach(x=>{const id=String(x[key]??'');if(!byId.has(id))byId.set(id,[]);byId.get(id).push(x)});const keep=new Set();for(const row of (newRows||[])){const id=String(row?.[0]??'');keep.add(id);const group=byId.get(id)||[];if(group.length)await updateRow(l,table,group[0].__index,row,token);else await appendRows(l,table,[row],token)}const stale=[];for(const x of scoped){const id=String(x[key]??''),group=byId.get(id)||[];if(!keep.has(id)||group.indexOf(x)>0)stale.push(x.__index)}if(stale.length)await deleteIndexes(l,table,stale,token)}
async function upsertNoDelete(l,table,headers,newRows,token){const all=objRows(await listRows(l,table,token),headers),key=headers[0],byId=new Map();all.forEach(x=>{const id=String(x[key]??'');if(id&&!byId.has(id))byId.set(id,x)});for(const row of (newRows||[])){const id=String(row?.[0]??'');if(!id)continue;const old=byId.get(id);if(old)await updateRow(l,table,old.__index,row,token);else await appendRows(l,table,[row],token)}}
async function children(l,parentId,token){const j=await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(parentId)}/children?$select=id,name,folder,webUrl&$top=200`,token);return j.value||[]}
async function ensureFolder(l,parentId,name,token){const list=await children(l,parentId,token),found=list.find(x=>x.folder&&String(x.name).toLowerCase()===String(name).toLowerCase());if(found)return found;return graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(parentId)}/children`,token,{method:'POST',body:{name,folder:{},'@microsoft.graph.conflictBehavior':'rename'}})}
async function photoFolder(l,recordId,token){const root=await ensureFolder(l,l.parentId,'ELEVATTA_FOTOS',token);return ensureFolder(l,root.id,sanitize(recordId),token)}
async function uploadBlob(l,recordId,p,i,blob,token,folder){const dest=folder||await photoFolder(l,recordId,token),name=String(i+1).padStart(2,'0')+'_'+sanitize(p.id||('foto_'+i))+'.jpg';return graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(dest.id)}:/${encodeURIComponent(name)}:/content`,token,{method:'PUT',body:blob,raw:true,headers:{'Content-Type':blob.type||'image/jpeg'}})}
async function cleanupPhotoFolder(l,folder,expected,token){if(!folder)return;const keep=new Set(expected);for(const x of await children(l,folder.id,token)){if(!x.folder&&!keep.has(String(x.name||'')))await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(x.id)}`,token,{method:'DELETE'})}}
async function deletePhotoFolder(l,recordId,token){try{const base=(await children(l,l.parentId,token)).find(x=>x.folder&&String(x.name||'').toLowerCase()==='elevatta_fotos');if(!base)return;const rec=(await children(l,base.id,token)).find(x=>x.folder&&String(x.name||'').toLowerCase()===sanitize(recordId).toLowerCase());if(rec)await graph(`/drives/${encodeURIComponent(l.driveId)}/items/${encodeURIComponent(rec.id)}`,token,{method:'DELETE'})}catch(e){console.warn('photo folder delete',e)}}

async function pushAp(l,b,token){const current=(b.rows||[]),currentIds=new Set(current.map(r=>String(r?.[0]??''))),owned=new Set([...(b.ownedIds||[]).map(String),...currentIds]),scoped=objRows(await listRows(l,TABLE.ap,token),HEAD.ap).filter(o=>String(o.DATA)===String(b.date)&&owned.has(String(o.ID||''))),byId=new Map();scoped.forEach(x=>{const id=String(x.ID||'');if(!byId.has(id))byId.set(id,[]);byId.get(id).push(x)});for(const row of current){const id=String(row?.[0]??''),group=byId.get(id)||[];if(group.length)await updateRow(l,TABLE.ap,group[0].__index,row,token);else await appendRows(l,TABLE.ap,[row],token)}const stale=[];for(const x of scoped){const id=String(x.ID||''),group=byId.get(id)||[];if(!currentIds.has(id)||group.indexOf(x)>0)stale.push(x.__index)}if(stale.length)await deleteIndexes(l,TABLE.ap,stale,token)}
async function pushFvs(l,stored,token){
  const b=stored.bundle||{},r=b.rec||{},s=b.snapshot||{},q=s.quality||{},upd=new Date().toISOString(),sit=Number(q.NC||0)>0?'NÃO CONFORME':Number(q.P||0)>0?'PENDENTE':'APROVADA',code=r.serviceCode||s.service?.code||'',name=r.serviceName||s.service?.name||'',stage=r.serviceStage||s.service?.etapa||'',fid=frontId({obra:r.obra,pacote:code,pavimento:r.pavimento,local:r.local}),dev=r.deviceId||'';
  await replaceWhere(l,TABLE.fvs,HEAD.fvs,o=>String(o.ID)===String(r.id),[[r.id,r.date||'',r.obra||'',r.responsavel||'',code,name,stage,r.type==='MEDICAO_FISICA'?'FVS + MEDIÇÃO FÍSICA':'SOMENTE FVS',r.local||'',sit,Number(q.C||0),Number(q.NC||0),Number(q.NA||0),Number(q.P||0),Number(r.photoCount||b.photos?.length||0),r.note||'',upd,fid,r.pavimento||'',dev,'40']],token);
  const ir=(b.items||[]).map((x,i)=>[r.id+'-I-'+String(x.n||i+1),r.id,Number(x.n||i+1),x.desc||'',x.criterion||'',statusLabel(x.status),x.obs||'',Number(x.photos||0),upd,fid]);
  await replaceWhere(l,TABLE.items,HEAD.items,o=>String(o.REGISTRO_ID)===String(r.id),ir,token);
  if(r.type==='MEDICAO_FISICA'){
    const mr=[[r.id,r.id,r.date||'',r.obra||'',code,name,r.local||'',Number(r.previous||0)/100,Number(r.executed||0)/100,Number(r.approved||0)/100,Number(r.measured||0)/100,Number(r.accumulated||0)/100,Number(r.balance||0)/100,Number(r.pendingQuality||0)/100,Number(r.approvedUnmeasured||0)/100,r.note||'',upd,fid,r.pavimento||'']];
    await replaceWhere(l,TABLE.med,HEAD.med,o=>String(o.ID)===String(r.id),mr,token);
  }
  const photoRows=[],expected=[],folder=(b.photos||[]).length?await photoFolder(l,r.id,token):null;
  for(let i=0;i<(b.photos||[]).length;i++){
    const p=b.photos[i],local=await idbGet('photos',p.photoKey);const blob=local?.uploadBlob||local?.blob;
    if(!blob)throw new Error('Foto local ausente: '+String(p.id||i));
    const file=await uploadBlob(l,r.id,p,i,blob,token,folder);if(!file?.id)throw new Error('Upload da foto '+(i+1)+' não confirmado.');const appUrl='/api/erp/file?itemId='+encodeURIComponent(file.id)+'&name='+encodeURIComponent(file.name||('foto_'+i+'.jpg'));
    photoRows.push([p.id||r.id+'-P-'+i,r.id,r.date||'',r.obra||'',code,name,p.g||'','',p.cap||'',(p.marks||[]).length,appUrl,file.name||'',upd,fid,'FVS']);expected.push(String(file.name||''));
  }
  await replaceWhere(l,TABLE.photos,HEAD.photos,o=>String(o.REGISTRO_ID)===String(r.id),photoRows,token);if(folder)await cleanupPhotoFolder(l,folder,expected,token);else await deletePhotoFolder(l,r.id,token);
}

async function pushCatalog(l,b,token){await upsertNoDelete(l,TABLE.packs,HEAD.packs,b.packs||[],token);await upsertNoDelete(l,TABLE.funcs,HEAD.funcs,b.funcs||[],token)}

async function backgroundFlush(){
  const owner='SW:'+((crypto&&crypto.randomUUID)?crypto.randomUUID():Math.random().toString(36).slice(2));
  if(!(await acquireSyncLock(owner)))return 0;
  const lockBeat=setInterval(()=>acquireSyncLock(owner).catch(()=>{}),30000);
  try{
  const c=await config();if(!c.authConfigured||!c.workbookUrl)return 0;
  let token;
  try{token=await accessToken()}catch(e){if(e.reauthRequired)return 0;throw e}
  const l=await workbookLocator(c.workbookUrl,token);
  const list=(await idbAll('outbox')).sort((a,b)=>String(a.createdAt).localeCompare(String(b.createdAt)));
  let done=0;
  for(const x of list){
    if(Number(x.nextAttemptAt||0)>Date.now())continue;
    try{
      if(x.type==='AP_DAY')await pushAp(l,x.bundle,token);
      else if(x.type==='FVS'){
        const stored=x.recordId?await idbGet('records',x.recordId):null;
        if(stored)await pushFvs(l,stored,token);
        else throw new Error('Registro FVS local não encontrado.');
      } else if(x.type==='CATALOG')await pushCatalog(l,x.bundle||{},token);
      await idbDel('outbox',x.id);
      if(x.type==='FVS'&&x.recordId){const r=await idbGet('records',x.recordId);if(r){r.syncedAt=new Date().toISOString();r.updatedAt=r.syncedAt;await idbPut('records',r);for(const key of (r.photoKeys||[])){const p=await idbGet('photos',key).catch(()=>null);if(p&&p.uploadBlob){delete p.uploadBlob;await idbPut('photos',p).catch(()=>{})}}}}
      done++;
    }catch(e){x.attempts=Number(x.attempts||0)+1;x.lastError=String(e?.message||e);x.nextAttemptAt=Date.now()+retryDelay(x.attempts);x.updatedAt=new Date().toISOString();await idbPut('outbox',x);throw e}
  }
  return done;
  } finally { clearInterval(lockBeat);await releaseSyncLock(owner).catch(()=>{}); }
}



// ---------- Outbox do ERP integrado (IndexedDB elevatta-erp) ----------
let erpDbp=null;
function erpDb(){
  if(erpDbp)return erpDbp;
  erpDbp=new Promise((res,rej)=>{
    const q=indexedDB.open('elevatta-erp',1);
    q.onupgradeneeded=()=>{const d=q.result;for(const st of ['cache','outbox','assets','meta'])if(!d.objectStoreNames.contains(st))d.createObjectStore(st,{keyPath:'key'})};
    q.onsuccess=()=>{const d=q.result;d.onversionchange=()=>{try{d.close()}catch{}erpDbp=null};res(d)};
    q.onerror=()=>{erpDbp=null;rej(q.error)};
  });return erpDbp;
}
async function erpGet(store,key){const d=await erpDb();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).get(key);q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}
async function erpAll(store){const d=await erpDb();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).getAll();q.onsuccess=()=>res(q.result||[]);q.onerror=()=>rej(q.error)})}
async function erpPut(store,val){const d=await erpDb();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).put(val);tx.oncomplete=()=>res(val);tx.onerror=()=>rej(tx.error)})}
async function erpDel(store,key){const d=await erpDb();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).delete(key);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
async function erpLock(owner,ttl=120000){const d=await erpDb();return new Promise((res,rej)=>{let ok=false;const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('flush-lock');q.onsuccess=()=>{const cur=q.result;if(!cur||Number(cur.expiresAt||0)<Date.now()||cur.owner===owner){ok=true;os.put({key:'flush-lock',owner,expiresAt:Date.now()+ttl,updatedAt:new Date().toISOString()})}};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res(ok);tx.onerror=()=>rej(tx.error)})}
async function erpUnlock(owner){const d=await erpDb();return new Promise((res,rej)=>{const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('flush-lock');q.onsuccess=()=>{if(q.result?.owner===owner)os.delete('flush-lock')};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
async function erpApi(path,opt={}){const init={method:opt.method||'GET',credentials:'include',cache:'no-store',headers:Object.assign({},opt.headers||{})};if(opt.body!==undefined){if(opt.raw)init.body=opt.body;else{init.headers['Content-Type']='application/json';init.body=JSON.stringify(opt.body)}}const r=await fetchTimed(path,init,opt.timeout||60000),ct=r.headers.get('content-type')||'',b=ct.includes('json')?await r.json().catch(()=>({})):await r.blob();if(!r.ok){const e=new Error(b?.error||('HTTP '+r.status));e.status=r.status;e.payload=b;throw e}return b}
function erpSetPath(obj,path,value){const p=String(path||'').split('.').filter(Boolean);let cur=obj;for(let i=0;i<p.length-1;i++){cur[p[i]]=cur[p[i]]||{};cur=cur[p[i]]}if(p.length)cur[p[p.length-1]]=value}
async function erpHydrate(tx){const payload=JSON.parse(JSON.stringify(tx.payload||{}));for(const a of (tx.assets||[])){const x=await erpGet('assets',a.assetKey);if(!x?.blob)throw Object.assign(new Error('Arquivo offline não localizado: '+a.assetKey),{status:409});const qs=new URLSearchParams({folder:a.folder||x.folder||'',name:a.name||x.name||'arquivo'}),r=await erpApi('/api/erp/upload?'+qs.toString(),{method:'POST',body:x.blob,raw:true,headers:{'Content-Type':x.type||x.blob.type||'application/octet-stream'},timeout:60000});if(a.target)erpSetPath(payload,a.target,r.appUrl||r.webUrl||r.url||'')}return payload}
function erpPermanent(e){const st=Number(e?.status||0);return st>=400&&st<500&&![401,408,425,429].includes(st)}
async function erpBackgroundFlush(){
  const owner='SW-ERP:'+((crypto&&crypto.randomUUID)?crypto.randomUUID():Math.random().toString(36).slice(2));if(!(await erpLock(owner)))return 0;
  const beat=setInterval(()=>erpLock(owner).catch(()=>{}),30000);let done=0;
  try{
    const list=(await erpAll('outbox')).sort((a,b)=>String(a.createdAt).localeCompare(String(b.createdAt)));
    for(const tx of list){if(tx.blocked||Number(tx.nextAttemptAt||0)>Date.now())continue;try{
      const payload=await erpHydrate(tx);let r;
      if(tx.type==='TOOL_CHECKOUT')r=await erpApi('/api/erp/tools/checkout',{method:'POST',body:payload});
      else if(tx.type==='TOOL_RETURN')r=await erpApi('/api/erp/tools/return',{method:'POST',body:payload});
      else if(tx.type==='IMPORT_COMMIT')r=await erpApi('/api/erp/import/commit',{method:'POST',body:payload,timeout:90000});
      else r=await erpApi('/api/erp/batch',{method:'POST',body:payload});
      await erpDel('outbox',tx.key);for(const a of (tx.assets||[]))await erpDel('assets',a.assetKey).catch(()=>{});done++;
    }catch(e){tx.attempts=Number(tx.attempts||0)+1;tx.lastError=e.message||String(e);tx.updatedAt=new Date().toISOString();tx.blocked=erpPermanent(e);tx.blockedAt=tx.blocked?tx.updatedAt:'';tx.nextAttemptAt=tx.blocked?0:(Date.now()+retryDelay(tx.attempts));await erpPut('outbox',tx);break}}
    return done;
  }finally{clearInterval(beat);await erpUnlock(owner).catch(()=>{})}
}

async function notifyClients(type='ELEVATTA_SYNC_NOW'){
  const clients=await self.clients.matchAll({type:'window',includeUncontrolled:true});
  for(const client of clients)client.postMessage({type});
}

self.addEventListener('sync',event=>{
  if(event.tag==='elevatta-outbox')event.waitUntil(backgroundFlush().then(()=>notifyClients('ELEVATTA_SYNC_DONE')));
  if(event.tag==='elevatta-erp-outbox')event.waitUntil(erpBackgroundFlush().then(()=>notifyClients('ELEVATTA_ERP_SYNC_DONE')));
});

self.addEventListener('message',event=>{
  if(event.data?.type==='SKIP_WAITING')self.skipWaiting();
  if(event.data?.type==='REQUEST_SYNC')event.waitUntil?.(Promise.all([backgroundFlush(),erpBackgroundFlush()]).then(()=>notifyClients('ELEVATTA_SYNC_DONE')));
});
