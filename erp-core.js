/* Elevatta ERP Core REV40
   Infraestrutura compartilhada entre PWA mobile e ERP desktop.
   - IDs determinísticos de frente
   - IndexedDB offline-first
   - outbox transacional
   - uploads diferidos (assinaturas/fotos/importações)
   - gateway server-side para SharePoint/Excel
*/
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports) module.exports=api;
  root.ElevattaERP=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const E={};
  const VERSION='40.1.0';
  const DB_NAME='elevatta-erp';
  const DB_VERSION=1;
  const STORE=['cache','outbox','assets','meta'];
  let dbp=null,flushing=null;
  const state={online:typeof navigator==='undefined'?true:navigator.onLine,busy:false,pending:0,blocked:0,lastOkAt:0,lastError:'',schema:null,storagePersistent:false,storageUsage:0,storageQuota:0};

  const now=()=>new Date().toISOString();
  const norm=v=>String(v??'').trim();
  const upper=v=>norm(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/\s+/g,' ');
  const sanitize=v=>norm(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9._-]+/g,'_').replace(/^_+|_+$/g,'').slice(0,100)||'arquivo';
  function fnv1a(str,seed=0x811c9dc5){let h=seed>>>0;for(let i=0;i<str.length;i++){h^=str.charCodeAt(i);h=Math.imul(h,0x01000193)>>>0;}return h>>>0}
  function hash12(str){const s=String(str);return fnv1a(s,0x811c9dc5).toString(16).padStart(8,'0')+fnv1a(s,0x9e3779b9).toString(16).padStart(8,'0').slice(0,4)}
  function makeId(prefix='ID'){try{if(globalThis.crypto?.randomUUID)return prefix+'-'+crypto.randomUUID().replace(/-/g,'').slice(0,20).toUpperCase()}catch{}return prefix+'-'+Date.now().toString(36).toUpperCase()+Math.random().toString(36).slice(2,10).toUpperCase()}
  function frontKey(x={}){return [upper(x.obra||x.OBRA),upper(x.pacote||x.pacoteCod||x.PACOTE_COD),upper(x.pavimento||x.PAVIMENTO),upper(x.local||x.localFrente||x.LOCAL_FRENTE||x.LOCAL)].join('|')}
  function frontId(x={}){const obra=upper(x.obra||x.OBRA),pac=upper(x.pacote||x.pacoteCod||x.PACOTE_COD),pav=upper(x.pavimento||x.PAVIMENTO),loc=upper(x.local||x.localFrente||x.LOCAL_FRENTE||x.LOCAL);if(!obra||!pac||(!pav&&!loc))return '';const k=[obra,pac,pav,loc].join('|');return 'FRT-'+hash12(k).toUpperCase()}
  function employeeId(matricula){const m=upper(matricula);return m?'FUN-'+hash12(m).toUpperCase():''}
  function obraId(obra){const o=upper(obra);return o?'OBR-'+hash12(o).toUpperCase():''}
  function packageId(codigo){return upper(codigo)}
  function deviceId(){let id='';try{id=localStorage.getItem('elevatta_device_id_v1')||'';if(!id){id=makeId('DEV');localStorage.setItem('elevatta_device_id_v1',id)}}catch{id='DEV-SEM-STORAGE'}return id}
  function userLabel(){try{const a=JSON.parse(localStorage.getItem('elevatta_m365_account_v1')||'{}');return a.email||a.name||''}catch{return ''}}

  function emit(type,detail={}){try{globalThis.dispatchEvent(new CustomEvent(type,{detail}))}catch{} }
  function setState(p){Object.assign(state,p);emit('elevatta-erp-status',Object.assign({},state))}
  E.state=state;E.version=VERSION;E.norm=norm;E.upper=upper;E.sanitize=sanitize;E.hash12=hash12;E.makeId=makeId;E.frontKey=frontKey;E.frontId=frontId;E.employeeId=employeeId;E.obraId=obraId;E.packageId=packageId;E.deviceId=deviceId;

  function db(){
    if(dbp)return dbp;
    if(typeof indexedDB==='undefined')return Promise.reject(new Error('IndexedDB indisponível neste navegador.'));
    dbp=new Promise((resolve,reject)=>{
      const r=indexedDB.open(DB_NAME,DB_VERSION);
      r.onupgradeneeded=()=>{const d=r.result;for(const s of STORE)if(!d.objectStoreNames.contains(s))d.createObjectStore(s,{keyPath:'key'})};
      r.onsuccess=()=>{const d=r.result;d.onversionchange=()=>{try{d.close()}catch{}dbp=null};resolve(d)};
      r.onerror=()=>{dbp=null;reject(r.error||new Error('Falha ao abrir armazenamento local.'))};
      r.onblocked=()=>console.warn('[Elevatta ERP] IndexedDB aguardando outra aba fechar.');
    });return dbp;
  }
  async function get(store,key){const d=await db();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).get(key);q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}
  async function all(store){const d=await db();return new Promise((res,rej)=>{const q=d.transaction(store,'readonly').objectStore(store).getAll();q.onsuccess=()=>res(q.result||[]);q.onerror=()=>rej(q.error)})}
  async function put(store,value){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).put(value);tx.oncomplete=()=>res(value);tx.onerror=()=>rej(tx.error)})}
  async function del(store,key){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction(store,'readwrite');tx.objectStore(store).delete(key);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
  async function acquireLock(owner,ttl=120000){const d=await db();return new Promise((res,rej)=>{let ok=false;const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('flush-lock');q.onsuccess=()=>{const cur=q.result;if(!cur||Number(cur.expiresAt||0)<Date.now()||cur.owner===owner){ok=true;os.put({key:'flush-lock',owner,expiresAt:Date.now()+ttl,updatedAt:now()})}};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res(ok);tx.onerror=()=>rej(tx.error)})}
  async function releaseLock(owner){const d=await db();return new Promise((res,rej)=>{const tx=d.transaction('meta','readwrite'),os=tx.objectStore('meta'),q=os.get('flush-lock');q.onsuccess=()=>{if(q.result?.owner===owner)os.delete('flush-lock')};q.onerror=()=>rej(q.error);tx.oncomplete=()=>res();tx.onerror=()=>rej(tx.error)})}
  E.db={get,all,put,del};

  async function api(path,opt={}){
    const ctl=new AbortController(),ms=Number(opt.timeout||30000),timer=setTimeout(()=>ctl.abort(),ms);
    try{
      const init={method:opt.method||'GET',credentials:'same-origin',cache:'no-store',signal:ctl.signal,headers:Object.assign({},opt.headers||{})};
      if(opt.body!==undefined){if(opt.raw)init.body=opt.body;else{init.headers['Content-Type']='application/json';init.body=JSON.stringify(opt.body)}}
      const r=await fetch(path,init),ct=r.headers.get('content-type')||'';const b=ct.includes('json')?await r.json().catch(()=>({})):await r.blob();
      if(!r.ok){const e=new Error(b?.error||b?.message||('HTTP '+r.status));e.status=r.status;e.reauthRequired=!!b?.reauthRequired;e.payload=b;throw e}return b;
    }catch(e){if(e?.name==='AbortError')throw new Error('Tempo de conexão esgotado; a operação permanece protegida no aparelho.');throw e}
    finally{clearTimeout(timer)}
  }
  E.api=api;

  async function loadSchema(force=false){if(state.schema&&!force)return state.schema;try{state.schema=await api('/api/erp/schema',{timeout:10000});await put('meta',{key:'schema',value:state.schema,updatedAt:now()});return state.schema}catch(e){const c=await get('meta','schema').catch(()=>null);if(c?.value){state.schema=c.value;return c.value}throw e}}
  E.loadSchema=loadSchema;

  async function cacheTable(table,rows){await put('cache',{key:'table:'+table,rows:Array.isArray(rows)?rows:[],updatedAt:now()});return rows}
  async function cachedTable(table){return (await get('cache','table:'+table).catch(()=>null))?.rows||[]}
  async function snapshot(tables=[],opt={}){
    const list=[...new Set((tables||[]).filter(Boolean))];
    if((typeof navigator!=='undefined'&&!navigator.onLine)||opt.localOnly){const out={offline:true,tables:{}};for(const t of list)out.tables[t]=await cachedTable(t);return out}
    try{const j=await api('/api/erp/snapshot?tables='+encodeURIComponent(list.join(',')),{timeout:opt.timeout||45000});for(const [t,rows] of Object.entries(j.tables||{}))await cacheTable(t,rows);setState({lastOkAt:Date.now(),lastError:''});return j}
    catch(e){const out={offline:true,stale:true,error:e.message,tables:{}};for(const t of list)out.tables[t]=await cachedTable(t);setState({lastError:e.message});return out}
  }
  E.snapshot=snapshot;E.cachedTable=cachedTable;E.cacheTable=cacheTable;

  async function saveAsset(blob,meta={}){if(!(blob instanceof Blob))throw new Error('Arquivo local inválido.');const key=meta.key||makeId('ASSET');await put('assets',{key,blob,name:meta.name||'arquivo',type:blob.type||meta.type||'application/octet-stream',folder:meta.folder||'',createdAt:now()});return key}
  async function asset(key){return get('assets',key)}
  E.saveAsset=saveAsset;

  async function queue(tx){
    const key=String(tx.key||tx.id||makeId('Q'));const old=await get('outbox',key).catch(()=>null);
    const item=Object.assign({},tx,{key,createdAt:old?.createdAt||tx.createdAt||now(),updatedAt:now(),attempts:Number(old?.attempts||0),nextAttemptAt:0,lastError:'',blocked:false,blockedAt:''});
    await put('outbox',item);await updatePending();requestSync();if(typeof navigator==='undefined'||navigator.onLine)setTimeout(()=>flush({silent:true}).catch(()=>{}),0);return item;
  }
  E.queue=queue;
  async function pending(){return all('outbox')}
  async function updatePending(){const list=await pending().catch(()=>[]),p=list.length,b=list.filter(x=>x.blocked).length;setState({pending:p,blocked:b});return p}
  E.pendingCount=updatePending;
  function retryDelay(n){return Math.min(15*60*1000,4000*Math.pow(2,Math.min(7,Math.max(0,n-1))))}
  async function requestSync(){try{const reg=await navigator.serviceWorker?.ready;if(reg?.sync)await reg.sync.register('elevatta-erp-outbox')}catch{}}

  async function uploadAsset(a){
    const x=await asset(a.assetKey);if(!x?.blob)throw new Error('Arquivo offline não localizado: '+a.assetKey);
    const qs=new URLSearchParams({folder:a.folder||x.folder||'',name:a.name||x.name||'arquivo'});
    return api('/api/erp/upload?'+qs.toString(),{method:'POST',body:x.blob,raw:true,headers:{'Content-Type':x.type||x.blob.type||'application/octet-stream'},timeout:60000});
  }
  function setPath(obj,path,value){const p=String(path||'').split('.').filter(Boolean);let cur=obj;for(let i=0;i<p.length-1;i++){cur[p[i]]=cur[p[i]]||{};cur=cur[p[i]]}if(p.length)cur[p[p.length-1]]=value}
  async function hydrateAssets(tx){const copy=JSON.parse(JSON.stringify(tx.payload||{})),uploaded=[];for(const a of (tx.assets||[])){const r=await uploadAsset(a);if(a.target)setPath(copy,a.target,r.appUrl||r.webUrl||r.url||'');uploaded.push({key:a.assetKey,result:r})}return {payload:copy,uploaded}}
  function permanentError(e){const st=Number(e?.status||0);return st>=400&&st<500&&![401,408,425,429].includes(st)}
  async function flush(opt={}){
    if(flushing)return flushing;
    flushing=(async()=>{
      if(typeof navigator!=='undefined'&&!navigator.onLine)return 0;
      const owner='PAGE:'+makeId('LOCK');if(!(await acquireLock(owner)))return 0;
      const beat=setInterval(()=>acquireLock(owner).catch(()=>{}),30000);
      try{
        setState({busy:true,lastError:''});let count=0;
        const list=(await pending()).sort((a,b)=>String(a.createdAt).localeCompare(String(b.createdAt)));
        for(const tx of list){if(tx.blocked)continue;if(!opt.force&&Number(tx.nextAttemptAt||0)>Date.now())continue;try{
          const h=await hydrateAssets(tx);let result;
          if(tx.type==='TOOL_CHECKOUT')result=await api('/api/erp/tools/checkout',{method:'POST',body:h.payload,timeout:60000});
          else if(tx.type==='TOOL_RETURN')result=await api('/api/erp/tools/return',{method:'POST',body:h.payload,timeout:60000});
          else if(tx.type==='IMPORT_COMMIT')result=await api('/api/erp/import/commit',{method:'POST',body:h.payload,timeout:90000});
          else result=await api('/api/erp/batch',{method:'POST',body:h.payload,timeout:60000});
          await del('outbox',tx.key);for(const a of (tx.assets||[]))await del('assets',a.assetKey).catch(()=>{});count++;setState({lastOkAt:Date.now(),lastError:''});emit('elevatta-erp-synced',{transaction:tx,result});
        }catch(e){tx.attempts=Number(tx.attempts||0)+1;tx.lastError=e.message||String(e);tx.updatedAt=now();tx.blocked=permanentError(e);tx.blockedAt=tx.blocked?now():'';tx.nextAttemptAt=tx.blocked?0:(Date.now()+retryDelay(tx.attempts));await put('outbox',tx);setState({lastError:tx.lastError});if(tx.blocked)emit('elevatta-erp-conflict',{transaction:tx,error:tx.lastError,status:e.status||0});break}}
        await updatePending();return count;
      }finally{clearInterval(beat);await releaseLock(owner).catch(()=>{});}
    })().finally(()=>{flushing=null;setState({busy:false})});return flushing;
  }
  E.flush=flush;
  E.retryTransaction=async key=>{const tx=await get('outbox',key);if(!tx)return false;tx.blocked=false;tx.blockedAt='';tx.lastError='';tx.nextAttemptAt=0;tx.updatedAt=now();await put('outbox',tx);await updatePending();requestSync();if(typeof navigator==='undefined'||navigator.onLine)flush({force:true}).catch(()=>{});return true};
  E.discardTransaction=async key=>{const tx=await get('outbox',key);if(!tx)return false;for(const a of (tx.assets||[]))await del('assets',a.assetKey).catch(()=>{});await del('outbox',key);await updatePending();emit('elevatta-erp-discarded',{transaction:tx});return true};

  async function sha256Blob(blob){try{const d=await crypto.subtle.digest('SHA-256',await blob.arrayBuffer());return [...new Uint8Array(d)].map(b=>b.toString(16).padStart(2,'0')).join('')}catch{return ''}}
  E.sha256Blob=sha256Blob;

  function frontRecord(x={}){const id=frontId(x),t=now();return {FRENTE_ID:id,OBRA_ID:x.obraId||obraId(x.obra),OBRA:norm(x.obra),PACOTE_COD:norm(x.pacote||x.pacoteCod),PACOTE_SERVICO:norm(x.pacoteServico),ETAPA:norm(x.etapa),PAVIMENTO:norm(x.pavimento),LOCAL_FRENTE:norm(x.local||x.localFrente),STATUS:norm(x.status||'ATIVA'),CRIADO_EM:x.criadoEm||t,UPDATED_AT:t}}
  E.frontRecord=frontRecord;

  async function toolCheckout(input={}){
    const tools=(input.tools||[]).filter(Boolean);if(!tools.length)throw new Error('Selecione pelo menos uma ferramenta.');
    if(!input.employee?.matricula&&!input.employee?.id)throw new Error('Selecione o funcionário que está retirando.');
    if(!input.obra)throw new Error('Selecione a obra.');if(!input.pacote?.codigo)throw new Error('Selecione o pacote de serviço.');if(!norm(input.local))throw new Error('Informe o local / frente onde a ferramenta será utilizada.');if(!norm(input.almoxarife))throw new Error('Informe quem entregou a ferramenta no almoxarifado.');
    const f=frontRecord({obra:input.obra,pacote:input.pacote.codigo,pacoteServico:input.pacote.nome,etapa:input.pacote.etapa,pavimento:input.pavimento,local:input.local});
    const termId=makeId('TERM'),time=input.dataHora||now(),dev=deviceId(),empId=input.employee.id||employeeId(input.employee.matricula),upd=now();
    const sigHash=norm(input.signatureHash)||(input.signatureBlob?await sha256Blob(input.signatureBlob):'');
    const term={TERMO_ID:termId,DATA_HORA_SAIDA:time,FUNCIONARIO_ID:empId,MATRICULA:norm(input.employee.matricula),FUNCIONARIO_NOME:norm(input.employee.nome),OBRA_ID:obraId(input.obra),OBRA:norm(input.obra),FRENTE_ID:f.FRENTE_ID,PACOTE_COD:norm(input.pacote.codigo),PACOTE_SERVICO:norm(input.pacote.nome),PAVIMENTO:norm(input.pavimento),LOCAL_FRENTE:norm(input.local),ASSINATURA_URL:'',ASSINATURA_HASH:sigHash,RESPONSAVEL_ALMOXARIFADO:norm(input.almoxarife),STATUS:'ABERTO',DATA_HORA_FECHAMENTO:'',OBSERVACAO:norm(input.observacao),DEVICE_ID:dev,UPDATED_AT:upd};
    const movements=tools.map(t=>({MOVIMENTO_ID:makeId('MOV'),TERMO_ID:termId,FERRAMENTA_ID:t.FERRAMENTA_ID||t.id,NUMERO:Number(t.NUMERO??t.numero)||0,FERRAMENTA_NOME:t.NOME||t.nome||'',TIPO:'SAIDA',DATA_HORA:time,FUNCIONARIO_ID:empId,FUNCIONARIO_NOME:norm(input.employee.nome),OBRA_ID:obraId(input.obra),OBRA:norm(input.obra),FRENTE_ID:f.FRENTE_ID,PACOTE_COD:norm(input.pacote.codigo),PACOTE_SERVICO:norm(input.pacote.nome),PAVIMENTO:norm(input.pavimento),LOCAL_FRENTE:norm(input.local),CONDICAO:norm(t.CONDICAO_ATUAL||t.condicao||'OK'),STATUS:'EM USO',FOTO_URL:'',OBSERVACAO:norm(input.observacao),RECEBIDO_POR:'',DEVICE_ID:dev,UPDATED_AT:upd}));
    const payload={term,front:f,movements,toolIds:movements.map(m=>m.FERRAMENTA_ID),user:userLabel()};const assets=[];
    if(input.signatureBlob){const key=await saveAsset(input.signatureBlob,{name:termId+'_assinatura.png',folder:'Assinaturas/Ferramentas/'+termId});assets.push({assetKey:key,folder:'Assinaturas/Ferramentas/'+termId,name:termId+'_assinatura.png',target:'term.ASSINATURA_URL'})}
    await queue({key:'TOOL-CHECKOUT-'+termId,type:'TOOL_CHECKOUT',payload,assets});
    return {termId,frontId:f.FRENTE_ID,term,movements,pending:true};
  }
  E.toolCheckout=toolCheckout;

  async function toolReturn(input={}){
    const items=(input.items||[]).filter(Boolean);if(!input.termId)throw new Error('Termo de retirada não informado.');if(!items.length)throw new Error('Selecione pelo menos uma ferramenta para devolver.');if(!norm(input.recebidoPor))throw new Error('Informe quem recebeu a devolução no almoxarifado.');
    const dev=deviceId(),time=input.dataHora||now(),upd=now(),movements=[],assets=[];
    for(const t of items){const m={MOVIMENTO_ID:makeId('MOV'),TERMO_ID:input.termId,FERRAMENTA_ID:t.FERRAMENTA_ID||t.id,NUMERO:Number(t.NUMERO??t.numero)||0,FERRAMENTA_NOME:t.NOME||t.nome||'',TIPO:'DEVOLUCAO',DATA_HORA:time,FUNCIONARIO_ID:norm(input.employeeId),FUNCIONARIO_NOME:norm(input.employeeName),OBRA_ID:norm(input.obraId),OBRA:norm(input.obra),FRENTE_ID:norm(input.frenteId),PACOTE_COD:norm(input.pacoteCod),PACOTE_SERVICO:norm(input.pacoteServico),PAVIMENTO:norm(input.pavimento),LOCAL_FRENTE:norm(input.local),CONDICAO:norm(t.returnCondition||input.condicao||'OK'),STATUS:'DEVOLVIDA',FOTO_URL:'',OBSERVACAO:norm(t.observacao||input.observacao),RECEBIDO_POR:norm(input.recebidoPor),DEVICE_ID:dev,UPDATED_AT:upd};movements.push(m);if(t.photoBlob){const key=await saveAsset(t.photoBlob,{name:m.MOVIMENTO_ID+'.jpg',folder:'Ferramentas/Devolucoes/'+input.termId});assets.push({assetKey:key,folder:'Ferramentas/Devolucoes/'+input.termId,name:m.MOVIMENTO_ID+'.jpg',target:'movements.'+(movements.length-1)+'.FOTO_URL'})}}
    const payload={termId:input.termId,movements,toolIds:movements.map(m=>m.FERRAMENTA_ID),receivedAt:time,user:userLabel()};await queue({key:'TOOL-RETURN-'+input.termId+'-'+Date.now(),type:'TOOL_RETURN',payload,assets});return {termId:input.termId,movements,pending:true};
  }
  E.toolReturn=toolReturn;

  async function commitImport(report,file){
    if(!report?.importacao)throw new Error('Relatório de validação inválido.');if(['REJEITADO','BLOQUEADO'].includes(report.importacao.STATUS)||Number(report.importacao.CRITICOS||0)>0)throw new Error('Arquivo bloqueado: corrija os erros críticos antes de importar.');
    const payload=JSON.parse(JSON.stringify(report));payload.user=userLabel();payload.deviceId=deviceId();const assets=[];
    if(file){const ext=(file.name||'apontamento.xlsx').split('.').pop()||'xlsx',name=report.importacao.IMPORTACAO_ID+'_'+sanitize(file.name||('arquivo.'+ext));const key=await saveAsset(file,{name,folder:'Importacoes/'+report.importacao.IMPORTACAO_ID});assets.push({assetKey:key,folder:'Importacoes/'+report.importacao.IMPORTACAO_ID,name,target:'importacao.ARQUIVO_URL'})}
    await queue({key:'IMPORT-'+report.importacao.IMPORTACAO_ID,type:'IMPORT_COMMIT',payload,assets});return report.importacao.IMPORTACAO_ID;
  }
  E.commitImport=commitImport;

  async function localToolState(){const [tools,terms,movs]=await Promise.all([cachedTable('ferramentas'),cachedTable('termosFerramentas'),cachedTable('movimentosFerramentas')]);return {tools,terms,movements:movs}}
  E.localToolState=localToolState;

  async function boot(){
    await updatePending().catch(()=>{});loadSchema().catch(()=>{});
    if(typeof navigator!=='undefined'){
      let persistent=false,usage=0,quota=0;
      try{if(navigator.storage?.persist)persistent=await navigator.storage.persist()}catch{}
      try{if(navigator.storage?.persisted)persistent=await navigator.storage.persisted()}catch{}
      try{const est=await navigator.storage?.estimate?.();usage=Number(est?.usage||0);quota=Number(est?.quota||0)}catch{}
      setState({online:navigator.onLine,storagePersistent:persistent,storageUsage:usage,storageQuota:quota});globalThis.addEventListener('online',()=>{setState({online:true});flush({force:true,silent:true}).catch(()=>{})});globalThis.addEventListener('offline',()=>setState({online:false}));
      if(typeof document!=='undefined')document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&navigator.onLine)flush({silent:true}).catch(()=>{})});
      setInterval(()=>{if(navigator.onLine)flush({silent:true}).catch(()=>{})},60000);
      navigator.serviceWorker?.addEventListener?.('message',e=>{if(e.data?.type==='ELEVATTA_ERP_SYNC')flush({force:true,silent:true}).catch(()=>{})});
    }
  }
  E.boot=boot;
  if(typeof document!=='undefined'){if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot,{once:true});else setTimeout(boot,0)}
  return E;
});
