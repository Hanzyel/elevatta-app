/* Elevatta Validator REV40 — motor de validação de apontamentos.
   Usa ExcelJS já embarcado no PWA. Não depende do nome das abas. */
(function(root,factory){const api=factory(root);if(typeof module==='object'&&module.exports)module.exports=api;root.ElevattaValidator=api})(typeof globalThis!=='undefined'?globalThis:this,function(root){
  'use strict';
  const V={version:'40.1.0'};
  const norm=x=>String(x??'').trim();
  const key=x=>norm(x).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');
  const upper=x=>norm(x).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');
  const COLS={
    data:['DATA'],mat:['MATRICULA'],nome:['NOME'],funcao:['FUNCAO'],regime:['REGIME'],obra:['OBRA','OBRAS'],etapa:['ETAPA'],cod:['COD','CODPACOTE','CODIGO','CODPACOTE'],serv:['SERVICOPACOTE','SERVICO','PACOTE'],local:['LOCALFRENTE','LOCAL','FRENTE'],pav:['PAVIMENTO','PAV'],equipe:['EQUIPE'],fr:['FRACAODODIA','FRACAO'],cl:['CLASSIFICACAO'],sit:['SITUACAODOPACOTE','SITUACAO'],obs:['OBSERVACAO','OBS'],resp:['RESPONSAVEL','ENCARREGADO']
  };
  function mapHeader(row){const out={};(row||[]).forEach((h,i)=>{const k=key(h);if(!k)return;for(const f in COLS)if(out[f]===undefined&&COLS[f].includes(k)){out[f]=i;break}});return out}
  function findHeader(rows){for(let i=0;i<Math.min(45,rows.length);i++){const ks=(rows[i]||[]).map(key);if(ks.includes('MATRICULA')&&ks.includes('NOME')&&(ks.includes('FRACAODODIA')||ks.includes('FRACAO')))return i}return -1}
  function isoValid(y,m,d){y=+y;m=+m;d=+d;if(!Number.isInteger(y)||!Number.isInteger(m)||!Number.isInteger(d)||y<2000||y>2100||m<1||m>12||d<1||d>31)return'';const dt=new Date(Date.UTC(y,m-1,d));return dt.getUTCFullYear()===y&&dt.getUTCMonth()===m-1&&dt.getUTCDate()===d?String(y).padStart(4,'0')+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0'):''}
  function iso(x){if(x instanceof Date&&!isNaN(x))return isoValid(x.getUTCFullYear(),x.getUTCMonth()+1,x.getUTCDate());if(typeof x==='number'&&isFinite(x)&&x>20000&&x<80000){const d=new Date(Math.round((x-25569)*86400000));return isoValid(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate())}const s=norm(x);let m=s.match(/^(\d{2})[\/-](\d{2})[\/-](\d{4})(?:\D|$)/);if(m)return isoValid(m[3],m[2],m[1]);m=s.match(/^(\d{4})-(\d{2})-(\d{2})(?:\D|$)/);return m?isoValid(m[1],m[2],m[3]):''}
  function number(x){if(typeof x==='number')return Number.isFinite(x)?x:NaN;let s=norm(x);if(!s)return NaN;const pct=/%$/.test(s);s=s.replace(/R\$/gi,'').replace(/%/g,'').replace(/\s/g,'').replace(/[^0-9,.-]/g,'');const ci=s.lastIndexOf(','),di=s.lastIndexOf('.');if(ci>=0&&di>=0)s=ci>di?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');else if(ci>=0)s=s.replace(/\./g,'').replace(',','.');const n=Number(s);return Number.isFinite(n)?(pct?n/100:n):NaN}
  function topDate(rows,end){for(let i=0;i<end;i++){for(let j=0;j<(rows[i]||[]).length;j++){if(key(rows[i][j])==='DATA'){for(let k=j+1;k<(rows[i]||[]).length;k++){const d=iso(rows[i][k]);if(d)return d}}}}return''}
  function cellValue(v){if(v instanceof Date)return v;if(v&&typeof v==='object'){if(v.formula!==undefined||v.sharedFormula!==undefined)return v.result??'';if(v.result!==undefined)return v.result;if(v.richText)return v.richText.map(t=>t.text).join('');if(v.text!==undefined)return v.text;return ''}return v}
  function sheetRows(ws){const rows=[];ws.eachRow({includeEmpty:true},r=>{const a=[];r.eachCell({includeEmpty:true},c=>a.push(cellValue(c.value)));rows.push(a)});return rows}
  async function sha256(file){if(root.crypto?.subtle&&file?.arrayBuffer){const b=await file.arrayBuffer(),d=await crypto.subtle.digest('SHA-256',b);return [...new Uint8Array(d)].map(x=>x.toString(16).padStart(2,'0')).join('')}return 'NOHASH-'+Date.now()}
  function issue(sev,type,o={}){return {OCORRENCIA_ID:(root.ElevattaERP?.makeId?.('OCO')||('OCO-'+Math.random().toString(36).slice(2))),SEVERIDADE:sev,TIPO:type,DATA:o.data||'',ARQUIVO:o.file||'',LINHA:o.line||0,CAMPO:o.field||'',VALOR:norm(o.value),DETALHE:o.detail||'',ACAO:o.action||'',UPDATED_AT:new Date().toISOString()}}
  function sig(l){return [l.DATA,l.MATRICULA,l.OBRA,l.PACOTE_COD,l.PAVIMENTO,l.LOCAL_FRENTE,l.FRACAO_DIA,l.CLASSIFICACAO,l.MOTIVO_FALTA].map(upper).join('|')}
  function catalogMaps(catalog={}){
    const funcs=new Map((catalog.funcionarios||[]).map(x=>[upper(x.MATRICULA||x.matricula),x]));
    const packs=new Map(),packNames=new Map();
    for(const x of (catalog.pacotes||[])){
      const code=upper(x.CODIGO||x.codigo),name=key(x.PACOTE_SERVICO||x.pacote_servico||x.NOME||x.nome);
      if(code)packs.set(code,x);
      if(name){const list=packNames.get(name)||[];list.push(x);packNames.set(name,list)}
    }
    const funcAliases=new Map();
    for(const a of (catalog.aliases||[])){
      const srcMat=upper(a.sourceMat||a.matriculaOrigem),srcName=key(a.sourceName||a.nomeOrigem),target=funcs.get(upper(a.targetMat||a.matriculaDestino));
      if(srcMat&&srcName&&target)funcAliases.set(srcMat+'|'+srcName,target);
    }
    return {funcs,packs,packNames,funcAliases};
  }
  function resolvePackage(cod,serv,maps){
    const raw=norm(cod),placeholder=!raw||raw==='-'||raw==='—';
    if(!placeholder){const p=maps.packs.get(upper(raw));return {code:raw,pack:p||null,source:p?'CODIGO':'NAO_ENCONTRADO'}}
    const byName=maps.packNames.get(key(serv))||[];
    if(byName.length===1){const p=byName[0];return {code:norm(p.CODIGO||p.codigo),pack:p,source:'NOME_UNICO'}}
    return {code:'',pack:null,source:byName.length>1?'NOME_AMBIGUO':'SEM_CORRESPONDENCIA'};
  }
  function validateRows(rows,fileName,hash,catalog){
    const issues=[],valid=[],quarantine=[],maps=catalogMaps(catalog),header=findHeader(rows);
    if(header<0){issues.push(issue('CRITICO','ESTRUTURA_NAO_RECONHECIDA',{file:fileName,detail:'Não foi encontrada tabela contendo MATRÍCULA, NOME e FRAÇÃO DO DIA.',action:'Use uma exportação do módulo Apontamento ou ajuste o cabeçalho.'}));return {valid,quarantine,issues}}
    const m=mapHeader(rows[header]),dTop=topDate(rows,header);
    if(m.mat===undefined||m.nome===undefined||m.fr===undefined||(m.cod===undefined&&m.serv===undefined&&m.cl===undefined)){issues.push(issue('CRITICO','COLUNAS_OBRIGATORIAS_AUSENTES',{file:fileName,line:header+1,detail:'O arquivo não possui todas as colunas mínimas para validação.',action:'Reexporte o apontamento no Elevatta.'}));return {valid,quarantine,issues}}
    const optional=['obra','etapa','pav','equipe','local'];for(const f of optional)if(m[f]===undefined)issues.push(issue('ALERTA','COLUNA_OPCIONAL_AUSENTE',{file:fileName,line:header+1,field:f.toUpperCase(),detail:'A dimensão '+f+' não está disponível neste layout.',action:'Atualize o modelo para melhorar os vínculos do ERP.'}));
    for(let i=header+1;i<rows.length;i++){
      const r=rows[i]||[],line=i+1,first=upper(r[0]),rawMat=norm(r[m.mat]);
      if(/^EQUIPES DO DIA/.test(first))break;
      if(/^LANCAMENTOS ADICIONAIS/.test(first))continue;
      const probeNome=upper(r[m.nome]),probeFunc=m.funcao!==undefined?upper(r[m.funcao]):'',probeReg=m.regime!==undefined?upper(r[m.regime]):'';
      if(!first&&!rawMat&&probeNome&&((probeFunc&&probeNome===probeFunc&&probeNome===probeReg)||/ASSINATURA DO COORDENADOR/.test(probeNome)))break;
      const data=m.data!==undefined?iso(r[m.data]):dTop;
      const cod=norm(m.cod!==undefined?r[m.cod]:''),serv=norm(m.serv!==undefined?r[m.serv]:''),cl=norm(m.cl!==undefined?r[m.cl]:'');const fr=number(r[m.fr]);const falta=/^(FALTA|ATESTADO|FERIAS|FOLGA|CHUVA)/.test(key(cl));
      const semCodigo=!cod||cod==='-'||cod==='—',temServico=!semCodigo||!!serv,temFracao=Number.isFinite(fr)&&fr>0,temOperacao=falta||temServico||temFracao||!!cl;
      /* Linhas do quadro sem apontamento são apenas cadastro visual no modelo antigo. */
      if(!temOperacao)continue;
      const matricula=rawMat.replace(/\.0+$/,'').replace(/\s+/g,'');
      const problems=[];const add=(t,field,val,detail,action)=>problems.push(issue('CRITICO',t,{file:fileName,line,data,field,value:val,detail,action}));
      if(!rawMat)add('MATRICULA_AUSENTE','MATRICULA','', 'Há trabalho/ausência preenchido, mas a matrícula está vazia.','Cadastrar ou informar a matrícula antes de consolidar.');
      else if(!/^[0-9][A-Z0-9._\/-]*$/i.test(matricula))add('MATRICULA_INVALIDA','MATRICULA',rawMat,'Matrícula inválida.','Corrigir a matrícula na origem.');
      if(!data)add('DATA_INVALIDA','DATA',m.data!==undefined?r[m.data]:'','Não foi possível determinar uma data válida.','Corrigir a data ou o cabeçalho do arquivo.');
      if(falta&&temServico)add('AUSENCIA_COM_TRABALHO','CLASSIFICACAO',cl,'A mesma linha contém ausência e serviço.','Manter somente a situação real.');
      if(falta&&temFracao)add('AUSENCIA_COM_TRABALHO','CLASSIFICACAO',cl,'A mesma linha contém ausência e fração trabalhada.','Manter somente a situação real.');
      if(!falta&&!temServico&&temFracao)add('SERVICO_AUSENTE','PACOTE_COD',cod,'Existe esforço sem pacote/serviço.','Selecionar o pacote correto.');
      if(!falta&&temServico&&(!Number.isFinite(fr)||fr<=0||fr>1.000001))add('FRACAO_INVALIDA','FRACAO_DIA',r[m.fr],'Para trabalho, a fração deve ser > 0 e <= 1.','Corrigir a fração do dia.');
      if(problems.length){quarantine.push({line,row:r,problems});issues.push(...problems);continue}
      const obra=norm(m.obra!==undefined?r[m.obra]:'');const etapa=norm(m.etapa!==undefined?r[m.etapa]:'');const pav=norm(m.pav!==undefined?r[m.pav]:'');const local=norm(m.local!==undefined?r[m.local]:'');const equipe=norm(m.equipe!==undefined?r[m.equipe]:'');const nome=norm(r[m.nome]);const funcao=norm(m.funcao!==undefined?r[m.funcao]:'');const regime=norm(m.regime!==undefined?r[m.regime]:'');const situacao=norm(m.sit!==undefined?r[m.sit]:'');const obs=norm(m.obs!==undefined?r[m.obs]:'');
      const pres=falta?{code:'',pack:null,source:'AUSENCIA'}:resolvePackage(cod,serv,maps),pcat=pres.pack,canonCode=pres.code,canonServ=norm(pcat?.PACOTE_SERVICO||serv),canonEtapa=norm(pcat?.ETAPA||etapa);
      if(!falta&&maps.packs.size){
        if(pres.source==='NAO_ENCONTRADO')issues.push(issue('ALERTA','PACOTE_NAO_CADASTRADO',{file:fileName,line,data,field:'PACOTE_COD',value:cod,detail:'O código não foi encontrado no cadastro central de pacotes.',action:'Conferir o código ou criar um de/para antes da consolidação.'}));
        else if(pres.source==='NOME_AMBIGUO')issues.push(issue('ALERTA','PACOTE_DEPARA_NECESSARIO',{file:fileName,line,data,field:'PACOTE_SERVICO',value:serv,detail:'O serviço sem código corresponde a mais de um pacote cadastrado.',action:'Definir manualmente qual pacote representa este serviço.'}));
        else if(pres.source==='SEM_CORRESPONDENCIA')issues.push(issue('ALERTA','PACOTE_SEM_CODIGO',{file:fileName,line,data,field:'PACOTE_SERVICO',value:serv,detail:'O serviço não possui código e não teve correspondência única pelo nome no cadastro central.',action:'Definir um de/para para preservar o vínculo com FVS, medição e ferramentas.'}));
      }
      const directEmp=matricula?maps.funcs.get(upper(matricula)):null,aliasEmp=matricula?maps.funcAliases.get(upper(matricula)+'|'+key(nome)):null,canonEmp=directEmp||aliasEmp||null;
      if(matricula&&maps.funcs.size&&!canonEmp){const o=issue('ALERTA','FUNCIONARIO_NAO_CADASTRADO',{file:fileName,line,data,field:'MATRICULA',value:matricula,detail:'A matrícula não foi encontrada no cadastro central para '+(nome||'este colaborador')+'.',action:'Defina o de/para somente após confirmar a identidade do colaborador.'});o.NOME_ORIGEM=nome;issues.push(o)}
      else if(aliasEmp&&!directEmp){const o=issue('INFO','FUNCIONARIO_ALIAS_APLICADO',{file:fileName,line,data,field:'MATRICULA',value:matricula,detail:(matricula+' · '+nome)+' foi reconciliado com '+(aliasEmp.MATRICULA||'')+' · '+(aliasEmp.NOME||'')+'.',action:'Nenhuma ação necessária; de/para administrativo aplicado.'});o.NOME_ORIGEM=nome;issues.push(o)}
      if(!obra&&!falta)issues.push(issue('ALERTA','OBRA_AUSENTE',{file:fileName,line,data,field:'OBRA',detail:'A obra não está informada.',action:'Informar a obra para permitir vínculos com FVS e ferramentas.'}));
      if(!pav&&!local&&!falta)issues.push(issue('ALERTA','LOCAL_AUSENTE',{file:fileName,line,data,field:'LOCAL_FRENTE',detail:'Pavimento/local não informado.',action:'Informar o local para aumentar a precisão da FRENTE_ID.'}));
      const front=root.ElevattaERP?.frontId?root.ElevattaERP.frontId({obra,pacote:canonCode,pavimento:pav,local}):'';const upd=new Date().toISOString();
      valid.push({ID:'IMPAP-'+hash.slice(0,12).toUpperCase()+'-'+String(line).padStart(4,'0'),DATA:data,RESPONSAVEL:norm(m.resp!==undefined?r[m.resp]:''),MATRICULA:norm(canonEmp?.MATRICULA||matricula),NOME:norm(canonEmp?.NOME||nome),FUNCAO:funcao||norm(canonEmp?.FUNCAO),REGIME:regime||norm(canonEmp?.REGIME),OBRA:obra,ETAPA:canonEtapa,PACOTE_COD:canonCode,PACOTE_SERVICO:canonServ,PAVIMENTO:pav,EQUIPE:equipe,FRACAO_DIA:falta?0:Math.round(fr*10000)/10000,CLASSIFICACAO:cl,SITUACAO:situacao,OBSERVACAO:obs,MOTIVO_FALTA:falta?(cl||'Falta'):'',SAIU_CEDO:'NAO',UPDATED_AT:upd,FRENTE_ID:front,LOCAL_FRENTE:local,DEVICE_ID:root.ElevattaERP?.deviceId?.()||'',SYNC_VERSION:'40',PACOTE_ORIGEM:pres.source,MATRICULA_ORIGEM:matricula,NOME_ORIGEM:nome});
    }
    const seen=new Map();for(const l of valid){const s=sig(l);if(seen.has(s))issues.push(issue('ALERTA','LINHA_DUPLICADA',{file:fileName,line:0,data:l.DATA,detail:'Há lançamento operacional idêntico dentro do arquivo.',action:'Conferir se é repetição ou fracionamento intencional.'}));else seen.set(s,l)}
    const pd=new Map(),types=new Map();for(const l of valid){const k=l.DATA+'|'+l.MATRICULA;if(l.MOTIVO_FALTA){const s=types.get(k)||new Set();s.add('F');types.set(k,s)}else{pd.set(k,(pd.get(k)||0)+Number(l.FRACAO_DIA||0));const s=types.get(k)||new Set();s.add('T');types.set(k,s)}}
    for(const [k,v] of pd)if(v>1.001){const [data,mat]=k.split('|');issues.push(issue('CRITICO','PESSOA_ACIMA_1_DH',{file:fileName,data,field:'MATRICULA',value:mat,detail:'A pessoa soma '+v.toFixed(2)+' dia-homem na mesma data.',action:'Revisar os fracionamentos antes de consolidar.'}))}
    for(const [k,s] of types)if(s.has('F')&&s.has('T')){const [data,mat]=k.split('|');issues.push(issue('CRITICO','AUSENCIA_E_TRABALHO_NO_DIA',{file:fileName,data,field:'MATRICULA',value:mat,detail:'A pessoa aparece trabalhando e ausente na mesma data.',action:'Corrigir o apontamento diário.'}))}
    return {valid,quarantine,issues};
  }

  async function parseFile(file,catalog={}){
    if(!file)throw new Error('Selecione um arquivo.');const ext=(file.name||'').toLowerCase();if(!/\.(xlsx|csv)$/.test(ext))throw new Error('Formato não suportado. Use .xlsx ou .csv.');
    const hash=await sha256(file),importId='IMP-'+hash.slice(0,20).toUpperCase(),allRows=[];let structuralIssues=[];
    if(ext.endsWith('.csv')){let txt=await file.text();if(txt.charCodeAt(0)===0xfeff)txt=txt.slice(1);const first=txt.split(/\r?\n/)[0]||'',sep=(first.split(';').length>=first.split(',').length)?';':',';let row=[],field='',quoted=false;for(let i=0;i<txt.length;i++){const c=txt[i];if(quoted){if(c==='"'&&txt[i+1]==='"'){field+='"';i++}else if(c==='"')quoted=false;else field+=c}else if(c==='"')quoted=true;else if(c===sep){row.push(field);field=''}else if(c==='\n'){row.push(field.replace(/\r$/,''));allRows.push(row);row=[];field=''}else field+=c}if(field||row.length){row.push(field);allRows.push(row)}}
    else{
      if(!root.ExcelJS)throw new Error('ExcelJS não carregado.');const wb=new root.ExcelJS.Workbook();await wb.xlsx.load(await file.arrayBuffer());let chosen=null,best=-1;wb.eachSheet(ws=>{const rows=sheetRows(ws),h=findHeader(rows);if(h>=0){const score=Object.keys(mapHeader(rows[h])).length;if(score>best){best=score;chosen=rows}}});if(!chosen){structuralIssues.push(issue('CRITICO','PLANILHA_SEM_APONTAMENTO',{file:file.name,detail:'Nenhuma aba contém a estrutura de apontamento reconhecida.',action:'Use o arquivo exportado pelo módulo Apontamento.'}));}else for(const row of chosen)allRows.push(row)
    }
    const r=allRows.length?validateRows(allRows,file.name,hash,catalog):{valid:[],quarantine:[],issues:structuralIssues};r.issues=[...structuralIssues,...r.issues];
    const crit=r.issues.filter(x=>x.SEVERIDADE==='CRITICO').length,alerts=r.issues.filter(x=>x.SEVERIDADE==='ALERTA').length,dates=r.valid.map(x=>x.DATA).filter(Boolean).sort();let status='APROVADO';if(!r.valid.length||structuralIssues.length)status='REJEITADO';else if(crit)status='BLOQUEADO';else if(alerts||r.quarantine.length)status='APROVADO_COM_ALERTAS';const upd=new Date().toISOString();
    const importacao={IMPORTACAO_ID:importId,ARQUIVO_NOME:file.name,ARQUIVO_HASH:hash,DATA_INICIO:dates[0]||'',DATA_FIM:dates[dates.length-1]||'',STATUS:status,LINHAS_TOTAL:Math.max(0,allRows.length),LINHAS_ACEITAS:r.valid.length,LINHAS_QUARENTENA:r.quarantine.length,CRITICOS:crit,ALERTAS:alerts,ORIGEM_DEVICE:root.ElevattaERP?.deviceId?.()||'',IMPORTADO_POR:'',ARQUIVO_URL:'',UPDATED_AT:upd};
    for(const o of r.issues)o.IMPORTACAO_ID=importId;
    return {importacao,apontamentos:r.valid,ocorrencias:r.issues,quarentena:r.quarantine.map(q=>({line:q.line,problems:q.problems.map(x=>x.TIPO)}))};
  }
  V.parseFile=parseFile;V.validateRows=validateRows;V.findHeader=findHeader;V.mapHeader=mapHeader;V.iso=iso;V.number=number;V.sha256=sha256;return V;
});
