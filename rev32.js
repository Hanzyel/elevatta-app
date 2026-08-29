/* Elevatta REV32 FINAL — fluxo guiado + Excel profissional de FVS/Medição Física */
(function guidedFlow(){
  const $ = id => document.getElementById(id);
  const esc = x => String(x ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const pct = x => (Number(x) || 0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}) + '%';
  let built = false;
  let current = 1;
  let modeChosen = false;
  let lastService = '';

  const M = () => window.ElevattaMeasurement;
  function snap(){
    try { return M()?._getSnapshot?.() || window.RF?.auditSnapshot?.() || null; }
    catch(e){ return null; }
  }
  const mode = () => M()?._getMode?.() || 'fvs';
  const state = () => M()?._getState?.() || {};
  const calc = () => M()?._getCalc?.() || {previous:0,executed:0,approved:0,measured:0,accumulated:0,balance:100,pendingQuality:0,approvedUnmeasured:0};

  function obraOk(){
    const v = $('r_h_obra')?.value || '';
    const o = $('r_h_obra2')?.value?.trim() || '';
    return !!(v && (v !== 'OUT' || o));
  }
  const respOk = () => !!($('r_h_resp')?.value || '').trim();
  const stageOk = () => !!($('r_pkg_stage')?.value || '');
  const serviceOk = () => stageOk() && !!($('r_svc')?.value || '');
  function quality(){
    const s = snap();
    const q = s?.quality || {};
    return {C:+q.C||0, NC:+q.NC||0, NA:+q.NA||0, P:+q.P||0, photos:s?.photos?.length||0};
  }
  function inspectionOk(){
    const q = quality();
    return q.photos > 0 && q.P === 0;
  }
  function execOk(){
    const st = state();
    return mode() !== 'med' || (String(st.local || '').trim().length > 0 && Number(st.executed) > 0);
  }
  function closeValid(){
    if(mode() !== 'med') return true;
    return (M()?._validate?.() || []).length === 0;
  }
  function stepComplete(n){
    if(n === 1) return obraOk() && respOk();
    if(n === 2) return serviceOk() && modeChosen;
    if(n === 3) return execOk();
    if(n === 4) return inspectionOk();
    return closeValid();
  }
  const visibleSteps = () => mode() === 'med' ? [1,2,3,4,5] : [1,2,4,5];
  const stageEl = n => $('rf_stage' + n);

  function makeStage(n,title,desc){
    const el = document.createElement('section');
    el.className = 'rf-stage';
    el.id = 'rf_stage' + n;
    el.innerHTML = `
      <div class="rf-stage-head">
        <div class="rf-stage-num">${n}</div>
        <div><h3>${title}</h3><p>${desc}</p></div>
      </div>
      <div class="rf-stage-body"></div>
      <div class="rf-stage-nav"></div>`;
    return el;
  }

  function nav(stage, back, next, nextText){
    const box = stage.querySelector('.rf-stage-nav');
    box.innerHTML = '';
    if(back){
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = '← Voltar';
      b.onclick = () => go(back);
      box.appendChild(b);
    }
    const n = document.createElement('button');
    n.type = 'button';
    n.className = 'next';
    n.textContent = nextText || 'Continuar →';
    n.onclick = () => go(next);
    box.appendChild(n);
    if(!back) box.classList.add('one');
  }

  function build(){
    if(built || !$('r_flow')) return;
    built = true;
    const flow = $('r_flow');
    const parent = flow.parentElement;

    const shell = document.createElement('div');
    shell.id = 'rf_guide';
    shell.className = 'rf-guide-shell';
    shell.innerHTML = `
      <div class="rf-guide-top">
        <div>
          <div class="rf-guide-kicker">FLUXO GUIADO</div>
          <div class="rf-guide-title">FVS e Medição Física</div>
          <div class="rf-guide-sub" id="rf_guide_sub">Preencha uma etapa por vez. A próxima só é liberada quando houver informação suficiente.</div>
        </div>
        <button class="rf-guide-base" id="rf_guide_base" type="button"><b>BASE</b><small>local</small></button>
      </div>
      <div class="rf-guide-steps" id="rf_guide_steps"></div>
      <div class="rf-guide-progressline"><span id="rf_guide_progressline"></span></div>
      <div class="rf-guide-counter" id="rf_guide_counter">Etapa 1</div>`;
    parent.insertBefore(shell, flow);
    $('rf_guide_base').onclick = () => window.ElevattaM365?.open('records');

    const defs = [
      [1,'Obra e responsável','Identifique onde o registro está sendo feito e quem responde por ele.'],
      [2,'Pacote e finalidade','Escolha a etapa, o pacote de serviço e defina se será somente FVS ou FVS com medição física.'],
      [3,'Execução física','Informe a frente/unidade e o percentual efetivamente executado no período.'],
      [4,'Inspeção de qualidade','Registre as evidências, marque as fotos e classifique a conformidade.'],
      [5,'Revisão e fechamento','Consolide o que foi aprovado, o que será medido e gere os documentos.']
    ];
    const stages = defs.map(d => makeStage(...d));
    // Mantém a ordem semântica 1 → 5 no DOM. Isso melhora leitura, acessibilidade e rolagem no celular.
    flow.after(...stages);

    // ETAPA 1 — move os campos existentes, preservando IDs e eventos.
    const fields = $('r_h_obra')?.closest('.fields');
    const b1 = stageEl(1).querySelector('.rf-stage-body');
    b1.insertAdjacentHTML('beforeend','<div class="rf-stage-note">Comece pela <b>obra</b> e pelo <b>responsável</b>. Localização e situação ficam como complementos da identificação.</div>');
    if(fields){
      const order = [
        $('r_h_obra')?.closest('.f'),
        $('r_h_resp')?.closest('.f'),
        $('r_st_and')?.closest('.f'),
        $('r_campo_abre'),
        $('r_campo_conc'),
        $('r_loc_on')?.closest('.f')
      ];
      const newFields = document.createElement('div');
      newFields.className = 'fields';
      order.filter(Boolean).forEach(x => newFields.appendChild(x));
      b1.appendChild(newFields);
      fields.remove();
    }

    // ETAPA 2 — pacote e modalidade.
    const b2 = stageEl(2).querySelector('.rf-stage-body');
    b2.insertAdjacentHTML('beforeend','<div class="rf-stage-note">Siga a ordem: <b>etapa → pacote de serviço → tipo de registro</b>. A FVS e a Medição usam exatamente a mesma base de pacotes do Apontamento.</div>');
    const pick = $('r_svc')?.closest('.pick');
    if(pick) b2.appendChild(pick);
    if($('r_mode_card')) b2.appendChild($('r_mode_card'));

    // ETAPA 3 — somente dados conhecidos antes da inspeção.
    const b3 = stageEl(3).querySelector('.rf-stage-body');
    b3.insertAdjacentHTML('beforeend','<div class="rf-stage-note">Antes da FVS, informe apenas o que já é conhecido: <b>frente/local</b> e <b>% executado</b>. O percentual anterior vem do histórico.</div>');
    const start = document.createElement('div');
    start.className = 'med-start-card';
    start.id = 'r_med_start';
    b3.appendChild(start);
    const panel = $('r_med_panel');
    const grid = panel?.querySelector('.med-input-grid');
    const local = panel?.querySelector('.med-local');
    if(local) start.appendChild(local);
    if(grid){
      const existing = Array.from(grid.children);
      const early = document.createElement('div');
      early.className = 'med-input-grid';
      existing.slice(0,2).forEach(x => early.appendChild(x));
      start.appendChild(early);
    }
    start.insertAdjacentHTML('beforeend',`
      <div class="med-presets">
        <button type="button" data-p="5">5%</button>
        <button type="button" data-p="10">10%</button>
        <button type="button" data-p="25">25%</button>
        <button type="button" data-p="50">50%</button>
        <button type="button" data-p="saldo">usar saldo</button>
      </div>
      <div class="med-exec-help">
        <div><b id="rf_exec_prev">0,00%</b><span>já medido anteriormente</span></div>
        <div><b id="rf_exec_after">0,00%</b><span>potencial após esta execução</span></div>
      </div>`);
    start.querySelectorAll('.med-presets button').forEach(btn => {
      btn.onclick = () => {
        const c = calc();
        const v = btn.dataset.p === 'saldo' ? Math.max(0,100-c.previous) : Number(btn.dataset.p);
        const inp = $('r_med_exec');
        if(inp){
          inp.value = v;
          inp.dispatchEvent(new Event('input',{bubbles:true}));
        }
      };
    });

    // ETAPA 4 — evidências + FVS.
    const b4 = stageEl(4).querySelector('.rf-stage-body');
    b4.insertAdjacentHTML('beforeend',`
      <div class="rf-inspection-progress">
        <div><b id="rf_inspection_title">Inspeção em andamento</b><small id="rf_inspection_sub">Adicione fotos e classifique os itens fotografados.</small></div>
        <strong id="rf_inspection_pct">0%</strong>
      </div>`);
    const sect = document.querySelector('#mod_rf .sect');
    const qs = $('r_quality_summary');
    const eng = $('rf_engine');
    const cap = $('r_capture');
    const groups = $('r_groups');
    [sect,qs,eng,cap,groups].filter(Boolean).forEach(x => b4.appendChild(x));

    // ETAPA 5 — fechamento + exportações.
    const b5 = stageEl(5).querySelector('.rf-stage-body');
    b5.insertAdjacentHTML('beforeend',`
      <div class="rf-stage-note">A inspeção vem antes do fechamento. Em medição física, confirme agora o percentual <b>aprovado</b> e quanto efetivamente será <b>medido</b>.</div>
      <div class="rf-final-summary" id="rf_final_summary"></div>
      <div class="rf-final-message" id="rf_final_message"></div>`);
    if(panel) b5.appendChild(panel);
    const closePresets=document.createElement('div');
    closePresets.className='med-close-presets';
    closePresets.id='rf_close_presets';
    closePresets.innerHTML=`<div><b>Atalhos de fechamento</b><small>Use quando toda a execução estiver aprovada e/ou será medida integralmente.</small></div><button type="button" id="rf_approve_exec">Aprovado = executado</button><button type="button" id="rf_measure_approved">Medido = aprovado</button>`;
    b5.appendChild(closePresets);
    $('rf_approve_exec').onclick=()=>{const cc=calc(),inp=$('r_med_appr');if(inp){inp.value=cc.executed;inp.dispatchEvent(new Event('input',{bubbles:true}));setTimeout(refresh,80)}};
    $('rf_measure_approved').onclick=()=>{const cc=calc(),inp=$('r_med_meas');if(inp){inp.value=cc.approved;inp.dispatchEvent(new Event('input',{bubbles:true}));setTimeout(refresh,80)}};
    [$('r_signoff'),$('r_legend')].filter(Boolean).forEach(x => b5.appendChild(x));
    b5.insertAdjacentHTML('beforeend',`
      <div class="rf-final-actions">
        <button type="button" class="pdf" id="rf_final_pdf">PDF / Compartilhar</button>
        <button type="button" class="xlsx" id="rf_final_xlsx">Excel completo + fotos</button>
        <button type="button" class="base" id="rf_final_base">Registrar na base / histórico</button>
      </div>`);
    $('rf_final_pdf').onclick = () => window.RF?.openMenu?.();
    $('rf_final_xlsx').onclick = () => M()?.exportExcel?.();
    $('rf_final_base').onclick = () => M()?.registerBase?.();

    nav(stageEl(1),0,2,'Identificação concluída →');
    nav(stageEl(2),1,3,'Continuar →');
    nav(stageEl(3),2,4,'Ir para inspeção →');
    nav(stageEl(4),2,5,'Concluir inspeção →');
    const nav5 = stageEl(5).querySelector('.rf-stage-nav');
    nav5.classList.add('one');
    const back5 = document.createElement('button');
    back5.type='button'; back5.textContent='← Voltar para inspeção'; back5.onclick=()=>go(4);
    nav5.appendChild(back5);

    const steps = $('rf_guide_steps');
    const labels = {1:'Obra',2:'Pacote',3:'Execução',4:'Inspeção',5:'Fechamento'};
    [1,2,3,4,5].forEach(n => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'rf-guide-step';
      b.dataset.step = n;
      b.innerHTML = `<i>${n}</i><span>${labels[n]}</span>`;
      b.onclick = () => go(n);
      steps.appendChild(b);
    });

    [$('r_mode_fvs'),$('r_mode_med')].filter(Boolean).forEach(btn => {
      btn.addEventListener('click',() => {
        modeChosen = true;
        setTimeout(() => {
          if(mode() === 'fvs' && current === 3) current = 4;
          refresh();
        },30);
      });
    });

    ['r_h_obra','r_h_obra2','r_h_resp','r_pkg_stage','r_svc','r_med_local','r_med_exec','r_med_appr','r_med_meas'].forEach(id => {
      $(id)?.addEventListener('input',() => setTimeout(refresh,60));
      $(id)?.addEventListener('change',() => setTimeout(refresh,80));
    });
    $('r_pkg_stage')?.addEventListener('change',() => {
      modeChosen = false;
      lastService = '';
      current = 2;
      setTimeout(refresh,120);
    });
    $('r_svc')?.addEventListener('change',() => {
      const now = $('r_svc').value;
      if(now !== lastService){
        lastService = now;
        modeChosen = false;
        current = 2;
      }
      setTimeout(refresh,520);
    });
    document.addEventListener('click',e => {
      if(e.target.closest?.('.quality-pick,.cam-bar,.annot-bottom,.mark-photo,.del')) setTimeout(refresh,180);
    });
    if($('r_groups')){
      const mo = new MutationObserver(() => refresh());
      mo.observe($('r_groups'),{childList:true,subtree:true,attributes:true,attributeFilter:['class']});
    }
    current = 1;
    refresh();
  }

  function reason(step){
    if(step === 1){
      const a=[];
      if(!obraOk()) a.push('selecione a obra');
      if(!respOk()) a.push('informe o responsável');
      return a.join(' e ');
    }
    if(step === 2){
      const a=[];
      if(!stageOk()) a.push('selecione a etapa');
      else if(!serviceOk()) a.push('selecione o pacote');
      if(!modeChosen) a.push('escolha o tipo de registro');
      return a.join(' e ');
    }
    if(step === 3) return 'informe a frente/local e o percentual executado';
    if(step === 4) return 'adicione ao menos uma foto e classifique todas as evidências da FVS';
    return '';
  }

  function go(n){
    if(!built) return;
    if(mode() === 'fvs' && n === 3) n = 4;
    const vis = visibleSteps();
    const idx = vis.indexOf(n);
    if(idx < 0) return;
    for(let i=0;i<idx;i++){
      const st = vis[i];
      if(!stepComplete(st)){
        current = st;
        refresh();
        window.RF?.toast?.('Antes de avançar, ' + reason(st) + '.',3200);
        stageEl(st)?.scrollIntoView({behavior:'smooth',block:'start'});
        return;
      }
    }
    current = n;
    refresh();
    stageEl(n)?.scrollIntoView({behavior:'smooth',block:'start'});
  }

  function markNextRequired(){
    document.querySelectorAll('.rf-next-required').forEach(x=>x.classList.remove('rf-next-required'));
    let el=null;
    if(current===1){
      if(!obraOk()) el=$('r_h_obra')?.closest('.f');
      else if(!respOk()) el=$('r_h_resp')?.closest('.f');
    }else if(current===2){
      if(!stageOk()) el=$('r_pkg_stage')?.closest('.rf-pkg-field') || $('r_pkg_stage')?.closest('.pick');
      else if(!serviceOk()) el=$('r_svc')?.closest('.rf-pkg-field') || $('r_svc')?.closest('.pick');
      else if(!modeChosen) el=$('r_mode_card');
    }else if(current===3 && mode()==='med'){
      const st=state();
      if(!String(st.local||'').trim()) el=$('r_med_local')?.closest('.med-local') || $('r_med_local')?.closest('.med-field');
      else if(!(Number(st.executed)>0)) el=$('r_med_exec')?.closest('.med-field');
    }else if(current===4){
      const q=quality();
      if(!q.photos) el=$('r_capture');
      else if(q.P>0) el=$('r_groups');
    }else if(current===5 && mode()==='med'){
      const cc=calc();
      if(cc.approved>cc.executed || cc.approved<0) el=$('r_med_appr')?.closest('.med-field');
      else if(cc.measured>cc.approved || cc.previous+cc.measured>100) el=$('r_med_meas')?.closest('.med-field');
    }
    el?.classList.add('rf-next-required');
  }

  function refresh(){
    if(!built) return;
    const md = mode();
    stageEl(3)?.classList.toggle('skip',md !== 'med');
    document.querySelector('.rf-guide-step[data-step="3"]')?.classList.toggle('skip',md !== 'med');
    if(md === 'fvs' && current === 3) current = 4;
    const vis = visibleSteps();

    vis.forEach(n => stageEl(n)?.classList.toggle('current',n === current));
    [1,2,3,4,5].forEach(n => {
      const b = document.querySelector('.rf-guide-step[data-step="'+n+'"]');
      if(!b) return;
      b.classList.toggle('current',n === current);
      b.classList.toggle('done',stepComplete(n));
      const ni = vis.indexOf(n);
      const prior = ni >= 0 ? vis.slice(0,ni) : [];
      b.classList.toggle('locked',prior.some(x => !stepComplete(x)));
    });

    let k=0;
    vis.forEach(n => {
      k++;
      const i = document.querySelector('.rf-guide-step[data-step="'+n+'"] i');
      if(i) i.textContent = stepComplete(n) ? '✓' : k;
      const sn = stageEl(n)?.querySelector('.rf-stage-num');
      if(sn) sn.textContent = k;
    });

    const s = snap();
    if($('rf_guide_sub')){
      const obra = obraOk() ? ((s?.header?.obraNome || $('r_h_obra2')?.value || $('r_h_obra')?.value) + ' · ') : '';
      const serv = serviceOk() ? ((s?.service?.code||'')+' · '+(s?.service?.name || 'Pacote selecionado')) : 'Selecione obra e pacote';
      const tipo = modeChosen ? (md === 'med' ? 'FVS + Medição Física' : 'Somente FVS') : 'defina o tipo de registro';
      $('rf_guide_sub').textContent = obra + serv + ' · ' + tipo;
    }

    const base = $('rf_guide_base');
    const db = $('r_med_db');
    if(base){
      const cloud = db?.classList.contains('cloud');
      base.innerHTML = `<b>${cloud?'M365':'LOCAL'}</b><small>${cloud?'SharePoint':'neste aparelho'}</small>`;
    }

    const c = calc();
    if($('rf_exec_prev')) $('rf_exec_prev').textContent = pct(c.previous);
    if($('rf_exec_after')) $('rf_exec_after').textContent = pct(Math.min(100,c.previous+c.executed));

    const q = quality();
    const done = q.C+q.NC+q.NA;
    const total = done+q.P;
    const ip = total ? Math.round(done/total*100) : 0;
    if($('rf_inspection_pct')) $('rf_inspection_pct').textContent = ip + '%';
    if($('rf_inspection_title')) $('rf_inspection_title').textContent = q.photos ? `${done} de ${total} evidências classificadas` : 'Nenhuma evidência ainda';
    if($('rf_inspection_sub')) $('rf_inspection_sub').textContent = q.photos ? `${q.photos} foto${q.photos!==1?'s':''} · ${q.C} conforme · ${q.NC} NC · ${q.NA} N/A` : 'Tire a primeira foto e vincule-a ao item verificado.';

    if(!modeChosen){
      $('r_mode_fvs')?.classList.remove('on');
      $('r_mode_med')?.classList.remove('on');
    }

    const fs = $('rf_final_summary');
    if(fs){
      const rows = md === 'med' ? [
        ['Anterior',pct(c.previous),''],['Executado',pct(c.executed),''],['Aprovado FVS',pct(c.approved),'good'],['Medido atual',pct(c.measured),'good'],
        ['Acumulado',pct(c.accumulated),'good'],['Saldo',pct(c.balance),''],['Pendente qualidade',pct(c.pendingQuality),c.pendingQuality?'warn':''],['Fotos',q.photos,'']
      ] : [
        ['Conformes',q.C,'good'],['Não conformes',q.NC,q.NC?'warn':''],['N/A',q.NA,''],['Fotos',q.photos,'']
      ];
      fs.innerHTML = rows.map(x => `<div class="rf-final-kpi ${x[2]}"><b>${x[1]}</b><span>${x[0]}</span></div>`).join('');
    }

    const fm = $('rf_final_message');
    if(fm){
      const errs = M()?._validate?.() || [];
      if(!inspectionOk()){
        fm.className = 'rf-final-message warn';
        fm.innerHTML = 'A inspeção ainda não está fechada. Volte e classifique as evidências pendentes.';
      }else if(md === 'med' && errs.length){
        fm.className = 'rf-final-message warn';
        fm.innerHTML = esc(errs.join(' · '));
      }else{
        fm.className = 'rf-final-message ok';
        fm.innerHTML = md === 'med' ? 'FVS concluída. Confirme os percentuais aprovado e medido, registre na base e exporte o Excel/PDF.' : 'FVS concluída. O registro está pronto para salvar na base e exportar.';
      }
    }

    const n1 = stageEl(1)?.querySelector('.next');
    if(n1) n1.disabled = !stepComplete(1);
    const n2 = stageEl(2)?.querySelector('.next');
    if(n2){
      n2.disabled = !stepComplete(2);
      n2.textContent = md === 'med' ? 'Ir para execução física →' : 'Ir para inspeção →';
      n2.onclick = () => go(md === 'med' ? 3 : 4);
    }
    const n3 = stageEl(3)?.querySelector('.next');
    if(n3) n3.disabled = !stepComplete(3);
    const n4 = stageEl(4)?.querySelector('.next');
    if(n4) n4.disabled = !stepComplete(4);

    const modeDisabled=!serviceOk();
    [$('r_mode_fvs'),$('r_mode_med')].filter(Boolean).forEach(b=>b.disabled=modeDisabled);
    if($('rf_close_presets')) $('rf_close_presets').style.display=md==='med'?'grid':'none';

    const activeIndex=Math.max(0,vis.indexOf(current));
    const completedBefore=vis.slice(0,activeIndex).filter(stepComplete).length;
    const fraction=(activeIndex+0.35)/(Math.max(1,vis.length));
    if($('rf_guide_progressline')) $('rf_guide_progressline').style.width=Math.min(100,Math.max(7,fraction*100))+'%';
    if($('rf_guide_counter')) $('rf_guide_counter').textContent='Etapa '+(activeIndex+1)+' de '+vis.length+(completedBefore?' · '+completedBefore+' concluída'+(completedBefore>1?'s':''):'');

    markNextRequired();
    if($('rf_final_base')) $('rf_final_base').disabled = !inspectionOk() || (md === 'med' && !closeValid());
    if($('rf_final_xlsx')) $('rf_final_xlsx').disabled = !inspectionOk();
  }

  window.ElevattaGuidedFlow = {go,refresh};
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded',() => setTimeout(build,80));
  else setTimeout(build,80);
})();

(function advancedMeasurementWorkbook(){
  function install(){
    const M = window.ElevattaMeasurement;
    if(!M?._getSnapshot) return setTimeout(install,150);

    M.exportExcel = async function(){
      const s = M._getSnapshot();
      const mode = M._getMode();
      const st = M._getState();
      const c = M._getCalc();
      const hist = M._getHistory();
      const frows = M._getFvsRows();
      if(!s?.service){ alert('Escolha um pacote de serviço antes de exportar.'); return; }
      if(typeof ExcelJS === 'undefined'){ alert('O gerador de planilhas ainda não terminou de carregar.'); return; }

      const q = s.quality || {};
      const photos = s.photos || [];
      const data = document.getElementById('r_h_data')?.value || new Date().toISOString().slice(0,10);
      const obra = s.header?.obraNome || document.getElementById('r_h_obra2')?.value || document.getElementById('r_h_obra')?.value || '';
      const local = st.local || '—';
      const resp = s.header?.resp || '—';
      const service = (s.service.etapa ? s.service.etapa + ' · ' : '') + (s.service.code || s.service.key) + ' · ' + s.service.name;

      const HX = h => ({argb:'FF'+h});
      const fill = h => ({type:'pattern',pattern:'solid',fgColor:HX(h)});
      const AR = (sz,b,col) => ({name:'Arial',size:sz,bold:!!b,color:HX(col||'000000')});
      const thin = {style:'thin',color:HX('C5CDD7')};
      const bd = {top:thin,left:thin,bottom:thin,right:thin};
      const NAVY='1F3864', ORANGE='E87722', GREEN='137333', RED='C5221F', AMBER='9A6700', GRAY='6B7280', LIGHT='E9EEF5', SOFT='F7F8FB', YELLOW='FFF2CC', WHITE='FFFFFF';
      const pval = v => (Number(v)||0)/100;
      const statusText = x => x.status==='C'?'CONFORME':x.status==='NC'?'NÃO CONFORME':x.status==='NA'?'N/A':'PENDENTE';
      const statusColor = x => x.status==='C'?GREEN:x.status==='NC'?RED:x.status==='NA'?GRAY:AMBER;

      try{
        window.ElevattaLoading?.show('Montando planilha profissional de medição…');
        const wb = new ExcelJS.Workbook();
        wb.creator = 'Elevatta Empreendimentos';
        wb.company = 'Elevatta Empreendimentos';
        wb.created = new Date();
        wb.modified = new Date();

        // 1 — FOLHA PRINCIPAL
        const mainName = mode === 'med' ? 'FOLHA DE MEDIÇÃO' : 'FOLHA FVS';
        const ws = wb.addWorksheet(mainName,{
          views:[{state:'frozen',ySplit:13,showGridLines:false}],
          pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,horizontalCentered:true,margins:{left:.28,right:.28,top:.42,bottom:.42,header:.18,footer:.18}}
        });
        [5,12,24,18,16,14,14,14,14,14,14,14,14,18,24].forEach((w,i)=>ws.getColumn(i+1).width=w);

        ws.mergeCells('A1:O1');
        let cell = ws.getCell('A1');
        cell.value = mode === 'med' ? 'FVS + MEDIÇÃO FÍSICA — FOLHA DE CONTROLE' : 'FVS — FOLHA DE CONTROLE DA QUALIDADE';
        cell.fill = fill(NAVY); cell.font = AR(14,true,WHITE); cell.alignment={horizontal:'center',vertical:'middle'}; ws.getRow(1).height=24;
        ws.mergeCells('A2:O2');
        cell = ws.getCell('A2');
        cell.value = mode === 'med' ? 'Controle físico percentual, inspeção FVS e evidências fotográficas. Documento sem valores financeiros.' : 'Inspeção FVS e evidências fotográficas do serviço.';
        cell.font=AR(8.5,false,GRAY); cell.alignment={horizontal:'center',vertical:'middle'}; ws.getRow(2).height=18;

        const meta = [['OBRA',obra],['SERVIÇO',service],['FRENTE / LOCAL',local],['RESPONSÁVEL',resp],['DATA',data]];
        meta.forEach((x,i)=>{
          const r=4+i;
          ws.mergeCells(r,1,r,2); ws.mergeCells(r,3,r,5);
          const a=ws.getCell(r,1), b=ws.getCell(r,3);
          a.value=x[0]; a.fill=fill(LIGHT); a.font=AR(8,true,NAVY); a.alignment={horizontal:'center',vertical:'middle'};
          b.value=x[1]; b.fill=fill(i<4?YELLOW:WHITE); b.font=AR(9,true,i===1?NAVY:'000000'); b.alignment={horizontal:'left',vertical:'middle',wrapText:true};
          for(let cc=1;cc<=5;cc++) ws.getCell(r,cc).border=bd;
          if(i===4){ try{ b.value=new Date(data+'T12:00:00'); b.numFmt='dd/mm/yyyy'; }catch(e){} }
          ws.getRow(r).height=i===1?24:18;
        });

        ws.mergeCells('F4:J4');
        cell=ws.getCell('F4'); cell.value=mode==='med'?'RESUMO DA MEDIÇÃO FÍSICA':'RESUMO DA FVS'; cell.fill=fill(NAVY); cell.font=AR(9,true,WHITE); cell.alignment={horizontal:'center',vertical:'middle'};
        if(mode==='med'){
          const labels=['ANTERIOR','EXECUTADO NO PERÍODO','APROVADO PELA FVS','MEDIDO NESTA ETAPA','ACUMULADO','SALDO FÍSICO'];
          const values=[c.previous,c.executed,c.approved,c.measured,c.accumulated,c.balance];
          labels.forEach((lab,i)=>{
            const r=5+i;
            ws.mergeCells(r,6,r,8); ws.mergeCells(r,9,r,10);
            const a=ws.getCell(r,6), b=ws.getCell(r,9);
            a.value=lab; a.fill=fill(LIGHT); a.font=AR(8,true,NAVY); a.alignment={vertical:'middle',indent:1};
            if(i===4) b.value={formula:'MIN(1,I5+I8)',result:pval(c.accumulated)};
            else if(i===5) b.value={formula:'MAX(0,1-I9)',result:pval(c.balance)};
            else b.value=pval(values[i]);
            b.numFmt='0.00%'; b.fill=fill(i===3?'FDEAD8':SOFT); b.font=AR(11,true,i===2||i===4?GREEN:i===3?ORANGE:NAVY); b.alignment={horizontal:'center',vertical:'middle'};
            for(let cc=6;cc<=10;cc++) ws.getCell(r,cc).border=bd;
          });
        }else{
          const items=[['CONFORMES',q.C||0,GREEN],['NÃO CONFORMES',q.NC||0,RED],['N/A',q.NA||0,GRAY],['PENDENTES',q.P||0,AMBER],['FOTOS',photos.length,NAVY]];
          items.forEach((x,i)=>{
            const r=5+i; ws.mergeCells(r,6,r,8); ws.mergeCells(r,9,r,10);
            ws.getCell(r,6).value=x[0]; ws.getCell(r,6).fill=fill(LIGHT); ws.getCell(r,6).font=AR(8,true,NAVY);
            ws.getCell(r,9).value=x[1]; ws.getCell(r,9).font=AR(11,true,x[2]); ws.getCell(r,9).alignment={horizontal:'center'};
            for(let cc=6;cc<=10;cc++) ws.getCell(r,cc).border=bd;
          });
        }

        ws.mergeCells('K4:O4');
        cell=ws.getCell('K4'); cell.value='QUALIDADE E EVIDÊNCIAS'; cell.fill=fill(NAVY); cell.font=AR(9,true,WHITE); cell.alignment={horizontal:'center'};
        const qm=[['CONFORMES',q.C||0,GREEN],['NÃO CONFORMES',q.NC||0,RED],['N/A',q.NA||0,GRAY],['PENDENTES',q.P||0,AMBER],['FOTOS',photos.length,NAVY],['ITENS FVS',frows.length,NAVY]];
        qm.forEach((x,i)=>{
          const r=5+i; ws.mergeCells(r,11,r,13); ws.mergeCells(r,14,r,15);
          const a=ws.getCell(r,11), b=ws.getCell(r,14);
          a.value=x[0]; a.fill=fill(LIGHT); a.font=AR(8,true,NAVY); a.alignment={vertical:'middle',indent:1};
          b.value=x[1]; b.fill=fill(SOFT); b.font=AR(11,true,x[2]); b.alignment={horizontal:'center',vertical:'middle'};
          for(let cc=11;cc<=15;cc++) ws.getCell(r,cc).border=bd;
        });

        ws.mergeCells('A12:O12');
        cell=ws.getCell('A12'); cell.value=mode==='med'?'MEDIÇÃO FÍSICA — HISTÓRICO E MEMÓRIA':'FVS — CONTROLE DA QUALIDADE'; cell.fill=fill(ORANGE); cell.font=AR(10,true,WHITE); cell.alignment={horizontal:'left',vertical:'middle',indent:1}; ws.getRow(12).height=19;

        let endRow=20;
        if(mode==='med'){
          const hdr=['Nº','DATA','LOCAL / FRENTE','ANTERIOR','EXECUTADO','APROVADO FVS','MEDIDO','ACUMULADO','SALDO','PEND. QUALIDADE','FOTOS','RESPONSÁVEL','OBSERVAÇÃO'];
          const merges=[[1,1],[2,2],[3,4],[5,5],[6,6],[7,7],[8,8],[9,9],[10,10],[11,11],[12,12],[13,13],[14,15]];
          hdr.forEach((h,i)=>{
            const [a,b]=merges[i]; if(b>a) ws.mergeCells(13,a,13,b);
            const cc=ws.getCell(13,a); cc.value=h; cc.fill=fill(NAVY); cc.font=AR(7.5,true,WHITE); cc.alignment={horizontal:'center',vertical:'middle',wrapText:true};
            for(let k=a;k<=b;k++) ws.getCell(13,k).border=bd;
          }); ws.getRow(13).height=28;

          const rows=hist.slice();
          const live={date:data,local:st.local||'',previous:c.previous,executed:c.executed,approved:c.approved,measured:c.measured,accumulated:c.accumulated,balance:c.balance,pendingQuality:c.pendingQuality,photoCount:photos.length,responsavel:resp,note:st.note||'',_live:true};
          if(c.executed||c.measured||!rows.length) rows.push(live);
          rows.forEach((x,i)=>{
            const r=14+i;
            const vals=[i+1,x.date||'',x.local||'',x.previous,x.executed,x.approved,x.measured,x.accumulated,x.balance,x.pendingQuality,x.photoCount||0,x.responsavel||resp,x.note||''];
            const map=[[1],[2],[3,4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14,15]];
            vals.forEach((v,j)=>{
              const cols=map[j], a=cols[0], b=cols[cols.length-1]; if(b>a) ws.mergeCells(r,a,r,b);
              const cc=ws.getCell(r,a);
              cc.value=(j>=3&&j<=9)?pval(v):v; if(j>=3&&j<=9) cc.numFmt='0.00%';
              cc.fill=fill(x._live?'FFF3ED':(i%2?SOFT:WHITE)); cc.font=AR(8,(j===7||x._live),j===7?GREEN:(x._live?ORANGE:'000000'));
              cc.alignment={horizontal:(j===0||j===1||(j>=3&&j<=10))?'center':'left',vertical:'middle',wrapText:j===2||j===11||j===12};
              for(let k=a;k<=b;k++) ws.getCell(r,k).border=bd;
            });
            ws.getRow(r).height=22;
          });

          let r=14+rows.length+2;
          ws.mergeCells(r,1,r,15); cell=ws.getCell(r,1); cell.value='FVS — ITENS VERIFICADOS NESTE REGISTRO'; cell.fill=fill(NAVY); cell.font=AR(9,true,WHITE); cell.alignment={horizontal:'left',indent:1}; r++;
          const fh=['ITEM','VERIFICAÇÃO','CRITÉRIO','RESULTADO','OBSERVAÇÃO','FOTOS'];
          const fm=[[1],[2,3,4,5,6],[7,8,9],[10,11],[12,13,14],[15]];
          fh.forEach((h,i)=>{
            const a=fm[i][0], b=fm[i][fm[i].length-1]; if(b>a) ws.mergeCells(r,a,r,b);
            const cc=ws.getCell(r,a); cc.value=h; cc.fill=fill(LIGHT); cc.font=AR(8,true,NAVY); cc.alignment={horizontal:'center',vertical:'middle',wrapText:true};
            for(let k=a;k<=b;k++) ws.getCell(r,k).border=bd;
          }); r++;
          frows.forEach((x,i)=>{
            const vals=[x.n,x.desc,x.criterion,statusText(x),x.obs,x.photos];
            const cols=[[1],[2,3,4,5,6],[7,8,9],[10,11],[12,13,14],[15]];
            vals.forEach((v,j)=>{
              const a=cols[j][0], b=cols[j][cols[j].length-1]; if(b>a) ws.mergeCells(r,a,r,b);
              const cc=ws.getCell(r,a); cc.value=v; cc.fill=fill(i%2?SOFT:WHITE); cc.font=AR(8,j===3,j===3?statusColor(x):'000000'); cc.alignment={horizontal:(j===0||j===3||j===5)?'center':'left',vertical:'top',wrapText:true};
              for(let k=a;k<=b;k++) ws.getCell(r,k).border=bd;
            }); ws.getRow(r).height=28; r++;
          });
          r+=2;
          ws.mergeCells(r,3,r,6); ws.mergeCells(r,10,r,13);
          const a=ws.getCell(r,3), b=ws.getCell(r,10); a.value=resp; b.value=new Date(data+'T12:00:00'); b.numFmt='dd/mm/yyyy';
          [a,b].forEach(cc=>{cc.font=AR(10,true,NAVY);cc.alignment={horizontal:'center',vertical:'bottom'};cc.border={bottom:{style:'medium',color:HX(NAVY)}}});
          r++; ws.mergeCells(r,3,r,6); ws.mergeCells(r,10,r,13); ws.getCell(r,3).value='Responsável / Inspetor'; ws.getCell(r,10).value='Data';
          [ws.getCell(r,3),ws.getCell(r,10)].forEach(cc=>{cc.font=AR(7.5,true,GRAY);cc.alignment={horizontal:'center'}});
          endRow=r;
        }else{
          const fh=['ITEM','VERIFICAÇÃO','CRITÉRIO','RESULTADO','OBSERVAÇÃO','FOTOS'];
          const fm=[[1],[2,3,4,5,6],[7,8,9],[10,11],[12,13,14],[15]];
          fh.forEach((h,i)=>{
            const a=fm[i][0], b=fm[i][fm[i].length-1]; if(b>a) ws.mergeCells(13,a,13,b);
            const cc=ws.getCell(13,a); cc.value=h; cc.fill=fill(NAVY); cc.font=AR(8,true,WHITE); cc.alignment={horizontal:'center',vertical:'middle',wrapText:true};
            for(let k=a;k<=b;k++) ws.getCell(13,k).border=bd;
          });
          let r=14;
          frows.forEach((x,i)=>{
            const vals=[x.n,x.desc,x.criterion,statusText(x),x.obs,x.photos];
            const cols=[[1],[2,3,4,5,6],[7,8,9],[10,11],[12,13,14],[15]];
            vals.forEach((v,j)=>{
              const a=cols[j][0], b=cols[j][cols[j].length-1]; if(b>a) ws.mergeCells(r,a,r,b);
              const cc=ws.getCell(r,a); cc.value=v; cc.fill=fill(i%2?SOFT:WHITE); cc.font=AR(8,j===3,j===3?statusColor(x):'000000'); cc.alignment={horizontal:(j===0||j===3||j===5)?'center':'left',vertical:'top',wrapText:true};
              for(let k=a;k<=b;k++) ws.getCell(r,k).border=bd;
            }); ws.getRow(r).height=28; r++;
          });
          endRow=Math.max(20,r);
        }
        ws.pageSetup.printArea='A1:O'+endRow;
        ws.headerFooter.oddFooter='&LElevatta Empreendimentos&CQualidade e Medição Física&R&P de &N';

        // 2 — FVS DETALHADA
        const wf=wb.addWorksheet('FVS DETALHADA',{views:[{state:'frozen',ySplit:5,showGridLines:false}],pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.3,right:.3,top:.4,bottom:.4}}});
        [7,42,28,17,40,10].forEach((w,i)=>wf.getColumn(i+1).width=w);
        wf.mergeCells('A1:F1'); wf.getCell('A1').value='FVS — '+service; wf.getCell('A1').fill=fill(NAVY); wf.getCell('A1').font=AR(13,true,WHITE); wf.getCell('A1').alignment={horizontal:'center'};
        wf.mergeCells('A2:F2'); wf.getCell('A2').value=obra+' · '+local+' · '+resp+' · '+data; wf.getCell('A2').font=AR(8.5,false,GRAY); wf.getCell('A2').alignment={horizontal:'center'};
        wf.mergeCells('A3:F3'); wf.getCell('A3').value='Legenda: CONFORME = verde · NÃO CONFORME = vermelho · N/A = cinza · PENDENTE = amarelo'; wf.getCell('A3').fill=fill(SOFT); wf.getCell('A3').font=AR(8,false,GRAY); wf.getCell('A3').alignment={horizontal:'center'};
        ['ITEM','VERIFICAÇÃO','CRITÉRIO','RESULTADO','OBSERVAÇÃO','FOTOS'].forEach((h,i)=>{const cc=wf.getCell(5,i+1);cc.value=h;cc.fill=fill(NAVY);cc.font=AR(8.5,true,WHITE);cc.alignment={horizontal:'center',vertical:'middle',wrapText:true};cc.border=bd;});
        wf.getRow(5).height=25;
        frows.forEach((x,i)=>{
          const r=6+i;
          [x.n,x.desc,x.criterion,statusText(x),x.obs,x.photos].forEach((v,j)=>{
            const cc=wf.getCell(r,j+1); cc.value=v; cc.border=bd; cc.fill=fill(i%2?SOFT:WHITE); cc.font=AR(9,j===3,j===3?statusColor(x):'000000'); cc.alignment={horizontal:(j===0||j===3||j===5)?'center':'left',vertical:'top',wrapText:true};
          });
          wf.getRow(r).height=34;
        });
        wf.autoFilter={from:'A5',to:'F'+Math.max(6,5+frows.length)};

        // 3 — MEMÓRIA DE MEDIÇÃO: mesma lógica de leitura e conferência da planilha de apontamentos.
        if(mode==='med'){
          const wm=wb.addWorksheet('MEMÓRIA DE MEDIÇÃO',{views:[{state:'frozen',ySplit:9,showGridLines:false}],pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,horizontalCentered:true,margins:{left:.28,right:.28,top:.4,bottom:.4,header:.18,footer:.18}}});
          [6,24,13,30,13,13,14,13,14,13,10,28,38].forEach((w,i)=>wm.getColumn(i+1).width=w);
          wm.mergeCells('A1:M1'); wm.getCell('A1').value='MEMÓRIA DE MEDIÇÃO FÍSICA — '+service; wm.getCell('A1').fill=fill(NAVY); wm.getCell('A1').font=AR(14,true,WHITE); wm.getCell('A1').alignment={horizontal:'center',vertical:'middle'}; wm.getRow(1).height=24;
          wm.mergeCells('A2:M2'); wm.getCell('A2').value='Histórico percentual do pacote. Sem valores financeiros. A linha laranja representa o registro atualmente em preenchimento.'; wm.getCell('A2').font=AR(8.5,false,GRAY); wm.getCell('A2').alignment={horizontal:'center'};
          wm.mergeCells('A4:C4'); wm.getCell('A4').value='OBRA'; wm.getCell('A4').fill=fill(LIGHT); wm.getCell('A4').font=AR(8,true,NAVY);
          wm.mergeCells('D4:G4'); wm.getCell('D4').value=obra; wm.getCell('D4').fill=fill(YELLOW); wm.getCell('D4').font=AR(9,true);
          wm.mergeCells('H4:I4'); wm.getCell('H4').value='RESPONSÁVEL'; wm.getCell('H4').fill=fill(LIGHT); wm.getCell('H4').font=AR(8,true,NAVY);
          wm.mergeCells('J4:M4'); wm.getCell('J4').value=resp; wm.getCell('J4').fill=fill(YELLOW); wm.getCell('J4').font=AR(9,true);
          for(let cc=1;cc<=13;cc++) wm.getCell(4,cc).border=bd;

          wm.mergeCells('A6:C6'); wm.getCell('A6').value='AVANÇO FÍSICO ACUMULADO'; wm.getCell('A6').fill=fill(NAVY); wm.getCell('A6').font=AR(9,true,WHITE); wm.getCell('A6').alignment={horizontal:'center'};
          wm.mergeCells('D6:E6'); wm.getCell('D6').value=pval(c.accumulated); wm.getCell('D6').numFmt='0.00%'; wm.getCell('D6').fill=fill('EAF4EE'); wm.getCell('D6').font=AR(13,true,GREEN); wm.getCell('D6').alignment={horizontal:'center'};
          wm.mergeCells('F6:H6'); wm.getCell('F6').value='SALDO FÍSICO'; wm.getCell('F6').fill=fill(NAVY); wm.getCell('F6').font=AR(9,true,WHITE); wm.getCell('F6').alignment={horizontal:'center'};
          wm.mergeCells('I6:J6'); wm.getCell('I6').value=pval(c.balance); wm.getCell('I6').numFmt='0.00%'; wm.getCell('I6').fill=fill(SOFT); wm.getCell('I6').font=AR(13,true,NAVY); wm.getCell('I6').alignment={horizontal:'center'};
          wm.mergeCells('K6:M6'); wm.getCell('K6').value=(c.pendingQuality>0?'PENDÊNCIA DE QUALIDADE: '+c.pendingQuality.toFixed(2)+'%':'SEM PENDÊNCIA DE QUALIDADE'); wm.getCell('K6').fill=fill(c.pendingQuality>0?'FFF4E5':'EAF4EE'); wm.getCell('K6').font=AR(8.5,true,c.pendingQuality>0?AMBER:GREEN); wm.getCell('K6').alignment={horizontal:'center',vertical:'middle',wrapText:true};
          for(let cc=1;cc<=13;cc++) wm.getCell(6,cc).border=bd;

          const mh=['Nº','ID','DATA','LOCAL / FRENTE','ANTERIOR','EXECUTADO','APROVADO FVS','MEDIDO','ACUMULADO','SALDO','FOTOS','RESPONSÁVEL','OBSERVAÇÃO'];
          mh.forEach((h,i)=>{const cc=wm.getCell(9,i+1);cc.value=h;cc.fill=fill(NAVY);cc.font=AR(8.2,true,WHITE);cc.alignment={horizontal:'center',vertical:'middle',wrapText:true};cc.border=bd;}); wm.getRow(9).height=27;
          const memoryRows=hist.map(x=>Object.assign({},x));
          memoryRows.push({id:'EM PREENCHIMENTO',date:data,local:st.local||'',previous:c.previous,executed:c.executed,approved:c.approved,measured:c.measured,accumulated:c.accumulated,balance:c.balance,photoCount:photos.length,responsavel:resp,note:st.note||'',_live:true});
          memoryRows.forEach((x,i)=>{
            const rr=10+i;
            const vv=[i+1,x.id||'',x.date||'',x.local||'',x.previous,x.executed,x.approved,x.measured,x.accumulated,x.balance,x.photoCount||0,x.responsavel||resp,x.note||''];
            vv.forEach((v,j)=>{const cc=wm.getCell(rr,j+1); if(j>=4&&j<=9){cc.value=pval(v);cc.numFmt='0.00%';}else cc.value=v; cc.border=bd;cc.fill=fill(x._live?'FFF3ED':(i%2?SOFT:WHITE));cc.font=AR(8.5,(j===8||x._live),j===8?GREEN:(x._live?ORANGE:'000000'));cc.alignment={horizontal:(j===0||j===2||(j>=4&&j<=10))?'center':'left',vertical:'middle',wrapText:j===3||j===11||j===12};});
            if(x._live){wm.getCell(rr,9).value={formula:'MIN(1,E'+rr+'+H'+rr+')',result:pval(c.accumulated)};wm.getCell(rr,10).value={formula:'MAX(0,1-I'+rr+')',result:pval(c.balance)};wm.getCell(rr,9).numFmt=wm.getCell(rr,10).numFmt='0.00%';}
            wm.getRow(rr).height=23;
          });
          const last=9+memoryRows.length;
          wm.autoFilter={from:'A9',to:'M'+last};
          wm.pageSetup.printArea='A1:M'+last;
          wm.headerFooter.oddFooter='&LElevatta Empreendimentos&CMemória de Medição Física&R&P de &N';
        }

        // 4 — REGISTRO FOTOGRÁFICO
        const wp=wb.addWorksheet('REGISTRO FOTOGRÁFICO',{views:[{showGridLines:false}],pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.35,right:.35,top:.4,bottom:.4}}});
        [4,18,18,18,18,18,18,4].forEach((w,i)=>wp.getColumn(i+1).width=w);
        wp.mergeCells('A1:H1'); wp.getCell('A1').value='REGISTRO FOTOGRÁFICO — '+service; wp.getCell('A1').fill=fill(NAVY); wp.getCell('A1').font=AR(13,true,WHITE); wp.getCell('A1').alignment={horizontal:'center'};
        wp.mergeCells('A2:H2'); wp.getCell('A2').value=obra+' · '+local+' · '+data; wp.getCell('A2').font=AR(8.5,false,GRAY); wp.getCell('A2').alignment={horizontal:'center'};
        let pr=4;
        for(let i=0;i<photos.length;i++){
          const p=photos[i];
          wp.mergeCells(pr,2,pr,7);
          const hh=wp.getCell(pr,2); hh.value='FOTO '+String(i+1).padStart(2,'0')+' · '+(p.dt||data); hh.fill=fill(i%2?ORANGE:NAVY); hh.font=AR(9,true,WHITE); hh.alignment={horizontal:'left',indent:1};
          pr++;
          try{
            const img=await M._markedPhoto(p,1100);
            const id=wb.addImage({base64:img.data,extension:'jpeg'});
            const ratio=img.w/img.h, w=420, h=Math.min(300,420/Math.max(.7,ratio));
            wp.addImage(id,{tl:{col:1.05,row:pr-1+.1},ext:{width:w,height:h},editAs:'oneCell'});
            const rows=Math.ceil(h/20);
            for(let rr=0;rr<rows;rr++) wp.getRow(pr+rr).height=15;
            const infoEnd=pr+Math.max(4,Math.min(rows-1,8));
            wp.mergeCells(pr,6,infoEnd,7);
            const info=wp.getCell(pr,6);
            const vinculo=p.g==='G'?'Registro geral':'Item FVS '+p.g;
            info.value='VÍNCULO\n'+vinculo+'\n\nLEGENDA\n'+(p.cap||'Sem legenda')+'\n\nMARCAÇÕES\n'+(p.marcas||0)+' elemento(s)';
            info.fill=fill(SOFT); info.font=AR(8,false,NAVY); info.alignment={vertical:'top',wrapText:true}; info.border=bd;
            pr+=rows+1;
          }catch(e){
            wp.mergeCells(pr,2,pr+2,7); wp.getCell(pr,2).value='Não foi possível incorporar esta foto.'; pr+=4;
          }
          pr++;
        }

        // 5 — BASE NORMALIZADA / importável
        const wbse=wb.addWorksheet('BASE MEDIÇÕES',{views:[{state:'frozen',ySplit:2,showGridLines:false}]});
        const bh=['ID','DATA','OBRA','CÓD. PACOTE','PACOTE','LOCAL / FRENTE','ANTERIOR','EXECUTADO','APROVADO FVS','MEDIDO','ACUMULADO','SALDO','PEND. QUALIDADE','FOTOS','RESPONSÁVEL','OBSERVAÇÃO'];
        [24,12,22,14,36,28,12,12,14,12,12,12,14,9,24,40].forEach((w,i)=>wbse.getColumn(i+1).width=w);
        wbse.mergeCells('A1:P1'); wbse.getCell('A1').value='BASE DE MEDIÇÕES FÍSICAS — LINHAS NORMALIZADAS'; wbse.getCell('A1').fill=fill(NAVY); wbse.getCell('A1').font=AR(10,true,WHITE); wbse.getCell('A1').alignment={horizontal:'center'};
        bh.forEach((h,i)=>{const cc=wbse.getCell(2,i+1);cc.value=h;cc.fill=fill(LIGHT);cc.font=AR(8,true,NAVY);cc.border=bd;cc.alignment={horizontal:'center',vertical:'middle',wrapText:true};});
        const baseRows=hist.map(x=>Object.assign({},x));
        if(mode==='med' && (c.executed||c.measured||!baseRows.length)) baseRows.push({id:'EM PREENCHIMENTO',date:data,obra,serviceCode:s.service.code||s.service.key,serviceName:s.service.name,local:st.local,previous:c.previous,executed:c.executed,approved:c.approved,measured:c.measured,accumulated:c.accumulated,balance:c.balance,pendingQuality:c.pendingQuality,photoCount:photos.length,responsavel:resp,note:st.note});
        baseRows.forEach((x,i)=>{
          const vals=[x.id,x.date,x.obra||obra,x.serviceCode||s.service.code||s.service.key,x.serviceName||s.service.name,x.local,x.previous,x.executed,x.approved,x.measured,x.accumulated,x.balance,x.pendingQuality,x.photoCount||0,x.responsavel||resp,x.note||''];
          vals.forEach((v,j)=>{
            const cc=wbse.getCell(3+i,j+1);
            cc.value=(j>=6&&j<=12)?pval(v):v; if(j>=6&&j<=12) cc.numFmt='0.00%';
            cc.border=bd; cc.fill=fill(i%2?SOFT:WHITE); cc.font=AR(8); cc.alignment={vertical:'middle',wrapText:j===4||j===5||j===14||j===15};
          });
        });
        if(baseRows.length) wbse.autoFilter={from:'A2',to:'P'+(2+baseRows.length)};

        const buf=await wb.xlsx.writeBuffer();
        const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        const safe=x=>String(x||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9_-]+/g,'_').replace(/^_+|_+$/g,'');
        const name=safe(obra)+'_'+safe(s.service.code||s.service.key)+'_'+(mode==='med'?'MEDICAO_FISICA':'FVS')+'_'+data+'.xlsx';
        await window.RF.saveFile(blob,name);
      }catch(e){
        console.error(e);
        alert('Não foi possível gerar a planilha: '+(e?.message||e));
      }finally{
        window.ElevattaLoading?.hide(true);
      }
    };
  }
  install();
})();
