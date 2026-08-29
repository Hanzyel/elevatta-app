/* Elevatta REV36 — compatibility hooks and Microsoft 365 integration */
(function(){
  'use strict';
  const $=id=>document.getElementById(id);
  function boot(){
    // Replace guided-flow base action/status with Microsoft 365.
    const b=$('rf_guide_base');if(b){b.onclick=()=>window.ElevattaM365?.open('records');}
    // Expose a safe reload hook to refresh the previous measurement after pulling SharePoint history.
    if(window.ElevattaMeasurement&&!ElevattaMeasurement._reload){
      // The measurement module in REV33 publishes this hook in index.html; keep fallback harmless.
      ElevattaMeasurement._reload=()=>{try{ElevattaMeasurement.render?.()}catch(e){}};
    }
    // Final AP export remains local-first; SharePoint sync runs without blocking PDF/Excel sharing.
    if(window.AP&&!AP.__m365SavePatched){
      AP.__m365SavePatched=true;
      const original=AP.salvar;
      AP.salvar=async function(){const r=await original.apply(this,arguments);setTimeout(()=>window.ElevattaM365?.syncApontamento({silent:true}).catch(()=>{}),60);return r};
      const trocar=AP.trocarDia;
      AP.trocarDia=async function(){const date=$('a_data')?.value;if(date&&window.ElevattaM365?.isConfigured?.()&&!localStorage.getItem('ap_dia_'+date)){try{await ElevattaM365.pullApontamentoDate(date,{silent:true})}catch(e){}}return trocar.apply(this,arguments)};
    }
    // Pull the latest accumulated measurement when package/obra changes, but never block the field flow.
    let tm=0;const auto=()=>{clearTimeout(tm);tm=setTimeout(()=>window.ElevattaM365?.syncCurrentMeasurementHistory({silent:true}).catch(()=>{}),650)};
    ['r_svc','r_h_obra'].forEach(id=>$(id)?.addEventListener('change',auto));$('r_h_obra2')?.addEventListener('change',auto);
    // Update guided status after M365 module has initialized.
    setInterval(()=>{const base=$('rf_guide_base'),badge=$('m365_home_badge');if(base){const connected=badge?.textContent==='conectado';base.innerHTML=`<b>${connected?'M365':'BASE'}</b><small>${connected?'SharePoint':'local'}</small>`;}},2500);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
