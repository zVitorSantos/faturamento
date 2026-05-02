// ══════════════════════════════════════════════════════════
//  ESTADO GLOBAL
// ══════════════════════════════════════════════════════════
const STATE={rootDir:null,cache:{},periods:[],curIdx:0,filiais:[],history:{},projection:null,historyDir:null};
const $=id=>document.getElementById(id);
const $$=s=>document.querySelectorAll(s);

// Persistência de pasta usando IndexedDB
async function saveRootDirHandle(handle){
  if(!('indexedDB'in window))return;
  try{
    const db=await idbOpen('freight_db',1);
    const tx=db.transaction('handles','readwrite');
    const store=tx.objectStore('handles');
    await store.put(handle,'rootDir');
    await tx.done;
  }catch(e){console.error('Erro ao salvar handle:',e);}
}

async function loadRootDirHandle(){
  if(!('indexedDB'in window))return null;
  try{
    const db=await idbOpen('freight_db',1);
    const tx=db.transaction('handles','readonly');
    const store=tx.objectStore('handles');
    const handle=await store.get('rootDir');
    await tx.done;
    if(handle){
      const permission=await handle.queryPermission({mode:'read'});
      if(permission==='granted')return handle;
      if(permission==='prompt'){
        const request=await handle.requestPermission({mode:'read'});
        if(request==='granted')return handle;
      }
    }
  }catch(e){console.error('Erro ao carregar handle:',e);}
  return null;
}

async function idbOpen(name,version){
  return new Promise((resolve,reject)=>{
    const request=indexedDB.open(name,version);
    request.onupgradeneeded=e=>{
      const db=e.target.result;
      if(!db.objectStoreNames.contains('handles'))db.createObjectStore('handles');
    };
    request.onsuccess=()=>resolve(request.result);
    request.onerror=()=>reject(request.error);
  });
}

function toggleTemplatePanel(){const p=$('template-panel'),l=$('template-toggle-lbl'),o=p.style.display==='none';p.style.display=o?'block':'none';l.textContent=o?'▲ recolher':'▼ expandir';}

// ══════════════════════════════════════════════════════════
//  DIAS ÚTEIS
// ══════════════════════════════════════════════════════════
const FIXED_HOL=['01-01','04-21','05-01','09-07','10-12','11-02','11-15','11-20','12-25'];
const VAR_HOL={2024:['2024-02-12','2024-02-13','2024-03-29'],2025:['2025-03-03','2025-03-04','2025-04-18'],2026:['2026-02-16','2026-02-17','2026-04-03'],2027:['2027-03-01','2027-03-02','2027-03-26']};
function workingDays(year,month){const days=[],d=new Date(year,month-1,1),varH=VAR_HOL[year]||[];while(d.getMonth()===month-1){const dow=d.getDay();if(dow>0&&dow<6){const iso=d.toISOString().split('T')[0];if(!FIXED_HOL.includes(iso.slice(5))&&!varH.includes(iso))days.push(iso);}d.setDate(d.getDate()+1);}return days;}

const MONTHS=['','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const DN=['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
const dn=d=>DN[new Date(d+'T12:00:00').getDay()];
const fd=d=>new Date(d+'T12:00:00').toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit'});
const R$=v=>isFinite(v)?'R$ '+Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}):'—';
const pct=v=>isFinite(v)?(v>=0?'+':'')+v.toFixed(1)+'%':'—';
const bcl=v=>v>=0?'up':'dn';
const padM=m=>String(m).padStart(2,'0');
function decodePeriod(p){const[y,m]=p.split('-');return{year:parseInt(y),month:parseInt(m)};}

function projection(totCur,wdCur){const today=new Date().toISOString().split('T')[0];const elapsed=wdCur.filter(d=>d<=today).length;if(elapsed<=0)return 0;return(totCur/elapsed)*wdCur.length;}
function proRata(totCur,totPrev,wdCur,wdPrev){const today=new Date().toISOString().split('T')[0];const elapsed=wdCur.filter(d=>d<=today).length;if(elapsed<=0||wdPrev.length<=0)return 0;const pr=(totPrev/wdPrev.length)*elapsed;return pr>0?((totCur-pr)/pr*100):0;}

// ══════════════════════════════════════════════════════════
//  CLASSIFY FILE — suporte a SC
// ══════════════════════════════════════════════════════════
function classifyFile(name){
  const n=name.toLowerCase();
  if(!n.endsWith('.xls')&&!n.endsWith('.xlsx'))return null;
  // FILIAL-SC.xls  →  source: 'sc'
  const mSC=n.match(/^([a-z0-9]+)-sc\.xlsx?$/i);
  if(mSC)return{filial:mSC[1].toUpperCase(),source:'sc'};
  const m=n.match(/^([a-z0-9]+)-(carteira|extra)\.xlsx?$/i);
  return m?{filial:m[1].toUpperCase(),source:m[2].toLowerCase()}:null;
}

// ══════════════════════════════════════════════════════════
//  FILIAIS DINÂMICAS
// ══════════════════════════════════════════════════════════
function discoverFiliaisFromData(){
  const set=new Set();
  Object.values(STATE.cache).forEach(d=>{[...d.cur,...d.prev,...(d.sc||[])].forEach(a=>{if(a.filial)set.add(a.filial);});});
  STATE.filiais=[...set].sort();
}

function buildDynamicTabs(){
  const container=$('filial-tabs');if(!container)return;
  container.innerHTML='';
  STATE.filiais.forEach(f=>{
    const div=document.createElement('div');
    div.className='tab';div.id=`tab-${f.toLowerCase()}`;div.textContent=f;
    div.onclick=function(){showPane('filial_'+f,this);};
    container.appendChild(div);
  });
}

function buildDynamicPanes(){
  const container=$('filial-panes-container');if(!container)return;
  container.innerHTML='';
  STATE.filiais.forEach(f=>{
    const pfx=f.toLowerCase();
    const div=document.createElement('div');
    div.className='pane filial-pane';div.id=`pane-filial_${f}`;
    div.innerHTML=`
      <div class="sec-hdr" style="flex-wrap:wrap;gap:10px">
        <div class="sec-title"><span class="dot"></span> <span id="title-${pfx}">Filial ${f}</span> — <span id="lbl-${pfx}">—</span></div>
        <div class="filial-fresh" id="fresh-${pfx}"></div>
      </div>
      <div id="${pfx}-nodata" class="empty" style="display:none">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
        <div style="margin-top:8px">Nenhum arquivo ${f} para este período.</div>
      </div>
      <div id="${pfx}-data">
        <div class="kpi-grid">
          <div class="kpi a"><div class="kpi-lbl">Faturado</div><div class="kpi-val" id="${pfx}-tot">—</div><div class="kpi-sub"><span class="badge" id="${pfx}-badge"></span></div></div>
          <div class="kpi b"><div class="kpi-lbl">Carteira</div><div class="kpi-val" id="${pfx}-cart">—</div><div class="kpi-sub" id="${pfx}-cart-sub"></div></div>
          <div class="kpi c"><div class="kpi-lbl">Extra</div><div class="kpi-val" id="${pfx}-extra">—</div><div class="kpi-sub" id="${pfx}-extra-sub"></div></div>
          <div class="kpi sc"><div class="kpi-lbl">SC Potencial</div><div class="kpi-val" id="${pfx}-sc-kpi" style="color:var(--sc)">—</div><div class="kpi-sub" id="${pfx}-sc-kpi-sub"></div></div>
          <div class="kpi d"><div class="kpi-lbl">Média Diária</div><div class="kpi-val" id="${pfx}-avg">—</div><div class="kpi-sub" id="${pfx}-avg-sub"></div></div>
          <div class="kpi a"><div class="kpi-lbl">Projeção Mês</div><div class="kpi-val" id="${pfx}-proj">—</div><div class="kpi-sub"><span class="badge" id="${pfx}-proj-b"></span></div></div>
          <div class="kpi b"><div class="kpi-lbl">AWBs</div><div class="kpi-val" id="${pfx}-awbs">—</div><div class="kpi-sub" id="${pfx}-awbs-sub"></div></div>
        </div>
        <div class="prog-wrap">
          <div class="prog-hdr"><span>Progresso vs total <span class="lbl-prev"></span></span><span style="font-family:var(--mono);color:var(--text2)" id="${pfx}-prog-pct">—</span></div>
          <div class="prog-track"><div class="prog-bar" id="${pfx}-prog-bar" style="width:0%"></div></div>
          <div style="display:flex;justify-content:space-between;margin-top:6px;font-size:11px;color:var(--text3);font-family:var(--mono)"><span>R$ 0</span><span id="${pfx}-prog-meta"></span></div>
        </div>
        <div class="charts2">
          <div class="chart-card">
            <div class="chart-ttl">Diário ${f} <div class="chart-leg"><span style="color:var(--accent)">■ atual</span><span style="color:var(--cyan)">— ant.</span><span style="color:var(--sc)">▪ SC</span></div></div>
            <canvas id="ch-${pfx}-daily" height="130"></canvas>
          </div>
          <div style="display:grid;grid-template-rows:1fr 1fr;gap:16px">
            <div class="chart-card"><div class="chart-ttl">Tipo de Frete (CIF / FOB)</div><div style="height:150px"><canvas id="ch-${pfx}-tipo"></canvas></div></div>
            <div class="chart-card"><div class="chart-ttl">Modal de Transporte</div><div style="height:150px"><canvas id="ch-${pfx}-modal"></canvas></div></div>
          </div>
        </div>
        <div class="tbl-card">
          <div class="tbl-top">
            <div class="tbl-top-ttl">Faturamento por Dia — ${f}</div>
            <div style="display:flex;gap:8px">
              <button class="btn btn-s" onclick="exportCSV('${f}')">Exportar CSV</button>
              <button class="btn btn-s" onclick="exportSCCSV('${f}')" style="color:var(--sc);border-color:var(--sc-border)">Exportar SC</button>
            </div>
          </div>
          <div class="tbl-wrap">
            <table id="tbl-${pfx}">
              <thead><tr>
                <th onclick="srt('tbl-${pfx}',0)">Data</th><th>Dia</th>
                <th onclick="srt('tbl-${pfx}',2)">Carteira</th>
                <th onclick="srt('tbl-${pfx}',3)">Extra</th>
                <th onclick="srt('tbl-${pfx}',4)" style="color:var(--sc)">SC</th>
                <th onclick="srt('tbl-${pfx}',5)">Total Comiss.</th>
                <th onclick="srt('tbl-${pfx}',6)">AWBs</th>
                <th onclick="srt('tbl-${pfx}',7)">Volumes</th>
                <th onclick="srt('tbl-${pfx}',8)">Ant. equiv.</th>
                <th onclick="srt('tbl-${pfx}',9)">Δ%</th>
                <th title="Clientes novos/retorno">Clientes</th>
                <th title="SC do dia" style="color:var(--sc)">SC↕</th>
              </tr></thead>
              <tbody id="body-${pfx}"></tbody>
            </table>
          </div>
        </div>
        <div class="chart-card" style="margin-bottom:20px">
          <div class="chart-ttl">Top Clientes — ${f}</div>
          <div style="height:200px"><canvas id="ch-${pfx}-clientes"></canvas></div>
        </div>
      </div>`;
    container.appendChild(div);
  });

  // Atualiza filtro de filial na busca AWB
  const afEl=$('af');
  if(afEl){afEl.innerHTML='<option value="">Todas</option>';STATE.filiais.forEach(f=>{const opt=document.createElement('option');opt.value=f;opt.textContent=f;afEl.appendChild(opt);});}
  rebuildOverviewFilialKPIs();
}

function rebuildOverviewFilialKPIs(){
  const kpiGrid=$('ov-filial-kpis');if(!kpiGrid)return;
  kpiGrid.innerHTML='';
  const colors=['a','e','b','c','d'];
  STATE.filiais.forEach((f,i)=>{
    const c=colors[i%colors.length];
    kpiGrid.innerHTML+=`
      <div class="kpi ${c}">
        <div class="kpi-lbl">${f} <span class="lbl-cur"></span></div>
        <div class="kpi-val" id="ov-${f.toLowerCase()}">—</div>
        <div class="kpi-sub" id="ov-${f.toLowerCase()}-sub"></div>
      </div>`;
  });
}

function applyBranchLabels(){if($('hdr-filiais-sub')&&STATE.filiais.length>0)$('hdr-filiais-sub').textContent=STATE.filiais.join(' · ');}

// ══════════════════════════════════════════════════════════
//  FRESHNESS — última atualização por filial/source × esperado
// ══════════════════════════════════════════════════════════
function lastWorkingDayOnOrBefore(iso){
  // assume iso = 'YYYY-MM-DD', retorna o último dia útil <= iso
  const d=new Date(iso+'T12:00:00');
  while(true){
    const dow=d.getDay();
    const isoCur=d.toISOString().slice(0,10);
    const monthDay=isoCur.slice(5);
    const varH=VAR_HOL[d.getFullYear()]||[];
    if(dow>0&&dow<6&&!FIXED_HOL.includes(monthDay)&&!varH.includes(isoCur))return isoCur;
    d.setDate(d.getDate()-1);
  }
}

function diffWorkingDays(fromIso,toIso){
  if(!fromIso||!toIso||fromIso>=toIso)return 0;
  const out=[];const d=new Date(fromIso+'T12:00:00');d.setDate(d.getDate()+1);
  const end=new Date(toIso+'T12:00:00');
  while(d<=end){
    const dow=d.getDay();const iso=d.toISOString().slice(0,10);const monthDay=iso.slice(5);
    const varH=VAR_HOL[d.getFullYear()]||[];
    if(dow>0&&dow<6&&!FIXED_HOL.includes(monthDay)&&!varH.includes(iso))out.push(iso);
    d.setDate(d.getDate()+1);
  }
  return out.length;
}

function computeFreshnessStatus(){
  const today=new Date().toISOString().slice(0,10);
  const yesterday=(()=>{const d=new Date();d.setDate(d.getDate()-1);return d.toISOString().slice(0,10);})();
  const expected=lastWorkingDayOnOrBefore(yesterday);
  const curPeriod=STATE.periods[STATE.periods.length-1];
  const data=curPeriod?STATE.cache[curPeriod]:null;
  const out={expected,today,byFilial:{}};
  STATE.filiais.forEach(f=>{
    const sources={carteira:null,extra:null,sc:null};
    if(data){
      data.cur.filter(a=>a.filial===f).forEach(a=>{if(a.data<=today&&(!sources[a.source]||a.data>sources[a.source]))sources[a.source]=a.data;});
      (data.sc||[]).filter(a=>a.filial===f).forEach(a=>{if(a.data<=today&&(!sources.sc||a.data>sources.sc))sources.sc=a.data;});
    }
    const sourcesPresent=Object.entries(sources).filter(([,v])=>v).map(([k])=>k);
    const hasSC=filialHasSC(f);
    // worst-case: source mais defasado
    const reportSources=['carteira','extra'];if(hasSC)reportSources.push('sc');
    let worstGap=0,worstSource=null;
    reportSources.forEach(s=>{
      if(!sources[s]){worstGap=Math.max(worstGap,99);worstSource=worstSource||s;return;}
      const gap=diffWorkingDays(sources[s],expected);
      if(gap>worstGap){worstGap=gap;worstSource=s;}
    });
    let status='ok';
    if(worstGap===0)status='ok';
    else if(worstGap===1)status='warn';
    else status='late';
    out.byFilial[f]={sources,worstGap,worstSource,status,sourcesPresent,hasSC};
  });
  STATE.freshness=out;
  return out;
}

const ICONS={
  check:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>',
  warn:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
  alert:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
  refresh:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>',
  folder:'<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
  users:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
  trend:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>',
  calendar:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
  cash:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>',
  bolt:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>',
  chart:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>',
  search:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
  download:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>',
  bell:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>',
  trophy:'<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6"/><path d="M18 9h1.5a2.5 2.5 0 0 0 0-5H18"/><path d="M4 22h16"/><path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22"/><path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22"/><path d="M18 2H6v7a6 6 0 0 0 12 0V2Z"/></svg>'
};

function freshnessChipHTML(f,info){
  const colorMap={ok:'var(--success)',warn:'var(--warning)',late:'var(--danger)'};
  const bgMap={ok:'rgba(16,185,129,.12)',warn:'rgba(245,158,11,.14)',late:'rgba(239,68,68,.14)'};
  const ic=info.status==='ok'?ICONS.check:info.status==='warn'?ICONS.warn:ICONS.alert;
  const expectedFmt=info.sources.carteira||info.sources.extra?fd(STATE.freshness.expected):'—';
  let txt;
  if(info.status==='ok')txt=`atualizado até ${expectedFmt}`;
  else if(info.worstSource&&!info.sources[info.worstSource])txt=`${info.worstSource} sem dados`;
  else txt=`${info.worstSource} ${info.worstGap}d atrás`;
  return `<span class="fresh-chip" style="background:${bgMap[info.status]};color:${colorMap[info.status]};border:1px solid ${colorMap[info.status]}33"><span style="display:inline-flex;align-items:center">${ic}</span><strong style="margin-left:5px">${f}</strong><span style="margin-left:6px;font-weight:600">${txt}</span></span>`;
}

function renderFreshnessBanner(){
  if(!STATE.filiais.length||!STATE.freshness)return;
  let banner=$('freshness-banner');
  if(!banner){
    banner=document.createElement('div');
    banner.id='freshness-banner';
    banner.className='fresh-banner';
    const content=document.querySelector('.content');
    if(content)content.insertBefore(banner,content.firstChild);
  }
  const expected=STATE.freshness.expected;
  const list=STATE.filiais.map(f=>freshnessChipHTML(f,STATE.freshness.byFilial[f]||{status:'late',worstGap:99})).join('');
  const allOk=STATE.filiais.every(f=>STATE.freshness.byFilial[f]?.status==='ok');
  const headerColor=allOk?'var(--success)':'var(--warning)';
  banner.innerHTML=`
    <div class="fresh-banner-hdr">
      <div class="fresh-banner-ttl">
        <span style="color:${headerColor}">${allOk?ICONS.check:ICONS.warn}</span>
        <span>Status dos arquivos</span>
        <span class="fresh-banner-sub">esperado: ${fd(expected)} (último dia útil)</span>
      </div>
      <button class="fresh-refresh-btn" onclick="checkAndReloadIfChanged(false)" title="Recarregar">
        ${ICONS.refresh}<span>Atualizar</span>
      </button>
    </div>
    <div class="fresh-chips">${list}</div>
  `;
}

document.addEventListener('DOMContentLoaded',()=>{
  const y=new Date().getFullYear();
  if($('cfg-cur-year')&&!$('cfg-cur-year').value)$('cfg-cur-year').value=y;
  if($('cfg-prev-year')&&!$('cfg-prev-year').value)$('cfg-prev-year').value=y-1;
  loadNegData();
  tryAutoOpenFolder();
});

// ══════════════════════════════════════════════════════════
//  GERAR PASTA MODELO
// ══════════════════════════════════════════════════════════
async function generateTemplateStructure(){
  if(!window.showDirectoryPicker){toast('Use Chrome ou Edge.','err');return;}
  try{
    const curYear=parseInt($('cfg-cur-year')?.value||new Date().getFullYear(),10)||new Date().getFullYear();
    const prevYear=parseInt($('cfg-prev-year')?.value||(curYear-1),10)||(curYear-1);
    const exemplar=($('cfg-example-filial')?.value||'ABC').trim().toUpperCase()||'ABC';
    const base=await window.showDirectoryPicker({mode:'readwrite'});
    async function writeReadme(d,fn,c){const fh=await d.getFileHandle(fn,{create:true});const w=await fh.createWritable();await w.write(c);await w.close();}
    async function createAnnualYear(year){const yd=await base.getDirectoryHandle(String(year),{create:true});const fd=await yd.getDirectoryHandle(exemplar,{create:true});await writeReadme(fd,'LEIA-ME.txt',`Coloque aqui os arquivos XLS da filial ${exemplar} para o ano ${year}.\nFormato: ${exemplar}-Carteira.xls, ${exemplar}-Extra.xls`);}
    await createAnnualYear(prevYear);
    for(const hYear of[prevYear-1,prevYear-2].filter(y=>y>=2020))await createAnnualYear(hYear);
    const curDir=await base.getDirectoryHandle(String(curYear),{create:true});
    for(let m=1;m<=12;m++){const md=await curDir.getDirectoryHandle(padM(m),{create:true});const fd=await md.getDirectoryHandle(exemplar,{create:true});await writeReadme(fd,'LEIA-ME.txt',`Coloque aqui os XLS da filial ${exemplar} — ${padM(m)}/${curYear}.\nArquivos: ${exemplar}-Carteira.xls, ${exemplar}-Extra.xls, ${exemplar}-SC.xls`);}
    toast('Pasta modelo criada com sucesso!','ok');
  }catch(e){if(e.name!=='AbortError')toast('Erro: '+e.message,'err');}
}

// ══════════════════════════════════════════════════════════
//  HISTÓRICO
// ══════════════════════════════════════════════════════════
async function loadHistory(){
  const raw={};
  const currentYear=STATE.periods.length?decodePeriod(STATE.periods[0]).year:null;
  const processedYears=new Set();
  async function absorbYear(yearHandle,year){
    const yearAWBs=await readAnnualAllMonths(yearHandle,year);
    yearAWBs.forEach(a=>{const mo=parseInt((a.data||'').slice(5,7));if(!mo||mo<1||mo>12)return;if(!raw[a.filial])raw[a.filial]={};if(!raw[a.filial][year])raw[a.filial][year]={};raw[a.filial][year][mo]=(raw[a.filial][year][mo]||0)+a.valor_frete;});
  }
  for await(const[yearStr,yearHandle]of STATE.rootDir.entries()){
    if(yearHandle.kind!=='directory')continue;if(yearStr==='historico')continue;
    const year=parseInt(yearStr);if(isNaN(year)||year<2015||year>2040)continue;
    if(year===currentYear)continue;
    await absorbYear(yearHandle,year);processedYears.add(year);
  }
  if(STATE.historyDir){for await(const[yearStr,yearHandle]of STATE.historyDir.entries()){if(yearHandle.kind!=='directory')continue;const year=parseInt(yearStr);if(isNaN(year)||year<2015||year>2040)continue;if(processedYears.has(year))continue;await absorbYear(yearHandle,year);}}
  STATE.history={};
  Object.keys(raw).forEach(filial=>{STATE.history[filial]=computeFilialHistory(filial,raw[filial]);});
}

async function readAnnualAllMonths(yearHandle,year){
  const dirs=[yearHandle];
  for(const n of['ANUAL','Anual','anual','CONSOLIDADO','Consolidado','consolidado']){try{dirs.push(await yearHandle.getDirectoryHandle(n,{create:false}));}catch(_){}}
  const all=[];for(const dir of dirs)all.push(...await readDirAWBs(dir,year));
  const seen=new Set();return all.filter(a=>{const k=[a.awb,a.data,a.filial,a.source].join('|');if(seen.has(k))return false;seen.add(k);return true;});
}

function computeFilialHistory(filial,byYear){
  const allMonths=[];const byYearTotals={};
  Object.entries(byYear).forEach(([yr,months])=>{const year=parseInt(yr);let yt=0;Object.entries(months).forEach(([mo,total])=>{allMonths.push({year,month:parseInt(mo),total});yt+=total;});byYearTotals[year]=yt;});
  if(!allMonths.length)return null;
  const sorted=[...allMonths].sort((a,b)=>b.total-a.total);
  const bestMonth=sorted[0];const worstMonth=sorted[sorted.length-1];const topMonths=sorted.slice(0,5);
  const bestYearEntry=Object.entries(byYearTotals).sort((a,b)=>b[1]-a[1])[0];
  const bestYear=bestYearEntry?{year:parseInt(bestYearEntry[0]),total:bestYearEntry[1]}:null;
  const avgMonthly=allMonths.reduce((s,m)=>s+m.total,0)/allMonths.length;
  const byMonthNum={};allMonths.forEach(({month,total})=>{if(!byMonthNum[month])byMonthNum[month]=[];byMonthNum[month].push(total);});
  const seasonality={};Object.entries(byMonthNum).forEach(([mo,vals])=>{const avg=vals.reduce((s,v)=>s+v,0)/vals.length;seasonality[parseInt(mo)]=avgMonthly>0?avg/avgMonthly:1;});
  return{byYear:byYearTotals,byMonthNum,seasonality,bestMonth,bestYear,worstMonth,topMonths,avgMonthly,allMonths};
}

// ══════════════════════════════════════════════════════════
//  PROJEÇÃO ANUAL (inclui SC potencial do mês atual)
// ══════════════════════════════════════════════════════════
function computeAnnualProjection(){
  if(!STATE.periods.length)return;
  const curYear=decodePeriod(STATE.periods[0]).year;
  const yesterday=(()=>{const d=new Date();d.setDate(d.getDate()-1);return d.toISOString().split('T')[0];})();
  const tomorrow=(()=>{const d=new Date();d.setDate(d.getDate()+1);return d.toISOString().split('T')[0];})();
  const curMonth=new Date().getMonth()+1;
  const allWD={};for(let mo=1;mo<=12;mo++)allWD[mo]=workingDays(curYear,mo);
  const wdTotalAno=Object.values(allWD).reduce((s,d)=>s+d.length,0);

  const filialAccum={};
  STATE.periods.forEach(p=>{
    const{month}=decodePeriod(p);const d=STATE.cache[p];if(!d)return;
    const wdDoneCount=allWD[month].filter(dt=>dt<=yesterday).length;
    d.cur.forEach(a=>{if(a.data>yesterday)return;const f=a.filial;if(!filialAccum[f])filialAccum[f]={realizado:0,diasUteis:0,scPotencial:0,_periods:new Set()};filialAccum[f].realizado+=a.valor_frete;});
    // SC do mês atual: soma como potencial
    if(month===curMonth&&d.sc){d.sc.filter(a=>a.filial).forEach(a=>{const f=a.filial;if(!filialAccum[f])filialAccum[f]={realizado:0,diasUteis:0,scPotencial:0,_periods:new Set()};filialAccum[f].scPotencial=(filialAccum[f].scPotencial||0)+a.valor_frete;});}
    Object.keys(filialAccum).forEach(f=>{if(!filialAccum[f]._periods.has(p)&&d.cur.some(a=>a.filial===f&&a.data<=yesterday)){filialAccum[f]._periods.add(p);filialAccum[f].diasUteis+=wdDoneCount;}});
  });

  STATE.projection={byFilial:{},consolidated:null};
  STATE.filiais.forEach(filial=>{
    const accum=filialAccum[filial];const hist=STATE.history[filial]||null;
    const histAnualRef=hist?.bestYear?.total||null;const histAvgAnual=hist?Object.values(hist.byYear).reduce((s,v)=>s+v,0)/Object.keys(hist.byYear).length:null;
    if(!accum||accum.diasUteis<3){STATE.projection.byFilial[filial]={projectedTotal:histAnualRef||0,basedOnRunRate:0,seasonalAdjustment:0,confidence:'historico',vsBestYear:histAnualRef,vsHistoricalAvg:histAvgAnual,vsBestYearPct:null,vsHistAvgPct:null,dailyRate:0,diasRealizados:0,realizadoAteHoje:0,scPotencial:0};return;}
    const dailyRate=accum.realizado/accum.diasUteis;
    let wdRestantes=0;for(let mo=curMonth;mo<=12;mo++)wdRestantes+=allWD[mo].filter(d=>d>=tomorrow).length;
    let projSazonal=accum.realizado;let totalSeasonalAdj=0;
    for(let mo=curMonth;mo<=12;mo++){const wdLeft=allWD[mo].filter(d=>d>=tomorrow).length;if(wdLeft<=0)continue;const rawIdx=hist?.seasonality?.[mo]??1.0;const cappedIdx=Math.min(Math.max(rawIdx,0.5),2.0);const adjRate=dailyRate*cappedIdx;projSazonal+=adjRate*wdLeft;totalSeasonalAdj+=(adjRate-dailyRate)*wdLeft;}
    let finalProjection=projSazonal;let confidence='realista';
    if(histAnualRef&&accum.diasUteis<15){const blendW=Math.min(accum.diasUteis/15,1);const histDailyRate=histAnualRef/wdTotalAno;const histEstimate=accum.realizado+histDailyRate*wdRestantes;finalProjection=blendW*projSazonal+(1-blendW)*histEstimate;confidence=accum.diasUteis<8?'conservador':'realista';}
    else if(accum.diasUteis>=15){if(histAnualRef&&finalProjection>histAnualRef*1.08)confidence='agressivo';else if(histAnualRef&&finalProjection<histAnualRef*0.92)confidence='conservador';}
    const vsBestYearPct=histAnualRef?((finalProjection-histAnualRef)/histAnualRef*100):null;
    const vsHistAvgPct=histAvgAnual?((finalProjection-histAvgAnual)/histAvgAnual*100):null;
    STATE.projection.byFilial[filial]={projectedTotal:Math.round(finalProjection),basedOnRunRate:Math.round(accum.realizado+dailyRate*wdRestantes),seasonalAdjustment:Math.round(totalSeasonalAdj),confidence,vsBestYear:histAnualRef,vsHistoricalAvg:histAvgAnual,vsBestYearPct,vsHistAvgPct,dailyRate:Math.round(dailyRate),diasRealizados:accum.diasUteis,realizadoAteHoje:Math.round(accum.realizado),scPotencial:Math.round(accum.scPotencial||0)};
  });
  const vals=Object.values(STATE.projection.byFilial).filter(v=>v.realizadoAteHoje>0);
  if(vals.length){const conf=vals.every(v=>v.confidence==='agressivo')?'agressivo':vals.some(v=>v.confidence==='conservador')?'conservador':'realista';STATE.projection.consolidated={projectedTotal:vals.reduce((s,v)=>s+v.projectedTotal,0),realizadoAteHoje:vals.reduce((s,v)=>s+(v.realizadoAteHoje||0),0),scPotencial:vals.reduce((s,v)=>s+(v.scPotencial||0),0),confidence:conf};}
}

// ══════════════════════════════════════════════════════════
//  SCAN + LOAD PERÍODOS
// ══════════════════════════════════════════════════════════
async function openFolder(){if(!window.showDirectoryPicker){toast('Use Chrome ou Edge.','err');return;}try{STATE.rootDir=await window.showDirectoryPicker({mode:'read'});await saveRootDirHandle(STATE.rootDir);await scanAllPeriods();}catch(e){if(e.name!=='AbortError')toast('Erro: '+e.message,'err');}}

// Salvar último período usado
function saveLastPeriod(){if(STATE.periods.length>0){localStorage.setItem('freight_last_period',STATE.periods[STATE.curIdx]);localStorage.setItem('frecht_last_period_idx',STATE.curIdx);}}

// Carregar último período usado
function loadLastPeriod(){const lastPeriod=localStorage.getItem('freight_last_period');const lastIdx=localStorage.getItem('frecht_last_period_idx');return{period:lastPeriod,idx:lastIdx?parseInt(lastIdx):null};}

// Tentar abrir automaticamente a última pasta ao iniciar
async function tryAutoOpenFolder(){
  const handle=await loadRootDirHandle();
  if(handle){
    try{
      STATE.rootDir=handle;
      await scanAllPeriods();
      toast('Pasta anterior reaberta automaticamente','ok');
    }catch(e){
      console.error('Erro ao reabrir pasta:',e);
      toast('Erro ao reabrir pasta anterior','err');
    }
  }
}

// Monitoramento por foco — mais barato que polling 5s
let fileModDates={};let _refocusBound=false;

async function snapshotFileModDates(){
  if(!STATE.rootDir)return{};
  const out={};
  async function walk(dir,path){
    for await(const[name,h]of dir.entries()){
      if(h.kind==='file'){if(classifyFile(name)){const f=await h.getFile();out[`${path}/${name}`]=f.lastModified;}}
      else if(h.kind==='directory'&&!/^historico$/i.test(name)){await walk(h,`${path}/${name}`);}
    }
  }
  try{await walk(STATE.rootDir,'');}catch(e){}
  return out;
}

async function checkAndReloadIfChanged(silent=false){
  if(!STATE.rootDir||STATE._reloading)return;
  const cur=await snapshotFileModDates();
  const keys=new Set([...Object.keys(cur),...Object.keys(fileModDates)]);
  let changed=false;
  for(const k of keys){if(cur[k]!==fileModDates[k]){changed=true;break;}}
  if(changed){
    fileModDates=cur;
    if(!silent)toast('Arquivos atualizados. Recarregando...','ok');
    STATE._reloading=true;
    try{await scanAllPeriods({preserveTab:true});}finally{STATE._reloading=false;}
  }
}

function startFileMonitoring(){
  snapshotFileModDates().then(d=>{fileModDates=d;});
  if(_refocusBound)return;_refocusBound=true;
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')checkAndReloadIfChanged(true);});
  window.addEventListener('focus',()=>checkAndReloadIfChanged(true));
}

function stopFileMonitoring(){/* noop — eventos persistem */}

// Walker recursivo — encontra arquivos XLS em qualquer profundidade.
// Infere ano/mês via caminho (preferência) ou nome do arquivo.
// Retorna: { byPeriod: {YYYY-MM: [{handle,name,info,year,month}], ...}, allYears: Set, structureMode }
async function discoverAllFiles(){
  const collected=[];
  async function walk(dir,segments){
    let entries;
    try{entries=[];for await(const e of dir.entries())entries.push(e);}catch(e){return;}
    const fileTasks=[];
    for(const[name,h]of entries){
      if(h.kind==='file'){
        const info=classifyFile(name);
        if(!info)continue;
        fileTasks.push({handle:h,name,info,segments:[...segments]});
      }else if(h.kind==='directory'){
        if(/^historico$/i.test(name))continue;
        await walk(h,[...segments,name]);
      }
    }
    collected.push(...fileTasks);
  }
  await walk(STATE.rootDir,[]);
  return collected;
}

// Infere {year,month} do caminho de pastas. Retorna null se não conseguir.
function inferYearMonthFromPath(segments){
  let year=null,month=null;
  for(const s of segments){
    const n=parseInt(s,10);
    if(!isNaN(n)){
      if(n>=2015&&n<=2040)year=n;
      else if(n>=1&&n<=12&&!month)month=n;
    }
  }
  return{year,month};
}

async function scanAllPeriods(opts={}){
  const preserveTab=!!opts.preserveTab;
  const prevActiveTab=preserveTab?(document.querySelector('.tab.active')?.id||null):null;
  const prevPeriodKey=preserveTab&&STATE.periods[STATE.curIdx]?STATE.periods[STATE.curIdx]:null;

  showLoading('Varrendo pastas...');
  try{STATE.historyDir=await STATE.rootDir.getDirectoryHandle('historico',{create:false});}catch(e){STATE.historyDir=null;}

  const allFiles=await discoverAllFiles();
  if(!allFiles.length){hideLoading();toast('Nenhum arquivo XLS reconhecido encontrado.','err');return;}

  // Agrupa por (year,month). Quando ano/mês não vem do path, lê o XLS para descobrir pela maior data.
  const groups=new Map(); // key: "YYYY-MM" → array<filehandle entry>
  const needsContentInference=[];
  for(const f of allFiles){
    const{year,month}=inferYearMonthFromPath(f.segments);
    if(year&&month){
      const k=`${year}-${padM(month)}`;
      f.year=year;f.month=month;
      if(!groups.has(k))groups.set(k,[]);
      groups.get(k).push(f);
    }else{
      needsContentInference.push(f);
    }
  }

  // Para arquivos sem ano/mês no path, leitura prévia para inferir do conteúdo
  if(needsContentInference.length){
    showLoading('Identificando períodos...',`${needsContentInference.length} arquivo(s) sem ano/mês no caminho`,5);
    const inferred=await Promise.all(needsContentInference.map(async f=>{
      const recs=await parseXLS(f.handle,f.info.filial,f.info.source,new Date().getFullYear(),f.name);
      // pega data mais frequente por mês
      const monthCount={};
      recs.forEach(r=>{const k=r.data.slice(0,7);monthCount[k]=(monthCount[k]||0)+1;});
      const best=Object.entries(monthCount).sort((a,b)=>b[1]-a[1])[0];
      if(!best)return null;
      const[y,m]=best[0].split('-').map(Number);
      return{f,year:y,month:m,records:recs};
    }));
    inferred.filter(Boolean).forEach(({f,year,month})=>{
      const k=`${year}-${padM(month)}`;
      f.year=year;f.month=month;
      if(!groups.has(k))groups.set(k,[]);
      groups.get(k).push(f);
    });
  }

  if(!groups.size){hideLoading();toast('Não consegui identificar nenhum período (ano/mês) nos arquivos.','err');return;}

  const sortedKeys=[...groups.keys()].sort();
  const latestYear=Math.max(...sortedKeys.map(k=>parseInt(k.slice(0,4))));
  STATE.periods=sortedKeys.filter(k=>parseInt(k.slice(0,4))===latestYear);
  STATE.cache={};STATE.curIdx=STATE.periods.length-1;
  STATE._allGroups=groups;

  const apEl=$('ap');if(apEl){apEl.innerHTML='<option value="">Todos</option>';sortedKeys.forEach(p=>{const{year,month}=decodePeriod(p);const opt=document.createElement('option');opt.value=p;opt.textContent=`${MONTHS[month].substring(0,3)} ${year}`;apEl.appendChild(opt);});}

  $('pane-welcome').classList.remove('active');$('main-tabs').style.display='flex';$('period-nav').style.display='flex';$('hdr-status').style.display='inline';$('open-btn-lbl').textContent='Recarregar';

  // Paraleliza leitura de períodos do ano corrente
  const total=STATE.periods.length;
  showLoading(`Lendo ${total} mês${total>1?'es':''}...`,'Em paralelo',10);
  let done=0;
  await Promise.all(STATE.periods.map(async p=>{
    const{year,month}=decodePeriod(p);
    STATE.cache[p]=await loadPeriodFromGroup(groups.get(p)||[],year,month);
    done++;
    showLoading(`Lendo ${total} mês${total>1?'es':''}...`,`${done} de ${total} concluídos`,Math.round(10+(done/total)*80));
  }));

  showLoading('Lendo histórico...','Calculando métricas históricas...',92);
  await loadHistoryFromGroups(groups,latestYear);
  hideLoading();

  toast(`${done} mês(es) carregados${Object.keys(STATE.history||{}).length?' + histórico':''}${countSCTotal()>0?' + SC':''}`,'ok');
  saveLastPeriod();
  discoverFiliaisFromData();buildDynamicTabs();buildDynamicPanes();applyBranchLabels();
  computeFreshnessStatus();renderFreshnessBanner();
  renderIntelFreteAnalysis();renderIntelTopClientes();
  const ovTab=document.querySelector('.tab[onclick*="overview"]');if(ovTab)ovTab.style.display=STATE.filiais.length<=1?'none':'';
  rebuildAllAWBIndex();computeAnnualProjection();

  // Restaurar período/aba anterior
  if(prevPeriodKey&&STATE.periods.includes(prevPeriodKey))STATE.curIdx=STATE.periods.indexOf(prevPeriodKey);
  await loadAndRender();
  startFileMonitoring();

  $$('.tab').forEach(t=>t.classList.remove('active'));$$('.pane').forEach(p=>p.classList.remove('active'));
  if(prevActiveTab&&$(prevActiveTab)){
    const tab=$(prevActiveTab);tab.classList.add('active');
    const pane=$(prevActiveTab.replace(/^tab-/,'pane-')) || (tab.getAttribute('onclick')?.match(/'([^']+)'/)?.[1]?$(`pane-${tab.getAttribute('onclick').match(/'([^']+)'/)[1]}`):null);
    if(pane){pane.classList.add('active');renderPane(pane.id.replace('pane-',''));}
  }else if(STATE.filiais.length===1){const f=STATE.filiais[0];const tab=$(`tab-${f.toLowerCase()}`);const pane=$(`pane-filial_${f}`);if(tab)tab.classList.add('active');if(pane){pane.classList.add('active');renderPane('filial_'+f);}}
  else{const ot=document.querySelector('.tab[onclick*="overview"]');if(ot)ot.classList.add('active');const op=$('pane-overview');if(op){op.classList.add('active');renderPane('overview');}}
}

// Lê todos arquivos de um período em paralelo
async function loadPeriodFromGroup(filesInGroup,year,month){
  if(!filesInGroup||!filesInGroup.length){
    // pode ainda querer carregar comparativo prev — preserva fluxo antigo
    const prevDir=await getMonthDir(year-1,padM(month));
    const prev=prevDir?await readDirAWBs(prevDir,year-1):await readAnnualYearAWBs(year-1,month);
    return{cur:[],prev,sc:[]};
  }
  const results=await Promise.all(filesInGroup.map(f=>parseXLS(f.handle,f.info.filial,f.info.source,year,f.name)));
  const all=results.flat();
  const cur=all.filter(a=>a.source!=='sc'&&a.source!=='neg');
  const sc=all.filter(a=>a.source==='sc');
  // Comparativo do mesmo mês ano anterior
  const prevDir=await getMonthDir(year-1,padM(month));
  let prev=prevDir?await readDirAWBs(prevDir,year-1):[];
  if(!prev.length){
    // tenta achar no allGroups o mesmo mês ano anterior
    const prevKey=`${year-1}-${padM(month)}`;
    const prevGroup=STATE._allGroups?.get(prevKey);
    if(prevGroup&&prevGroup.length){
      const prevResults=await Promise.all(prevGroup.map(f=>parseXLS(f.handle,f.info.filial,f.info.source,year-1,f.name)));
      prev=prevResults.flat().filter(a=>a.source!=='sc');
    }else{
      prev=await readAnnualYearAWBs(year-1,month);
    }
  }
  // registrar SC
  if(!STATE.scFiliais)STATE.scFiliais=new Set();
  sc.forEach(a=>STATE.scFiliais.add(a.filial));
  return{cur,prev,sc};
}

async function loadHistoryFromGroups(groups,latestYear){
  const raw={};
  const prevYears=new Set();
  for(const[k,files]of groups.entries()){
    const y=parseInt(k.slice(0,4));
    if(y===latestYear)continue;
    prevYears.add(y);
    const results=await Promise.all(files.map(f=>parseXLS(f.handle,f.info.filial,f.info.source,y,f.name)));
    results.flat().forEach(a=>{const mo=parseInt((a.data||'').slice(5,7));if(!mo||mo<1||mo>12)return;if(!raw[a.filial])raw[a.filial]={};if(!raw[a.filial][y])raw[a.filial][y]={};raw[a.filial][y][mo]=(raw[a.filial][y][mo]||0)+a.valor_frete;});
  }
  // historicoDir antigo
  if(STATE.historyDir){
    for await(const[yearStr,yearHandle]of STATE.historyDir.entries()){
      if(yearHandle.kind!=='directory')continue;const year=parseInt(yearStr);if(isNaN(year)||year<2015||year>2040)continue;if(prevYears.has(year))continue;
      const yearAWBs=await readAnnualAllMonths(yearHandle,year);
      yearAWBs.forEach(a=>{const mo=parseInt((a.data||'').slice(5,7));if(!mo||mo<1||mo>12)return;if(!raw[a.filial])raw[a.filial]={};if(!raw[a.filial][year])raw[a.filial][year]={};raw[a.filial][year][mo]=(raw[a.filial][year][mo]||0)+a.valor_frete;});
    }
  }
  STATE.history={};
  Object.keys(raw).forEach(filial=>{STATE.history[filial]=computeFilialHistory(filial,raw[filial]);});
}

function countSCTotal(){let t=0;Object.values(STATE.cache).forEach(d=>{if(d.sc)t+=d.sc.length;});return t;}

async function navMonth(dir){const newIdx=STATE.curIdx+dir;if(newIdx<0||newIdx>=STATE.periods.length)return;STATE.curIdx=newIdx;await loadAndRender();}
function updateNavBtns(){$('btn-prev').disabled=STATE.curIdx<=0;$('btn-next').disabled=STATE.curIdx>=STATE.periods.length-1;}

async function loadAndRender(){
  const periodKey=STATE.periods[STATE.curIdx];const{year:curYear,month}=decodePeriod(periodKey);const prevYear=curYear-1;
  updateNavBtns();updatePeriodChip(curYear,month,prevYear);
  if(!STATE.cache[periodKey]){showLoading(`Lendo ${MONTHS[month]} ${curYear}...`);STATE.cache[periodKey]=await loadPeriodData(curYear,prevYear,month);hideLoading();}
  const data=STATE.cache[periodKey];const wdCur=workingDays(curYear,month);const wdPrev=workingDays(prevYear,month);
  const _allDts=data.cur.map(a=>a.data).sort();const _dtRange=_allDts.length?` · ${_allDts[0]} → ${_allDts[_allDts.length-1]}`:'';
  const scCount=data.sc?data.sc.length:0;
  $('hdr-status').textContent=`${data.cur.length} cur + ${data.prev.length} prev + ${scCount} SC${_dtRange}`;
  $$('.lbl-cur').forEach(el=>el.textContent=curYear);$$('.lbl-prev').forEach(el=>el.textContent=prevYear);
  renderCurrentPane(data,wdCur,wdPrev,curYear,prevYear,month);
}

async function getMonthDir(year,mo){try{let d=STATE.rootDir;d=await d.getDirectoryHandle(String(year),{create:false});d=await d.getDirectoryHandle(mo,{create:false});return d;}catch(e){return null;}}
async function getYearDir(year){try{return await STATE.rootDir.getDirectoryHandle(String(year),{create:false});}catch(e){return null;}}

async function readAnnualYearAWBs(year,month){
  const yearDir=await getYearDir(year);if(!yearDir)return[];
  const dirs=[yearDir];for(const n of['ANUAL','Anual','anual','CONSOLIDADO','Consolidado','consolidado']){try{dirs.push(await yearDir.getDirectoryHandle(n,{create:false}));}catch(e){}}
  const all=[];for(const dir of dirs)all.push(...await readDirAWBs(dir,year));
  const seen=new Set();return all.filter(a=>{const recMonth=new Date(a.data+'T12:00:00').getMonth()+1;if(recMonth!==month)return false;const key=[a.awb,a.data,a.filial,a.source,a.valor_frete].join('|');if(seen.has(key))return false;seen.add(key);return true;});
}

// loadPeriodData — agora retorna {cur, prev, sc}
async function loadPeriodData(curYear,prevYear,month){
  const mo=padM(month);
  const curDir=await getMonthDir(curYear,mo);const prevDir=await getMonthDir(prevYear,mo);
  const allCur=curDir?await readDirAWBs(curDir,curYear):[];
  const prev=prevDir?await readDirAWBs(prevDir,prevYear):await readAnnualYearAWBs(prevYear,month);
  // Separar SC do cur
  const cur=allCur.filter(a=>a.source!=='sc');
  const sc=allCur.filter(a=>a.source==='sc');
  return{cur,prev,sc};
}

// ══════════════════════════════════════════════════════════
//  PARSE XLS
// ══════════════════════════════════════════════════════════
async function readDirAWBs(dirHandle,year){
  // Coleta tarefas e dispara em paralelo
  const tasks=[];
  for await(const[name,handle]of dirHandle.entries()){
    if(handle.kind==='file'){const info=classifyFile(name);if(!info)continue;tasks.push(parseXLS(handle,info.filial,info.source,year,name));}
    else if(handle.kind==='directory'){for await(const[subName,subHandle]of handle.entries()){if(subHandle.kind!=='file')continue;const info=classifyFile(subName);if(!info)continue;tasks.push(parseXLS(subHandle,info.filial,info.source,year,subName));}}
  }
  return(await Promise.all(tasks)).flat();
}

function pNum(v){return parseFloat(String(v||'').replace(',','.'))||0;}
function normModal(v){const s=String(v||'').trim().toUpperCase();if(!s)return'N/D';if(s==='A'||s.includes('AER'))return'Aéreo';if(s==='R'||s.includes('ROD'))return'Rodoviário';return s;}
function modalClass(v){const s=String(v||'N/D').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();if(s.includes('aereo'))return'aereo';if(s.includes('rodoviario'))return'rodoviario';return'nd';}

async function parseXLS(handle,filial,source,year,filename){
  try{
    const file=await handle.getFile();
    const lastModified=new Date(file.lastModified).toLocaleDateString('pt-BR');
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array',cellDates:true});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});
    let hi=0;for(let i=0;i<Math.min(8,rows.length);i++){if(rows[i].some(c=>/AWB/i.test(String(c)))){hi=i;break;}}
    // Fallback: se não encontrou AWB, procura Data/Hora no título e começa na linha seguinte
    if(hi===0&&source==='sc'){for(let i=0;i<Math.min(10,rows.length);i++){if(rows[i].some(c=>/Remetente|Data|Modal/i.test(String(c)))){hi=i;break;}}}
    const hdr=rows[hi].map(h=>String(h).trim());
    const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
    const normHdr=hdr.map(norm);
    const ci=(...ns)=>{for(const n of ns){const nk=norm(n);const exact=normHdr.findIndex(h=>h===nk);if(exact>=0)return exact;}for(const n of ns){const nk=norm(n);const partial=normHdr.findIndex(h=>h.includes(nk));if(partial>=0)return partial;}return -1;};
    const I={
      awb:ci('Nro. AWB','Numero AWB','Nr. AWB','Num. AWB','AWB','Nro','Nr.','Numero'),
      data:ci('Data Emissão','Data Emissao','Dt. Emissão','Dt. Emissao','Dt Emissao','Data de Emissão','Emitido','Emissão','Data','Dt.'),
      rem:ci('Remetente','Nome Remetente','Razão Social Rem','Razao Social Rem','Remet'),
      dest:ci('Destinatário','Destinatario','Nome Destinatário','Razão Social Dest','Razao Social Dest','Dest'),
      cid:ci('Cidade Destino','Cidade Dest','Municipio','Município','Cidade'),
      uf:ci('UF Destino','UF Dest','Estado Destino','Estado','UF'),
      vol:ci('Volumes','Qtd Volumes','Qtd. Volumes','Quantidade','Volume','Vol','Qtd'),
      peso:ci('Peso Real','Peso Bruto','Peso Kg','Peso (kg)','Peso'),
      merc:ci('Valor Mercantil','Valor da Mercadoria','Vl Mercantil','Vlr Mercantil','Vl. Merc.','Mercantil','Merc'),
      fret:ci('Valor do Frete','Valor Frete Total','Valor Frete','Vl. do Frete','Vl Frete','Vlr Frete','Vl. Frete','Total Frete','Frete'),
      tipo:ci('Tipo de Frete','Tipo Frete','Tp. Frete','Tp Frete','Frete Tipo'),
      modal:ci('Modal de Transporte','Modalidade Transporte','Modalidade','Modal'),
      ramo:ci('Ramo de Atividade','Ramo Atividade','Segmento Atividade','Ramo','Segmento'),
      cnpjRem:ci('CNPJ Remetente','CNPJ/CPF Remetente','CNPJ Rem','CNPJ do Remetente','CNPJ Remet.'),
      cnpjDest:ci('CNPJ Destinatário','CNPJ Destinatario','CNPJ/CPF Destinatário','CNPJ/CPF Destinatario','CNPJ Dest'),
      cnpj:ci('CNPJ Remetente','CNPJ/CPF','C.N.P.J.','C.N.P.J','CNPJ','Cnpj'),
    };
    if(I.data<0){for(let col=0;col<hdr.length;col++){const sample=rows[hi+1]?.[col];if(!sample&&sample!==0)continue;if(sample instanceof Date){I.data=col;break;}if(typeof sample==='number'&&sample>40000&&sample<55000){I.data=col;break;}if(typeof sample==='string'&&/\d{2}\/\d{2}\/\d{4}/.test(sample)){I.data=col;break;}if(typeof sample==='string'&&/\d{4}-\d{2}-\d{2}/.test(sample)){I.data=col;break;}}}
    const recs=[];
    for(let i=hi+1;i<rows.length;i++){
      const r=rows[i];const awbVal=r[I.awb];if(!awbVal&&awbVal!==0)continue;
      let dt=I.data>=0?(r[I.data]??''):'';
      if(dt instanceof Date){const y=dt.getFullYear(),m=String(dt.getMonth()+1).padStart(2,'0'),d=String(dt.getDate()).padStart(2,'0');dt=`${y}-${m}-${d}`;}
      else if(typeof dt==='number'&&dt>0){try{const msUtc=Math.round((dt-25569)*86400*1000);const jsDate=new Date(msUtc);if(!isNaN(jsDate.getTime())){const y=jsDate.getUTCFullYear(),m=String(jsDate.getUTCMonth()+1).padStart(2,'0'),d=String(jsDate.getUTCDate()).padStart(2,'0');dt=`${y}-${m}-${d}`;}else continue;}catch(e){continue;}}
      else{dt=String(dt).trim().split(' ')[0].split('T')[0];if(!dt)continue;if(dt.includes('/')){const p=dt.split('/').map(v=>v.trim());if(p.length===3){let[a,b,c]=p;a=a.padStart(2,'0');b=b.padStart(2,'0');if(c.length===2)c='20'+c;dt=`${c}-${b}-${a}`;}}const tmp=new Date(dt+'T12:00:00');if(isNaN(tmp.getTime()))continue;dt=tmp.toISOString().slice(0,10);}
      const cnpjRem=I.cnpjRem>=0?(String(r[I.cnpjRem]||'').replace(/\D/g,'').replace(/^0+/,'')||''):(I.cnpj>=0?(String(r[I.cnpj]||'').replace(/\D/g,'').replace(/^0+/,'')||''):'');
      const cnpjDest=I.cnpjDest>=0?(String(r[I.cnpjDest]||'').replace(/\D/g,'').replace(/^0+/,'')||''):'';
      recs.push({awb:String(awbVal).trim(),data:dt,filial,source,year,remetente:String(r[I.rem]||'').trim().substring(0,45),destinatario:String(r[I.dest]||'').substring(0,45),cidade:String(r[I.cid]||''),uf:String(r[I.uf]||''),volumes:parseInt(r[I.vol])||0,peso:pNum(r[I.peso]),valor_mercantil:pNum(r[I.merc]),valor_frete:pNum(r[I.fret]),tipo_frete:String(r[I.tipo]||''),modal:normModal(r[I.modal]),ramo:String(r[I.ramo]||'').substring(0,40),cnpj:cnpjRem,cnpj_remetente:cnpjRem,cnpj_destinatario:cnpjDest,periodo:STATE.periods[STATE.curIdx]||'',arquivo:filename||'',data_modificacao:lastModified||''});
    }
    return recs;
  }catch(e){return[];}
}

// ══════════════════════════════════════════════════════════
//  AGREGAÇÕES — inclui SC
// ══════════════════════════════════════════════════════════
function stats(awbs,filial){
  const r=filial?awbs.filter(a=>a.filial===filial):awbs;
  const cart=r.filter(a=>a.source==='carteira').reduce((s,a)=>s+a.valor_frete,0);
  const extra=r.filter(a=>a.source==='extra').reduce((s,a)=>s+a.valor_frete,0);
  return{total:cart+extra,carteira:cart,extra,awbs:r.length,volumes:r.reduce((s,a)=>s+a.volumes,0),aereo:r.filter(a=>a.modal==='Aéreo').reduce((s,a)=>s+a.valor_frete,0),rodoviario:r.filter(a=>a.modal==='Rodoviário').reduce((s,a)=>s+a.valor_frete,0),awbsAereo:r.filter(a=>a.modal==='Aéreo').length,awbsRodoviario:r.filter(a=>a.modal==='Rodoviário').length};
}

function scStats(sc,filial){
  if(!sc||!sc.length)return{total:0,awbs:0};
  const r=filial?sc.filter(a=>a.filial===filial):sc;
  const total=r.reduce((s,a)=>s+a.valor_frete,0);
  return{total,awbs:r.length};
}

// dmap — inclui SC
function dmap(awbs,filial,sc){
  const map={};
  (filial?awbs.filter(a=>a.filial===filial):awbs).forEach(a=>{if(!map[a.data])map[a.data]={carteira:0,extra:0,total:0,volumes:0,awbs:0};map[a.data][a.source]=(map[a.data][a.source]||0)+a.valor_frete;map[a.data].total+=a.valor_frete;map[a.data].volumes+=a.volumes;map[a.data].awbs++;});
  // SC por dia
  if(sc&&sc.length){(filial?sc.filter(a=>a.filial===filial):sc).forEach(a=>{if(!map[a.data])map[a.data]={carteira:0,extra:0,total:0,volumes:0,awbs:0,sc:0,scAwbs:0};map[a.data].sc=(map[a.data].sc||0)+a.valor_frete;map[a.data].scAwbs=(map[a.data].scAwbs||0)+1;});}
  return map;
}

// ══════════════════════════════════════════════════════════
//  RENDER DISPATCHER
// ══════════════════════════════════════════════════════════
let _data=null,_wdC=null,_wdP=null,_curY=null,_prevY=null,_mo=null;
function renderCurrentPane(data,wdCur,wdPrev,curYear,prevYear,month){_data=data;_wdC=wdCur;_wdP=wdPrev;_curY=curYear;_prevY=prevYear;_mo=month;const pane=document.querySelector('.pane.active')?.id?.replace('pane-','')||'overview';renderPane(pane);}
function renderPane(name){if(!_data)return;if(name==='overview')renderOverview();else if(name.startsWith('filial_'))renderFilial(name.replace('filial_',''));else if(name==='awb')initAWB();else if(name==='neg')renderNeg();else if(name==='intel')renderIntel();else if(name==='comissao')renderComissao();}

// ══════════════════════════════════════════════════════════
//  OVERVIEW
// ══════════════════════════════════════════════════════════
function renderOverview(){
  const{cur,prev,sc}=_data;
  const sCurAll=stats(cur,null);const sPrevAll=stats(prev,null);
  const scAll=scStats(sc,null);
  const filialStats=STATE.filiais.map(f=>({filial:f,cur:stats(cur,f),prev:stats(prev,f),sc:scStats(sc,f),dmCur:dmap(cur,f,sc)}));
  const today=new Date().toISOString().split('T')[0];
  const elapsedDays=_wdC.filter(d=>d<=today).length;
  const proj=projection(sCurAll.total,_wdC);const pctVs=proRata(sCurAll.total,sPrevAll.total,_wdC,_wdP);const gap=proj-sPrevAll.total;

  if($('wd-cur-year'))$('wd-cur-year').textContent=_curY;if($('wd-prev-year'))$('wd-prev-year').textContent=_prevY;if($('wd-prev-year2'))$('wd-prev-year2').textContent=_prevY;
  if($('wd-cur-total'))$('wd-cur-total').textContent=`${_wdC.length} dias`;if($('wd-done'))$('wd-done').textContent=`${elapsedDays} dias`;if($('wd-left'))$('wd-left').textContent=`${_wdC.filter(d=>d>today).length} dias`;if($('wd-prev-total'))$('wd-prev-total').textContent=`${_wdP.length} dias`;if($('wd-prev-avg'))$('wd-prev-avg').textContent=R$(_wdP.length>0?sPrevAll.total/_wdP.length:0);
  if($('ov-period'))$('ov-period').textContent=`${MONTHS[_mo]} ${_curY}`;if($('ov-note'))$('ov-note').textContent=`${cur.length+prev.length} AWBs comissionados · ${scAll.awbs} SC`;


  set('ov-tot-cur',R$(sCurAll.total),'color:var(--accent)');set('ov-tot-prev',R$(sPrevAll.total));
  const pEl=$('ov-pct');if(pEl){pEl.textContent=pct(pctVs);pEl.style.color=pctVs>=0?'var(--success)':'var(--danger)';}
  set('ov-proj',R$(proj));const gEl=$('ov-gap');if(gEl){gEl.textContent=isFinite(gap)?R$(Math.abs(gap)):'—';gEl.style.color=gap>=0?'var(--success)':'var(--danger)';}
  filialStats.forEach(({filial,cur:fCur,prev:fPrev})=>{const fl=filial.toLowerCase();const pF=proRata(fCur.total,fPrev.total,_wdC,_wdP);const ovEl=$(`ov-${fl}`);const ovSub=$(`ov-${fl}-sub`);if(ovEl)set(`ov-${fl}`,R$(fCur.total));if(ovSub)ovSub.innerHTML=fCur.total>0?`<span class="badge ${bcl(pF)}">${pct(pF)} vs ${_prevY}</span>`:`<span style="color:var(--text3);font-size:11px">sem dados</span>`;});
  set('ov-awbs',(sCurAll.awbs).toLocaleString('pt-BR'));if($('ov-awbs-sub'))$('ov-awbs-sub').textContent=`${_prevY}: ${sPrevAll.awbs.toLocaleString('pt-BR')}`;
  set('ov-cart',R$(sCurAll.carteira));if($('ov-cart-sub'))$('ov-cart-sub').textContent=`${_prevY}: ${R$(sPrevAll.carteira)}`;
  set('ov-extra',R$(sCurAll.extra));if($('ov-extra-sub'))$('ov-extra-sub').textContent=`${_prevY}: ${R$(sPrevAll.extra)}`;
  const med=elapsedDays>0?sCurAll.total/elapsedDays:0;set('ov-avg',R$(med));if($('ov-avg-sub'))$('ov-avg-sub').textContent=`Média ${_prevY}: ${R$(_wdP.length>0?sPrevAll.total/_wdP.length:0)}`;
  renderOvCharts(cur,prev,sc,filialStats,proj,sPrevAll.total);
}

function renderOvCharts(cur,prev,sc,filialStats,proj,totPrev){
  const labels=_wdC.map(d=>fd(d));const colors=FILIAL_COLORS;const dmPrevAll=dmap(prev,null,null);
  const vPrev=_wdC.map((_,i)=>{const dp=_wdP[i];return dp?(dmPrevAll[dp]?.total||0):0;});
  const datasets=filialStats.map(({filial,cur:fCur,dmCur:fDm},i)=>({label:`${filial} ${_curY}`,data:_wdC.map(d=>fDm[d]?.total||0),backgroundColor:colors[i%colors.length].bar,borderRadius:3,stack:'c'}));
  // SC overlay
  if(sc&&sc.length){const dmScAll=dmap([],null,sc);datasets.push({label:'SC',data:_wdC.map(d=>dmScAll[d]?.sc||0),type:'bar',backgroundColor:'rgba(167,139,250,.35)',borderRadius:3,stack:'sc'});}
  datasets.push({label:String(_prevY),data:vPrev,type:'line',borderColor:'rgba(100,116,139,.5)',backgroundColor:'transparent',tension:.3,pointRadius:1,borderWidth:1.5,order:0});
  dch('chOvD');if($('ch-ov-daily'))CHARTS.chOvD=mkChart('ch-ov-daily','bar',{labels,datasets},co({stacked:true}));
  const totals=filialStats.map(({cur:fCur})=>fCur.total);const totSp=totals.reduce((s,v)=>s+v,0);
  dch('chOvSp');if($('ch-ov-split'))CHARTS.chOvSp=mkChart('ch-ov-split','bar',{labels:filialStats.map(f=>f.filial),datasets:[{data:totals,backgroundColor:filialStats.map((_,i)=>colors[i%colors.length].bar),borderRadius:4,borderSkipped:false}]},{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+R$(c.raw)+' ('+(c.raw/(totSp||1)*100).toFixed(1)+'%)'}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}},y:{grid:{display:false},ticks:{color:'#7c93b0',font:{size:12}}}}});
  const today=new Date().toISOString().split('T')[0];const dmAllCur=dmap(cur,null,null);let run=0,days=0;
  const projLine=_wdC.map(d=>{const t=dmAllCur[d]?.total||0;if(d<=today){run+=t;days++;return days>0?Math.round(run/days*_wdC.length):null;}return null;});
  dch('chOvPr');if($('ch-ov-proj'))CHARTS.chOvPr=mkChart('ch-ov-proj','line',{labels,datasets:[{label:'Projeção',data:projLine,borderColor:'var(--success)',backgroundColor:'rgba(16,185,129,.08)',fill:true,tension:.3,pointRadius:0},{label:`Meta ${_prevY}`,data:new Array(_wdC.length).fill(totPrev),borderColor:'rgba(249,115,22,.5)',borderDash:[6,4],backgroundColor:'transparent',pointRadius:0,borderWidth:1.5}]},co({}));
  filialStats.slice(0,3).forEach(({filial},i)=>{const id=`ch-ov-r${i}`;let card=$(id);if(!card){const container=$('ov-ramos-charts');if(container){const div=document.createElement('div');div.className='chart-card';div.innerHTML=`<div class="chart-ttl" id="ch-ov-r${i}-ttl">Top Ramos ${filial}</div><canvas id="${id}" height="120"></canvas>`;container.appendChild(div);card=$(id);}}const ttlEl=$(`ch-ov-r${i}-ttl`);if(ttlEl)ttlEl.textContent=`Top Ramos ${filial}`;if(card)ramos(id,`chR${i}`,filial,colors[i%colors.length].bar,cur);});
  let tfCard=$('ch-ov-tipo');if(!tfCard){const container=$('ov-tipo-charts');if(container){const div=document.createElement('div');div.style.cssText='margin-bottom:20px';div.innerHTML=`<div class="chart-card"><div class="chart-ttl">Tipo de Frete — Consolidado</div><div style="height:120px"><canvas id="ch-ov-tipo"></canvas></div></div>`;container.appendChild(div);tfCard=$('ch-ov-tipo');}}if(tfCard)tipoFreteBar('ch-ov-tipo','chOvTF',null,cur);
}

function ramos(canvId,key,filial,color,awbs){
  const map={};awbs.filter(a=>a.filial===filial).forEach(a=>{const r=a.ramo||'Outros';map[r]=(map[r]||0)+a.valor_frete;});
  const top=Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,7);dch(key);if(!top.length)return;
  CHARTS[key]=mkChart(canvId,'bar',{labels:top.map(r=>r[0].substring(0,22)),datasets:[{data:top.map(r=>r[1]),backgroundColor:color,borderRadius:3}]},{indexAxis:'y',responsive:true,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+R$(c.raw)}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}},y:{grid:{display:false},ticks:{color:'#7c93b0',font:{size:10}}}}});
}

function tipoFreteBar(canvId,key,filial,awbs){
  const src=filial?awbs.filter(a=>a.filial===filial):awbs;const map={};src.forEach(a=>{const t=(String(a.tipo_frete||'').trim().toUpperCase())||'N/D';map[t]=(map[t]||0)+a.valor_frete;});
  const sorted=Object.entries(map).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);const total=sorted.reduce((s,[,v])=>s+v,0);const colors=['rgba(6,182,212,.85)','rgba(249,115,22,.85)','rgba(139,92,246,.85)','rgba(16,185,129,.85)','rgba(148,163,184,.85)'];
  dch(key);if(!sorted.length)return;
  CHARTS[key]=mkChart(canvId,'bar',{labels:sorted.map(([k])=>k),datasets:[{data:sorted.map(([,v])=>v),backgroundColor:sorted.map((_,i)=>colors[i%colors.length]),borderRadius:4,borderSkipped:false}]},{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+R$(c.raw)+' ('+(c.raw/total*100).toFixed(1)+'%)'}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}},y:{grid:{display:false},ticks:{color:'#7c93b0',font:{size:11}}}}});
}

function topClientes(canvId,key,filial,color,awbs){
  const map={};awbs.filter(a=>a.filial===filial).forEach(a=>{const nome=(a.remetente||'').trim()||'Desconhecido';map[nome]=(map[nome]||0)+a.valor_frete;});
  const sorted=Object.entries(map).sort((a,b)=>b[1]-a[1]);const total=sorted.reduce((s,[,v])=>s+v,0);const top5=sorted.slice(0,5);const resto=sorted.slice(5).reduce((s,[,v])=>s+v,0);const entries=[...top5];if(resto>0)entries.push(['Demais clientes',resto]);
  const bgColors=entries.map(([k])=>k==='Demais clientes'?'rgba(100,116,139,.55)':color);dch(key);if(!entries.length)return;
  CHARTS[key]=mkChart(canvId,'bar',{labels:entries.map(([k])=>k.substring(0,28)),datasets:[{data:entries.map(([,v])=>v),backgroundColor:bgColors,borderRadius:4,borderSkipped:false}]},{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>{const p=total>0?(c.raw/total*100).toFixed(1):'0';return' '+R$(c.raw)+' · '+p+'% do total';}}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}},y:{grid:{display:false},ticks:{color:'#7c93b0',font:{size:10}}}}});
}

// ══════════════════════════════════════════════════════════
//  RENDER FILIAL — com SC
// ══════════════════════════════════════════════════════════
const FILIAL_COLORS=[{color:'var(--accent)',bar:'rgba(249,115,22,.85)'},{color:'var(--pet)',bar:'rgba(16,185,129,.85)'},{color:'var(--cyan)',bar:'rgba(6,182,212,.85)'},{color:'var(--warning)',bar:'rgba(245,158,11,.85)'},{color:'#a78bfa',bar:'rgba(139,92,246,.85)'}];
function filialColorIdx(filial){return STATE.filiais.indexOf(filial)%FILIAL_COLORS.length;}

function renderFilial(filial){
  const pfx=filial.toLowerCase();const ci=filialColorIdx(filial);const color=FILIAL_COLORS[ci].color;const barCol=FILIAL_COLORS[ci].bar;
  const{cur,prev,sc}=_data;
  const nodata=$(`${pfx}-nodata`);const dataEl=$(`${pfx}-data`);const has=cur.some(a=>a.filial===filial)||( sc&&sc.some(a=>a.filial===filial));
  if(nodata)nodata.style.display=has?'none':'block';if(dataEl)dataEl.style.display=has?'block':'none';if(!has)return;
  const sCur=stats(cur,filial);const sPrev=stats(prev,filial);const scF=scStats(sc,filial);
  const dmCur=dmap(cur,filial,sc);const dmPrev=dmap(prev,filial,null);
  const today=new Date().toISOString().split('T')[0];const elapsedDays=_wdC.filter(d=>d<=today).length;
  const proj=elapsedDays>0?(sCur.total/elapsedDays)*_wdC.length:0;const med=elapsedDays>0?sCur.total/elapsedDays:0;
  const pVal=proRata(sCur.total,sPrev.total,_wdC,_wdP);const projPct=sPrev.total>0?(proj/sPrev.total*100):0;const prog=sPrev.total>0?Math.min(sCur.total/sPrev.total*100,100):0;

  if($(`lbl-${pfx}`))$(`lbl-${pfx}`).textContent=`${MONTHS[_mo]} ${_curY}`;
  // Status de freshness por source da filial
  const freshEl=$(`fresh-${pfx}`);
  if(freshEl&&STATE.freshness){
    const info=STATE.freshness.byFilial[filial];
    if(info){
      const items=[];
      const srcLabels={carteira:'Carteira',extra:'Extra',sc:'SC'};
      ['carteira','extra'].concat(info.hasSC?['sc']:[]).forEach(s=>{
        const last=info.sources[s];
        if(!last){items.push(`<span class="filial-fresh-item late">${srcLabels[s]}: sem arquivo</span>`);return;}
        const gap=diffWorkingDays(last,STATE.freshness.expected);
        const cls=gap===0?'ok':gap===1?'warn':'late';
        const ic=cls==='ok'?ICONS.check:cls==='warn'?ICONS.warn:ICONS.alert;
        items.push(`<span class="filial-fresh-item ${cls}">${ic}<span>${srcLabels[s]}: ${fd(last)}${gap>0?` (-${gap}d)`:''}</span></span>`);
      });
      freshEl.innerHTML=items.join('');
    }
  }

  // KPIs
  set(`${pfx}-tot`,R$(sCur.total),`color:${color}`);bb(`${pfx}-badge`,pct(pVal)+` vs ${_prevY}`,bcl(pVal));
  set(`${pfx}-cart`,R$(sCur.carteira));if($(`${pfx}-cart-sub`))$(`${pfx}-cart-sub`).textContent=`${_prevY}: ${R$(sPrev.carteira)}`;
  set(`${pfx}-extra`,R$(sCur.extra));if($(`${pfx}-extra-sub`))$(`${pfx}-extra-sub`).textContent=`${_prevY}: ${R$(sPrev.extra)}`;
  // SC KPI
  set(`${pfx}-sc-kpi`,R$(scF.total),'color:var(--sc)');
  if($(`${pfx}-sc-kpi-sub`))$(`${pfx}-sc-kpi-sub`).innerHTML=scF.awbs>0?`${scF.awbs} AWBs`:'Sem SC neste mês';
  set(`${pfx}-avg`,R$(med));if($(`${pfx}-avg-sub`))$(`${pfx}-avg-sub`).textContent=`${elapsedDays} dias úteis dec. / ${_wdC.length} úteis no mês`;
  set(`${pfx}-proj`,R$(proj));bb(`${pfx}-proj-b`,pct(projPct-100)+' vs meta',bcl(projPct-100));
  set(`${pfx}-awbs`,sCur.awbs.toLocaleString('pt-BR'));if($(`${pfx}-awbs-sub`))$(`${pfx}-awbs-sub`).textContent=`${_prevY}: ${sPrev.awbs.toLocaleString('pt-BR')} · Aéreo ${sCur.awbsAereo} · Rod. ${sCur.awbsRodoviario}`;
  if($(`${pfx}-prog-bar`))$(`${pfx}-prog-bar`).style.width=prog.toFixed(1)+'%';
  if($(`${pfx}-prog-pct`))$(`${pfx}-prog-pct`).textContent=prog.toFixed(1)+`% do total ${_prevY}`;
  if($(`${pfx}-prog-meta`))$(`${pfx}-prog-meta`).textContent='Meta: '+R$(sPrev.total);

  // Tabela diária com coluna SC
  const tbody=$(`body-${pfx}`);
  if(tbody){
    tbody.innerHTML='';
    _wdC.forEach((dt,i)=>{
      if(dt>today&&!dmCur[dt])return;
      const r=dmCur[dt];const dp=_wdP[i];const v25=dp?(dmPrev[dp]?.total||0):0;const tot=r?.total||0;
      const scDia=r?.sc||0;const scDiaAwbs=r?.scAwbs||0;
      const diff=v25>0?((tot-v25)/v25*100):null;
      const isScDay=scDia>0;
      const tr=document.createElement('tr');tr.className='row-clickable';tr.title='Ver clientes novos/retorno';tr.onclick=()=>openNovos(dt,filial);
      tr.innerHTML=`
        <td class="mono"><strong>${fd(dt)}</strong></td>
        <td style="color:var(--text2)">${dn(dt)}</td>
        <td>${r?R$(r.carteira):'—'}</td>
        <td>${r?R$(r.extra):'—'}</td>
        <td class="col-sc${false?' col-sc-new':''}" style="cursor:${isScDay?'pointer':'default'};font-weight:${isScDay?'700':'400'};color:${isScDay?'var(--sc)':'var(--text3)'}" onclick="${isScDay?`event.stopPropagation();openSCDay('${dt}','${filial}')`:'void(0)'}" title="${isScDay?`SC do dia: ${scDiaAwbs} AWBs (${0} novos)`:'Sem SC neste dia'}">${scDia>0?R$(scDia)+'<br><span style="font-size:10px;color:var(--text3)">'+scDiaAwbs+' AWBs</span>':'—'}</td>
        <td class="mono" style="font-weight:700;color:${tot>0?color:'var(--text3)'}">${tot>0?R$(tot):'—'}</td>
        <td>${r?.awbs||'—'}</td><td>${r?.volumes||'—'}</td>
        <td style="color:var(--text3)">${v25>0?R$(v25):'—'}</td>
        <td style="font-weight:700;color:${diff===null?'var(--text3)':diff>=0?'var(--success)':'var(--danger)'}">${diff!==null?pct(diff):'—'}</td>
        <td style="color:var(--text3);font-size:11px;text-align:center">+</td>
        <td style="text-align:center">${false?`<span onclick="event.stopPropagation();openSCDay('${dt}','${filial}')" style="cursor:pointer;background:var(--sc-bg);color:var(--sc);padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700">${0}↑</span>`:'—'}</td>`;
      tbody.appendChild(tr);
    });
  }

  // Gráfico diário com SC
  const labels=_wdC.map(d=>fd(d));const vC=_wdC.map(d=>dmCur[d]?.total||0);const vP=_wdC.map((_,i)=>{const dp=_wdP[i];return dp?(dmPrev[dp]?.total||0):0;});const vSC=_wdC.map(d=>dmCur[d]?.sc||0);
  const dkD=`ch_${pfx}_daily`;dch(dkD);
  if($(`ch-${pfx}-daily`))CHARTS[dkD]=mkChart(`ch-${pfx}-daily`,'bar',{labels,datasets:[{label:String(_curY),data:vC,backgroundColor:barCol,borderRadius:3,stack:'a'},{label:'SC',data:vSC,backgroundColor:'rgba(167,139,250,.55)',borderRadius:3,stack:'sc'},{label:String(_prevY),data:vP,type:'line',borderColor:'rgba(6,182,212,.7)',backgroundColor:'transparent',tension:.3,pointRadius:1.5,borderWidth:1.5}]},co({stacked:false}));

  // Tipo frete e modal (igual antes)
  const tf={};cur.filter(a=>a.filial===filial).forEach(a=>{const t=(String(a.tipo_frete||'').trim().toUpperCase())||'N/D';tf[t]=(tf[t]||0)+a.valor_frete;});
  const lbTf=Object.keys(tf).filter(k=>tf[k]>0).sort((a,b)=>tf[b]-tf[a]);const colorsTf=['rgba(6,182,212,.85)','rgba(249,115,22,.85)','rgba(139,92,246,.85)','rgba(16,185,129,.85)','rgba(148,163,184,.85)'];
  const dkT=`ch_${pfx}_tipo`;dch(dkT);if($(`ch-${pfx}-tipo`))CHARTS[dkT]=mkChart(`ch-${pfx}-tipo`,'bar',{labels:lbTf,datasets:[{data:lbTf.map(k=>tf[k]),backgroundColor:lbTf.map((_,i)=>colorsTf[i%colorsTf.length]),borderRadius:4,borderSkipped:false}]},{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+R$(c.raw)+' ('+(c.raw/(lbTf.reduce((s,k)=>s+tf[k],0)||1)*100).toFixed(1)+'%)'}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}},y:{grid:{display:false},ticks:{color:'#7c93b0',font:{size:11},padding:4}}}});
  const mm={'Rodoviário':0,'Aéreo':0,'N/D':0};cur.filter(a=>a.filial===filial).forEach(a=>{const t=a.modal||'N/D';if(t==='Rodoviário'||t==='Aéreo')mm[t]+=a.valor_frete;else mm['N/D']+=a.valor_frete;});
  const lbMm=Object.keys(mm).filter(k=>mm[k]>0).sort((a,b)=>mm[b]-mm[a]);const colMm={'Rodoviário':'rgba(245,158,11,.85)','Aéreo':'rgba(59,130,246,.85)','N/D':'rgba(148,163,184,.85)'};
  const dkM=`ch_${pfx}_modal`;dch(dkM);if($(`ch-${pfx}-modal`))CHARTS[dkM]=mkChart(`ch-${pfx}-modal`,'bar',{labels:lbMm,datasets:[{data:lbMm.map(k=>mm[k]),backgroundColor:lbMm.map(k=>colMm[k]||'rgba(148,163,184,.85)'),borderRadius:4,borderSkipped:false}]},{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+R$(c.raw)+' ('+(c.raw/(lbMm.reduce((s,k)=>s+mm[k],0)||1)*100).toFixed(1)+'%)'}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}},y:{grid:{display:false},ticks:{color:'#7c93b0',font:{size:11},padding:4}}}});
  const dkCli=`ch_${pfx}_cli`;if($(`ch-${pfx}-clientes`))topClientes(`ch-${pfx}-clientes`,dkCli,filial,barCol,cur);
}

// ══════════════════════════════════════════════════════════
//  MODAL SC DO DIA
// ══════════════════════════════════════════════════════════
let _scDayData=[];
function openSCDay(dateStr,filial){
  const{sc}=_data;if(!sc||!sc.length)return;
  _scDayData=sc.filter(a=>a.filial===filial&&a.data===dateStr);
  if(!_scDayData.length){toast('Sem SC neste dia.','ok');return;}
  const totalSC=_scDayData.reduce((s,a)=>s+a.valor_frete,0);
  $('sc-day-title').textContent=`SC — ${filial} · ${fd(dateStr)}`;
  $('sc-day-sub').textContent=`${_scDayData.length} AWBs · Total ${R$(totalSC)}`;
  $('sc-day-count').textContent='Sem Comissionado';
  const tbody=$('body-sc-day');tbody.innerHTML='';
  _scDayData.sort((a,b)=>b.valor_frete-a.valor_frete).forEach(a=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td class="mono" style="font-size:11px">${a.awb}</td><td>${new Date(a.data+'T12:00:00').toLocaleDateString('pt-BR')}</td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${a.remetente}</td><td style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${a.destinatario}</td><td>${a.cidade}/${a.uf}</td><td><span class="tag ${(a.tipo_frete||'nd').toLowerCase()}">${a.tipo_frete||'N/D'}</span></td><td><span class="tag ${modalClass(a.modal)}">${a.modal}</span></td><td style="text-align:right;font-family:var(--mono);font-weight:700;color:var(--sc)">${R$(a.valor_frete)}</td><td style="color:var(--text3);font-size:11px">SC</td>` ;
    tbody.appendChild(tr);
  });
  $('modal-sc-day').style.display='flex';document.body.style.overflow='hidden';
}

function exportSCDayCSV(){let csv='AWB,Data,Remetente,Destinatário,Cidade,UF,Tipo Frete,Modal,Frete R$,Status\n';_scDayData.forEach(a=>{csv+=`"${a.awb}","${a.data}","${a.remetente}","${a.destinatario}","${a.cidade}","${a.uf}","${a.tipo_frete}","${a.modal}",${a.valor_frete},"SC"\n`;});const el=document.createElement('a');el.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));el.download=`SC_dia.csv`;el.click();}

function exportSCCSV(filial){const{sc}=_data;if(!sc)return;const scF=sc.filter(a=>a.filial===filial);let csv='AWB,Data,Filial,Remetente,Destinatário,Cidade,UF,Tipo Frete,Modal,Frete R$,Status\n';scF.forEach(a=>{csv+=`"${a.awb}","${a.data}","${a.filial}","${a.remetente}","${a.destinatario}","${a.cidade}","${a.uf}","${a.tipo_frete}","${a.modal}",${a.valor_frete},"SC"\n`;});const el=document.createElement('a');el.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));el.download=`SC_${filial}_${_curY}-${padM(_mo)}.csv`;el.click();}

// ══════════════════════════════════════════════════════════
//  INTELIGÊNCIA ANUAL — inclui SC
// ══════════════════════════════════════════════════════════
function renderIntel(){
  const proj=STATE.projection;const hasData=proj&&(proj.consolidated||Object.keys(proj.byFilial||{}).length>0);
  const emptyEl=$('intel-empty');const contentEl=$('intel-content');
  if(!hasData){if(emptyEl)emptyEl.style.display='block';if(contentEl)contentEl.style.display='none';return;}
  if(emptyEl)emptyEl.style.display='none';if(contentEl)contentEl.style.display='block';
  const cons=proj.consolidated;
  if(cons){
    set('intel-proj-total',R$(cons.projectedTotal));set('intel-realizado',R$(cons.realizadoAteHoje));set('intel-dias-real',Object.values(proj.byFilial).reduce((s,v)=>Math.max(s,v.diasRealizados||0),0)+' dias úteis');
    // SC total anual
    const scAnual=Object.values(proj.byFilial).reduce((s,v)=>s+(v.scPotencial||0),0);if($('intel-sc-total'))set('intel-sc-total',scAnual>0?R$(scAnual):'—');
    const confLabel={conservador:'Conservador',realista:'Realista',agressivo:'Agressivo',historico:'Referência Histórica'};const confColor={conservador:'rgba(245,158,11,.18)',realista:'rgba(16,185,129,.18)',agressivo:'rgba(239,68,68,.18)',historico:'rgba(100,116,139,.18)'};const confText={conservador:'var(--warning)',realista:'var(--success)',agressivo:'var(--danger)',historico:'var(--text2)'};
    const cEl=$('intel-confidence-badge');if(cEl){cEl.textContent=confLabel[cons.confidence]||cons.confidence;cEl.style.background=confColor[cons.confidence]||confColor.realista;cEl.style.color=confText[cons.confidence]||confText.realista;}
    const subEl=$('intel-proj-sub');if(subEl){const curY=decodePeriod(STATE.periods[0]).year;subEl.textContent=`Estimativa para o ano completo de ${curY}`;}
  }
  const grid=$('intel-filiais-grid');
  if(grid){
    grid.innerHTML='';const colors=FILIAL_COLORS;
    STATE.filiais.forEach((filial,i)=>{
      const fp=proj.byFilial[filial];if(!fp)return;const col=colors[i%colors.length];const hist=STATE.history[filial];
      const histRows=fp.vsBestYear?`<div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-top:1px solid var(--border)"><span style="font-size:11px;color:var(--text2)">🏆 Melhor ano (${hist?.bestYear?.year||'—'})</span><span style="font-family:var(--mono);font-size:12px">${R$(fp.vsBestYear)}</span></div><div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0"><span style="font-size:11px;color:var(--text2)">Δ vs melhor ano</span><span class="badge ${fp.vsBestYearPct>=0?'up':'dn'}">${pct(fp.vsBestYearPct)}</span></div>`:'';
      const scRow=fp.scPotencial>0?`<div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-top:1px solid var(--sc-border);background:var(--sc-bg);border-radius:6px;padding:8px 12px;margin-top:6px"><span style="font-size:11px;color:var(--sc);font-weight:700">⚠ SC a recuperar este mês</span><span style="font-family:var(--mono);font-size:13px;font-weight:800;color:var(--sc)">${R$(fp.scPotencial)}</span></div>`:'';
      grid.innerHTML+=`<div class="chart-card" style="border-top:3px solid ${col.bar};padding:18px 20px">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px"><div style="font-weight:800;font-size:14px">${filial}</div><div style="font-size:11px;padding:3px 10px;border-radius:12px;font-weight:700;background:${fp.confidence==='agressivo'?'rgba(239,68,68,.15)':fp.confidence==='conservador'?'rgba(245,158,11,.15)':'rgba(16,185,129,.15)'};color:${fp.confidence==='agressivo'?'var(--danger)':fp.confidence==='conservador'?'var(--warning)':'var(--success)'}">${fp.confidence}</div></div>
        <div style="font-size:26px;font-weight:800;font-family:var(--mono);color:${col.color};margin-bottom:4px">${R$(fp.projectedTotal)}</div>
        <div style="font-size:11px;color:var(--text3);margin-bottom:14px">Projeção para o ano completo</div>
        <div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-top:1px solid var(--border)"><span style="font-size:11px;color:var(--text2)">✅ Realizado até ontem</span><span style="font-family:var(--mono);font-size:12px;color:var(--success)">${R$(fp.realizadoAteHoje)}</span></div>
        <div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0"><span style="font-size:11px;color:var(--text2)">⚡ Média diária</span><span style="font-family:var(--mono);font-size:12px">${R$(fp.dailyRate)}/dia</span></div>
        <div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0"><span style="font-size:11px;color:var(--text2)">📅 Dias realizados</span><span style="font-family:var(--mono);font-size:12px">${fp.diasRealizados} úteis</span></div>
        ${histRows}${scRow}
      </div>`;
    });
  }

  // Seção SC anual
  renderIntelSCSection();

  const histSection=$('intel-hist-section');const histGrid=$('intel-hist-grid');const hasHistory=Object.keys(STATE.history||{}).length>0;
  if(histSection)histSection.style.display=hasHistory?'block':'none';
  if(histGrid&&hasHistory){histGrid.innerHTML='';STATE.filiais.forEach((filial,i)=>{const h=STATE.history[filial];if(!h)return;const col=FILIAL_COLORS[i%FILIAL_COLORS.length];const topRows=h.topMonths.map((m,idx)=>`<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-top:1px solid var(--border)"><span style="font-size:11px;color:var(--text2)">${idx===0?'🥇':idx===1?'🥈':idx===2?'🥉':'  '} ${MONTHS[m.month]} ${m.year}</span><span style="font-family:var(--mono);font-size:12px;font-weight:700;color:${col.color}">${R$(m.total)}</span></div>`).join('');histGrid.innerHTML+=`<div class="chart-card" style="border-top:3px solid ${col.bar}"><div style="font-weight:800;font-size:13px;margin-bottom:12px">${filial} — Top Meses Históricos</div><div style="display:flex;justify-content:space-between;padding:4px 0;margin-bottom:4px"><span style="font-size:11px;color:var(--text2)">Melhor ano histórico</span><span style="font-family:var(--mono);font-size:12px;font-weight:700">${h.bestYear?h.bestYear.year+' · '+R$(h.bestYear.total):'—'}</span></div><div style="display:flex;justify-content:space-between;padding:4px 0;margin-bottom:8px"><span style="font-size:11px;color:var(--text2)">Média mensal histórica</span><span style="font-family:var(--mono);font-size:12px">${R$(h.avgMonthly)}</span></div>${topRows}</div>`;});}

  const seasonSection=$('intel-season-section');const hasSeasonData=STATE.filiais.some(f=>STATE.history[f]?.seasonality);if(seasonSection)seasonSection.style.display=hasSeasonData?'block':'none';
  if(hasSeasonData&&$('ch-intel-season')){dch('chIntelSeason');const labels=MONTHS.slice(1);const datasets=STATE.filiais.map((filial,i)=>{const h=STATE.history[filial];if(!h)return null;const col=FILIAL_COLORS[i%FILIAL_COLORS.length];return{label:filial,data:Array.from({length:12},(_,mo)=>{const v=h.seasonality[mo+1];return v!==undefined?Math.round(v*100)/100:null;}),borderColor:col.bar,backgroundColor:'transparent',tension:0.35,pointRadius:4,borderWidth:2};}).filter(Boolean);datasets.push({label:'Média (1.0)',data:new Array(12).fill(1),borderColor:'rgba(100,116,139,.4)',borderDash:[6,4],backgroundColor:'transparent',pointRadius:0,borderWidth:1.5});CHARTS.chIntelSeason=mkChart('ch-intel-season','line',{labels,datasets},{responsive:true,plugins:{legend:{labels:{color:'#7c93b0',font:{size:11}}},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${c.raw?.toFixed(2)??'—'}×`}}},scales:{x:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10}}},y:{grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>v.toFixed(1)+'×'},min:0}}});}
  renderIntelFreteAnalysis();renderIntelTopClientes();
}

// Seção SC na análise anual
function renderIntelSCSection(){
  const section=$('intel-sc-section');const grid=$('intel-sc-grid');if(!section||!grid)return;
  // Coleta SC de todos os períodos do ano atual
  const scByFilial={};STATE.periods.forEach(p=>{const d=STATE.cache[p];if(!d||!d.sc)return;const{month}=decodePeriod(p);d.sc.forEach(a=>{if(!scByFilial[a.filial])scByFilial[a.filial]={byDay:{},total:0,awbs:0,months:new Set()};scByFilial[a.filial].total+=a.valor_frete;scByFilial[a.filial].awbs++;scByFilial[a.filial].months.add(month);if(!scByFilial[a.filial].byDay[a.data])scByFilial[a.filial].byDay[a.data]={total:0,awbs:0};scByFilial[a.filial].byDay[a.data].total+=a.valor_frete;scByFilial[a.filial].byDay[a.data].awbs++;});});
  const hasAnySC=Object.keys(scByFilial).length>0;section.style.display=hasAnySC?'block':'none';if(!hasAnySC)return;
  grid.innerHTML='';
  Object.entries(scByFilial).forEach(([filial,s])=>{
    const fi=STATE.filiais.indexOf(filial);const col=FILIAL_COLORS[fi%FILIAL_COLORS.length];const pctMig=s.total>0?(s.migrado/s.total*100):0;const pctPend=100-pctMig;
    const topDays=Object.entries(s.byDay).sort((a,b)=>b[1].awbs-a[1].awbs).slice(0,5);
    grid.innerHTML+=`<div class="chart-card" style="border-top:3px solid var(--sc)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px"><div style="font-weight:800;font-size:13px;color:var(--sc)">${filial} — SC Acumulado</div><span class="tag sc">${MONTHS[_mo]} ${_curY}</span></div>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:14px">
        <div style="background:var(--card2);border:1px solid var(--border);border-radius:8px;padding:10px"><div style="font-size:10px;color:var(--text3);margin-bottom:4px">Total SC</div><div style="font-size:16px;font-weight:800;font-family:var(--mono);color:var(--sc)">${R$(s.total)}</div><div style="font-size:10px;color:var(--text3)">${s.awbs} AWBs</div></div>
        <div style="background:rgba(16,185,129,.08);border:1px solid rgba(16,185,129,.2);border-radius:8px;padding:10px"><div style="font-size:10px;color:var(--text3);margin-bottom:4px">✓ Migrou</div><div style="font-size:16px;font-weight:800;font-family:var(--mono);color:var(--success)">${R$(s.migrado)}</div><div style="font-size:10px;color:var(--text3)">${pctMig.toFixed(1)}%</div></div>
        <div style="background:var(--sc-bg);border:1px solid var(--sc-border);border-radius:8px;padding:10px"><div style="font-size:10px;color:var(--text3);margin-bottom:4px">⚠ A recuperar</div><div style="font-size:16px;font-weight:800;font-family:var(--mono);color:var(--warning)">${R$(s.aRecuperar)}</div><div style="font-size:10px;color:var(--text3)">${pctPend.toFixed(1)}%</div></div>
      </div>
      <div style="margin-bottom:10px"><div style="font-size:10px;color:var(--text3);margin-bottom:5px">Conversão para comissão</div><div style="height:8px;background:var(--border);border-radius:6px;overflow:hidden;display:flex"><div style="height:100%;width:${pctMig.toFixed(1)}%;background:var(--success);border-radius:6px 0 0 6px;transition:width .6s"></div><div style="height:100%;width:${pctPend.toFixed(1)}%;background:rgba(167,139,250,.5)"></div></div><div style="display:flex;justify-content:space-between;margin-top:4px;font-size:10px;color:var(--text3)"><span>✓ Comissionado ${pctMig.toFixed(1)}%</span><span>⚠ Pendente ${pctPend.toFixed(1)}%</span></div></div>
      ${topDays.length?`<div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.7px;margin-bottom:6px">Dias com mais SC</div>${topDays.map(([dt,v])=>`<div style="display:flex;justify-content:space-between;padding:4px 0;border-top:1px solid rgba(30,51,82,.4);font-size:11px"><span style="color:var(--text2)">${new Date(dt+'T12:00:00').toLocaleDateString('pt-BR')}</span><span>${v.awbs} AWBs · ${R$(v.total)}</span></div>`).join('')}`:''}
    </div>`;
  });
}

// ══════════════════════════════════════════════════════════
//  ANÁLISE POR TIPO DE FRETE
// ══════════════════════════════════════════════════════════
function renderIntelFreteAnalysis(){
  const grid=$('intel-frete-grid');if(!grid)return;grid.innerHTML='';
  STATE.filiais.forEach((filial,fi)=>{
    const col=FILIAL_COLORS[fi%FILIAL_COLORS.length];const pfx=filial.toLowerCase();const awbs=[];
    STATE.periods.forEach(p=>{const d=STATE.cache[p];if(!d)return;d.cur.filter(a=>a.filial===filial).forEach(a=>awbs.push(a));});
    if(!awbs.length)return;
    const totCIF=awbs.filter(a=>(a.tipo_frete||'').toUpperCase()==='CIF').reduce((s,a)=>s+a.valor_frete,0);const totFOB=awbs.filter(a=>(a.tipo_frete||'').toUpperCase()==='FOB').reduce((s,a)=>s+a.valor_frete,0);const totAereo=awbs.filter(a=>(a.modal||'').toLowerCase().includes('aér')||(a.modal||'').toLowerCase().includes('aer')).reduce((s,a)=>s+a.valor_frete,0);const totRodo=awbs.filter(a=>(a.modal||'').toLowerCase().includes('rod')).reduce((s,a)=>s+a.valor_frete,0);const totGeral=awbs.reduce((s,a)=>s+a.valor_frete,0);
    const pctCIF=totGeral>0?(totCIF/totGeral*100):0;const pctFOB=totGeral>0?(totFOB/totGeral*100):0;const pctAereo=totGeral>0?(totAereo/totGeral*100):0;const pctRodo=totGeral>0?(totRodo/totGeral*100):0;
    const awbsCIF=awbs.filter(a=>(a.tipo_frete||'').toUpperCase()==='CIF').length;const awbsFOB=awbs.filter(a=>(a.tipo_frete||'').toUpperCase()==='FOB').length;const awbsAereo=awbs.filter(a=>(a.modal||'').toLowerCase().includes('aér')||(a.modal||'').toLowerCase().includes('aer')).length;const awbsRodo=awbs.filter(a=>(a.modal||'').toLowerCase().includes('rod')).length;
    const card=document.createElement('div');card.className='chart-card';card.style.cssText=`border-top:3px solid ${col.bar}`;
    card.innerHTML=`<div class="chart-ttl" style="margin-bottom:16px"><span style="color:${col.color};font-size:13px;font-weight:800">${filial}</span><span style="font-weight:400;color:var(--text3)">Acumulado ano atual</span></div><div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px">${[{lbl:'CIF',val:totCIF,pctV:pctCIF,awbs:awbsCIF,color:'var(--cyan)'},{lbl:'FOB',val:totFOB,pctV:pctFOB,awbs:awbsFOB,color:'var(--accent)'},{lbl:'Aéreo',val:totAereo,pctV:pctAereo,awbs:awbsAereo,color:'#60a5fa'},{lbl:'Rodoviário',val:totRodo,pctV:pctRodo,awbs:awbsRodo,color:'#fbbf24'}].map(item=>`<div style="background:var(--card2);border:1px solid var(--border);border-radius:10px;padding:10px 12px"><div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.7px;margin-bottom:6px">${item.lbl}</div><div style="font-size:16px;font-weight:800;font-family:var(--mono);color:${item.color};margin-bottom:2px">${R$(item.val)}</div><div style="font-size:11px;color:var(--text3)">${item.pctV.toFixed(1)}% · ${item.awbs} AWBs</div><div style="margin-top:6px;height:4px;background:var(--border);border-radius:4px;overflow:hidden"><div style="height:100%;width:${Math.min(item.pctV,100).toFixed(1)}%;background:${item.color};border-radius:4px"></div></div></div>`).join('')}</div><div style="height:140px;position:relative"><canvas id="ch-intel-frete-${pfx}"></canvas></div>`;
    grid.appendChild(card);
    requestAnimationFrame(()=>{const ctx=document.getElementById(`ch-intel-frete-${pfx}`);if(!ctx)return;if(ctx._chartInst)ctx._chartInst.destroy();ctx._chartInst=new Chart(ctx,{type:'doughnut',data:{labels:['CIF','FOB','Aéreo','Rodoviário'],datasets:[{data:[totCIF,totFOB,totAereo,totRodo],backgroundColor:['rgba(6,182,212,.7)','rgba(249,115,22,.7)','rgba(96,165,250,.7)','rgba(251,191,36,.7)'],borderColor:['#06b6d4','#f97316','#60a5fa','#fbbf24'],borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'right',labels:{color:'#7c93b0',font:{size:10},boxWidth:10}},tooltip:{callbacks:{label:ctx=>{const t=ctx.dataset.data.reduce((s,v)=>s+v,0);return` ${ctx.label}: ${R$(ctx.raw)} (${t>0?(ctx.raw/t*100).toFixed(1):0}%)`;}}}}}}); });
  });
}

// ══════════════════════════════════════════════════════════
//  TOP CLIENTES ANUAL
// ══════════════════════════════════════════════════════════
function renderIntelTopClientes(){
  const grid=$('intel-clientes-grid');if(!grid)return;grid.innerHTML='';
  STATE.filiais.forEach((filial,fi)=>{
    const col=FILIAL_COLORS[fi%FILIAL_COLORS.length];const clientMap={};
    STATE.periods.forEach(p=>{const d=STATE.cache[p];if(!d)return;d.cur.filter(a=>a.filial===filial).forEach(a=>{const key=a.cnpj||a.remetente||'—';const nome=a.remetente||key;if(!clientMap[key])clientMap[key]={nome,cnpj:a.cnpj||'',total:0,awbs:0};clientMap[key].total+=a.valor_frete;clientMap[key].awbs++;});});
    const top=Object.values(clientMap).sort((a,b)=>b.total-a.total).slice(0,10);if(!top.length)return;
    const totalFilial=top.reduce((s,c)=>s+c.total,0);const maxVal=top[0]?.total||1;
    const card=document.createElement('div');card.className='tbl-card';card.style.cssText=`border-top:3px solid ${col.bar}`;
    card.innerHTML=`<div class="tbl-top"><div class="tbl-top-ttl" style="color:${col.color}">${filial} — Top 10 Clientes</div><span style="font-size:11px;color:var(--text3);font-family:var(--mono)">Acumulado ano atual</span></div><div class="tbl-wrap"><table><thead><tr><th style="width:28px">#</th><th>Cliente</th><th>CNPJ</th><th style="text-align:right">Frete R$</th><th style="text-align:right">AWBs</th><th style="text-align:right">% do total</th><th style="min-width:80px"></th></tr></thead><tbody>${top.map((c,i)=>{const p=totalFilial>0?(c.total/totalFilial*100):0;const barW=maxVal>0?(c.total/maxVal*100).toFixed(1):0;return`<tr><td style="color:var(--text3);font-family:var(--mono)">${i+1}</td><td style="font-weight:700;max-width:180px;overflow:hidden;text-overflow:ellipsis" title="${c.nome}">${c.nome}</td><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">${c.cnpj||'—'}</td><td style="text-align:right;font-family:var(--mono);font-weight:700;color:${col.color}">${R$(c.total)}</td><td style="text-align:right;font-family:var(--mono)">${c.awbs}</td><td style="text-align:right;font-family:var(--mono);color:var(--text2)">${p.toFixed(1)}%</td><td><div style="height:5px;background:var(--border);border-radius:3px;overflow:hidden"><div style="height:100%;width:${barW}%;background:${col.bar};border-radius:3px"></div></div></td></tr>`;}).join('')}</tbody></table></div>`;
    grid.appendChild(card);
  });
}

// ══════════════════════════════════════════════════════════
//  AWB SEARCH — inclui SC
// ══════════════════════════════════════════════════════════
let awbAll=[],awbPg=1;const PP=50;
function initAWB(){
  const all=[];Object.values(STATE.cache).forEach(d=>{all.push(...d.cur,...d.prev,...(d.sc||[]));});
  if(!all.length&&_data){all.push(..._data.cur,..._data.prev,...(_data.sc||[]));}
  awbAll=[...all];awbPg=1;updC();renderAWBPg();
}
function fAWB(){
  const q=$('aq').value.trim().toLowerCase();const r=$('ar').value.trim().toLowerCase();const dest=($('adest')?.value||'').trim().toLowerCase();const cidade=($('acidade')?.value||'').trim().toLowerCase();const f=$('af').value,p=$('ap').value,s=$('as').value;const modal=($('amodal')?.value||'');const dt1=($('adt1')?.value||''),dt2=($('adt2')?.value||'');const fmin=parseFloat($('afmin')?.value||0)||0;
  let base=[];Object.values(STATE.cache).forEach(d=>{base.push(...d.cur,...d.prev,...(d.sc||[]));});if(!base.length&&_data)base.push(..._data.cur,..._data.prev,...(_data.sc||[]));
  awbAll=base.filter(a=>(!q||a.awb.toLowerCase().includes(q))&&(!r||a.remetente.toLowerCase().includes(r))&&(!dest||a.destinatario.toLowerCase().includes(dest))&&(!cidade||(a.cidade+'/'+a.uf).toLowerCase().includes(cidade))&&(!f||a.filial===f)&&(!p||a.periodo===p||(p&&`${a.year}-${padM(new Date(a.data+'T12:00:00').getMonth()+1)}`===p))&&(!s||a.source===s)&&(!modal||a.modal===modal)&&(!dt1||a.data>=dt1)&&(!dt2||a.data<=dt2)&&(!fmin||a.valor_frete>=fmin));
  awbPg=1;updC();renderAWBPg();
}
function updC(){$('awb-count').textContent=awbAll.length.toLocaleString('pt-BR')+' registros';}
function renderAWBPg(){
  const tb=$('body-awb');tb.innerHTML='';const st=(awbPg-1)*PP;
  awbAll.slice(st,st+PP).forEach(a=>{
    const tr=document.createElement('tr');const tc=(a.tipo_frete||'').toLowerCase();const{month:am}={month:new Date(a.data+'T12:00:00').getMonth()+1};
    const isSC=a.source==='sc';
    tr.style.background=isSC?'rgba(167,139,250,.04)':'';
    tr.innerHTML=`<td class="mono" style="font-size:11px">${a.awb}</td><td>${new Date(a.data+'T12:00:00').toLocaleDateString('pt-BR')}</td><td><span class="tag" style="background:rgba(249,115,22,.12);color:var(--accent)">${a.filial}</span></td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${a.remetente}</td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${a.destinatario}</td><td>${a.cidade}/<strong>${a.uf}</strong></td><td>${a.volumes}</td><td class="mono">${a.peso} kg</td><td>${R$(a.valor_mercantil)}</td><td class="mono" style="font-weight:700;color:${isSC?'var(--sc)':'var(--accent)'}">${R$(a.valor_frete)}</td><td><span class="tag ${tc}">${a.tipo_frete}</span></td><td><span class="tag ${modalClass(a.modal)}">${a.modal||'N/D'}</span></td><td><span class="tag ${a.source}">${a.source}</span></td><td style="color:var(--text3);font-family:var(--mono)">${MONTHS[am].substring(0,3)} ${a.year}</td>`;
    tb.appendChild(tr);
  });
  renderPag('pg-awb',awbAll.length,awbPg,PP,p=>{awbPg=p;renderAWBPg();});
}
function cAWB(){['aq','ar','adest','acidade','afmin'].forEach(id=>{const el=$(id);if(el)el.value='';});['af','ap','as','amodal'].forEach(id=>{const el=$(id);if(el)el.value='';});['adt1','adt2'].forEach(id=>{const el=$(id);if(el)el.value='';});fAWB();}

// ══════════════════════════════════════════════════════════
//  ÍNDICE GLOBAL DE REMETENTES
// ══════════════════════════════════════════════════════════
let ALL_REM_DATES={};function normRem(s){return String(s||'').trim().toLowerCase().replace(/\s+/g,' ');}
function freightPayerInfo(a){const tipo=String(a?.tipo_frete||'').trim().toUpperCase();if(tipo==='FOB'){const nome=String(a?.destinatario||'').trim();const cnpj=String(a?.cnpj_destinatario||'').replace(/\D/g,'').replace(/^0+/,'');return{nome,cnpj,tipo,papel:'Destinatário'};}if(tipo==='CIF'){const nome=String(a?.remetente||'').trim();const cnpj=String(a?.cnpj_remetente||a?.cnpj||'').replace(/\D/g,'').replace(/^0+/,'');return{nome,cnpj,tipo,papel:'Remetente'};}return{nome:'',cnpj:'',tipo,papel:''};}
function rebuildAllAWBIndex(){ALL_REM_DATES={};Object.values(STATE.cache).forEach(d=>{[...d.cur,...d.prev].forEach(a=>{const payer=freightPayerInfo(a);const k=normRem(payer.nome);if(!k)return;if(!ALL_REM_DATES[k])ALL_REM_DATES[k]=new Set();ALL_REM_DATES[k].add(a.data);});});Object.keys(ALL_REM_DATES).forEach(k=>{ALL_REM_DATES[k]=[...ALL_REM_DATES[k]].sort();});}
function clientCategory(remNorm,dateStr){const dates=ALL_REM_DATES[remNorm];if(!dates||dates.length===0)return'new';const before=dates.filter(d=>d<dateStr);if(before.length===0)return'new';const lastBefore=before[before.length-1];const diffDays=(new Date(dateStr)-new Date(lastBefore))/86400000;return diffDays>90?'ret':null;}

// ══════════════════════════════════════════════════════════
//  MODAL CLIENTES NOVOS / RETORNO
// ══════════════════════════════════════════════════════════
let _mnAll=[],_mnFilter='all';
function openNovos(dateStr,filial){
  if(!Object.keys(ALL_REM_DATES).length)rebuildAllAWBIndex();
  let dayAWBs=[];Object.values(STATE.cache).forEach(d=>{dayAWBs.push(...d.cur.filter(a=>a.data===dateStr&&a.filial===filial));});const seenAWB=new Set();dayAWBs=dayAWBs.filter(a=>{if(seenAWB.has(a.awb))return false;seenAWB.add(a.awb);return true;});
  const clientMap={};dayAWBs.forEach(a=>{const payer=freightPayerInfo(a);const k=normRem(payer.nome);if(!k)return;if(!clientMap[k])clientMap[k]={pagador:payer.nome,cnpj:payer.cnpj||'',papel:payer.papel,tipoFrete:payer.tipo||'',filial:a.filial,awbs:[],totalFrete:0,sources:new Set()};clientMap[k].awbs.push(a.awb);clientMap[k].totalFrete+=a.valor_frete;clientMap[k].sources.add(a.source);});
  _mnAll=Object.entries(clientMap).map(([k,g])=>{const cat=clientCategory(k,dateStr);const dates=ALL_REM_DATES[k]||[];const before=dates.filter(d=>d<dateStr);const lastDate=before.length?before[before.length-1]:null;return{...g,cat,lastDate,normKey:k};}).filter(a=>a.cat!==null);_mnAll.sort((a,b)=>(a.cat==='new'?0:1)-(b.cat==='new'?0:1));
  $('mn-title').textContent=`Pagadores Novos / Retorno — ${filial} · ${fd(dateStr)}`;const totalClientes=Object.keys(clientMap).length;$('mn-sub').textContent=`${_mnAll.length} pagador(es) novo(s) ou retornando (de ${totalClientes} pagadores no dia, ${dayAWBs.length} NFs)`;
  _mnFilter='all';renderMN();$('modal-novos').classList.add('open');document.body.style.overflow='hidden';
}
function closeMN(){$('modal-novos').classList.remove('open');document.body.style.overflow='';}
function mnFilter(f){_mnFilter=f;['all','new','ret'].forEach(k=>{const btn=$(`mn-btn-${k}`);btn.style.background=k===f?'var(--accent)':'';btn.style.color=k===f?'#fff':'';});renderMN();}
function renderMN(){
  const q=($('mn-search')?.value||'').trim().toLowerCase();const base=_mnFilter==='all'?_mnAll:_mnAll.filter(a=>a.cat===_mnFilter);const list=q?base.filter(a=>(a.pagador||'').toLowerCase().includes(q)||(a.cnpj||'').includes(q)):base;
  const container=$('body-mn');container.innerHTML='';$('mn-count').textContent=`${list.length} pagador(es) exibido(s)`;$('mn-empty').style.display=list.length?'none':'block';
  list.forEach((a,i)=>{const isNew=a.cat==='new';const cnpjDisplay=a.cnpj?a.cnpj.replace(/^0+/,'')||a.cnpj:'—';const sourcesStr=[...a.sources].join('+');const lastFmt=a.lastDate?new Date(a.lastDate+'T12:00:00').toLocaleDateString('pt-BR'):null;const row=document.createElement('div');row.style.cssText=`display:grid;grid-template-columns:130px 1fr 60px 120px 56px 72px 148px 108px;gap:0;padding:9px 20px;border-bottom:1px solid var(--border);align-items:center;${i%2===1?'background:rgba(255,255,255,.012)':''}`;row.innerHTML=`<div class="mono" style="font-size:11px;color:var(--cyan);user-select:all;padding-right:8px">${cnpjDisplay}</div><div><div style="font-weight:700;font-size:13px;word-break:break-word;line-height:1.3;padding-right:12px">${a.pagador||'—'}</div><div style="font-size:10px;color:var(--text3);margin-top:3px">${a.tipoFrete||'N/D'} · ${a.papel||'Pagador'}</div></div><div class="mono" style="text-align:center;font-weight:700;font-size:14px">${a.awbs.length}</div><div class="mono" style="color:var(--accent);font-weight:800;font-size:12px">${R$(a.totalFrete)}</div><div><span style="background:rgba(249,115,22,.12);color:var(--accent);padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;font-family:var(--mono)">${a.filial}</span></div><div style="font-size:11px;color:var(--text2)">${sourcesStr}</div><div>${isNew?'<span style="background:rgba(16,185,129,.18);color:var(--success);padding:3px 9px;border-radius:6px;font-size:11px;font-weight:700">Novo</span>':'<span style="background:rgba(249,115,22,.18);color:var(--accent);padding:3px 9px;border-radius:6px;font-size:11px;font-weight:700">Retornou</span>'}</div><div style="font-size:12px;color:${lastFmt?'var(--text2)':'var(--text3)'}">${lastFmt||'— (1° vez)'}</div>`;container.appendChild(row);});
}
function exportMNCSV(){const list=_mnFilter==='all'?_mnAll:_mnAll.filter(a=>a.cat===_mnFilter);let csv='CNPJ,Pagador,Tipo Frete,Papel,AWBs,Frete Total,Filial,Origem,Status,Ultima vez\n';list.forEach(a=>{const status=a.cat==='new'?'Nunca transportou':'Retornou (+90 dias)';csv+=`"${a.cnpj||''}","${a.pagador}","${a.tipoFrete||''}","${a.papel||''}",${a.awbs.length},${a.totalFrete},"${a.filial}","${[...a.sources].join('+')}","${status}","${a.lastDate||''}"\n`;});const el=document.createElement('a');el.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));el.download='clientes_novos_retorno.csv';el.click();}

// ══════════════════════════════════════════════════════════
//  NEGOCIAÇÃO
// ══════════════════════════════════════════════════════════
let NEG_LIST=[],NEG_MATCHES=[],NEG_FILTERED=[];function cleanCNPJ(v){return String(v||'').replace(/\D/g,'');}

// Persistência de dados de negociação
function saveNegData(){
  try{
    localStorage.setItem('freight_neg_list',JSON.stringify(NEG_LIST));
  }catch(e){
    console.error('Erro ao salvar dados de negociação:',e);
  }
}

function loadNegData(){
  try{
    const saved=localStorage.getItem('freight_neg_list');
    if(saved){
      NEG_LIST=JSON.parse(saved);
      if(NEG_LIST.length>0){
        $('neg-file-info').textContent=`${NEG_LIST.length} CNPJs carregados (salvo)`;
        computeNegMatches();
        renderNeg();
        updateNegBadge();
      }
    }
  }catch(e){
    console.error('Erro ao carregar dados de negociação:',e);
  }
}

function openAddNegModal(){
  $('modal-add-neg').style.display='flex';
  document.body.style.overflow='hidden';
  $('add-neg-cnpj').value='';
  $('add-neg-nome').value='';
  $('add-neg-dt').value='';
  $('add-neg-tabela').value='';
  $('add-neg-vendedor').value='';
  $('add-neg-contato').value='';
  $('add-neg-obs').value='';
}

function closeAddNegModal(){
  $('modal-add-neg').style.display='none';
  document.body.style.overflow='';
}

function addNegEntry(){
  const cnpjInput=$('add-neg-cnpj').value.trim();
  if(!cnpjInput){
    toast('CNPJ é obrigatório','err');
    return;
  }
  const cnpj=cleanCNPJ(cnpjInput);
  if(cnpj.length<8){
    toast('CNPJ inválido','err');
    return;
  }
  
  const entry={
    cnpj:cnpj.replace(/^0+/,'')||cnpj,
    nome:$('add-neg-nome').value.trim(),
    dtNeg:$('add-neg-dt').value?new Date($('add-neg-dt').value+'T12:00:00').toLocaleDateString('pt-BR'):'',
    tabela:$('add-neg-tabela').value.trim(),
    vendedor:$('add-neg-vendedor').value.trim(),
    contato:$('add-neg-contato').value.trim(),
    obs:$('add-neg-obs').value.trim()
  };
  
  NEG_LIST.push(entry);
  saveNegData();
  computeNegMatches();
  renderNeg();
  updateNegBadge();
  closeAddNegModal();
  toast('CNPJ adicionado com sucesso!','ok');
  $('neg-file-info').textContent=`${NEG_LIST.length} CNPJs carregados`;
}

function parseNegDate(v){if(!v&&v!==0)return'';if(typeof v==='number'&&v>40000&&v<60000){const d=new Date(Math.round((v-25569)*86400*1000));if(!isNaN(d))return d.toLocaleDateString('pt-BR');}if(v instanceof Date)return v.toLocaleDateString('pt-BR');const s=String(v).trim();if(!s)return'';if(/^\d{4}-\d{2}-\d{2}/.test(s))return new Date(s+'T12:00:00').toLocaleDateString('pt-BR');if(/^\d{2}\/\d{2}\/\d{4}/.test(s))return s;return s;}
function downloadNegTemplate(){const bom='\uFEFF';const csv=bom+'CNPJ,Razão Social,Data Negociação,Tabela,Vendedor,Contato,Observação\n12345678000195,Empresa Exemplo Ltda,2026-01-15,Tabela A,João Silva,(11)99999-9999,Cliente prioritário\n';const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download='planilha_negociacao_modelo.csv';a.click();toast('Planilha modelo baixada!','ok');}
async function loadNegFile(){if(!window.showOpenFilePicker){toast('Use Chrome ou Edge.','err');return;}try{const[fh]=await window.showOpenFilePicker({types:[{description:'Planilha',accept:{'application/vnd.ms-excel':['.xls','.xlsx']}}]});const buf=await(await fh.getFile()).arrayBuffer();const wb=XLSX.read(buf,{type:'array',cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});if(rows.length<2){toast('Planilha vazia.','err');return;}const hdr=rows[0].map(h=>String(h).trim().toLowerCase());const ci=(...ns)=>{for(const n of ns){const i=hdr.findIndex(h=>h.includes(n.toLowerCase()));if(i>=0)return i;}return -1;};const iCNPJ=ci('cnpj'),iNome=ci('razão','razao','nome','social'),iDtNeg=ci('data neg','dt. neg','negociação','negociacao'),iTabela=ci('tabela','tab'),iVend=ci('vendedor','vend'),iCont=ci('contato','telefone','email'),iObs=ci('obs','observ','nota');if(iCNPJ<0){toast('Coluna CNPJ não encontrada.','err');return;}NEG_LIST=[];for(let i=1;i<rows.length;i++){const r=rows[i];let cnpj=cleanCNPJ(r[iCNPJ]);if(!cnpj||cnpj.length<8)continue;cnpj=cnpj.replace(/^0+/,'')||cnpj;NEG_LIST.push({cnpj,nome:iNome>=0?String(r[iNome]||'').trim():'',dtNeg:iDtNeg>=0?parseNegDate(r[iDtNeg]):'',tabela:iTabela>=0?String(r[iTabela]||'').trim():'',vendedor:iVend>=0?String(r[iVend]||'').trim():'',contato:iCont>=0?String(r[iCont]||'').trim():'',obs:iObs>=0?String(r[iObs]||'').trim():''});}$('neg-file-info').textContent=`${NEG_LIST.length} CNPJs carregados de "${fh.name}"`;toast(`✅ ${NEG_LIST.length} CNPJs de negociação carregados.`,'ok');_negPg=1;computeNegMatches();renderNeg();updateNegBadge();saveNegData();}catch(e){if(e.name!=='AbortError')toast('Erro ao ler planilha: '+e.message,'err');}}
function computeNegMatches(){const allAWBs=[];Object.values(STATE.cache).forEach(d=>allAWBs.push(...d.cur,...d.prev));const cnpjMap={};allAWBs.forEach(a=>{if(!a.cnpj)return;const c=cleanCNPJ(a.cnpj);if(!c)return;if(!cnpjMap[c])cnpjMap[c]=[];cnpjMap[c].push(a);});NEG_MATCHES=NEG_LIST.map(entry=>{const awbs=cnpjMap[entry.cnpj]||[];const totalFrete=awbs.reduce((s,a)=>s+a.valor_frete,0);const sorted=[...awbs].sort((a,b)=>a.data>b.data?-1:1);const lastAWB=sorted[0]||null;return{...entry,awbs,totalFrete,lastDate:lastAWB?.data||null,movimentou:awbs.length>0};});}
function updateNegBadge(){const count=NEG_MATCHES.filter(m=>m.movimentou).length;const badge=$('neg-badge-tab');if(count>0){badge.style.display='inline';badge.textContent=count;}else badge.style.display='none';const ib=$('neg-alert-count-badge');if(ib){if(count>0){ib.style.display='inline';ib.textContent=count+' movimentaram';}else ib.style.display='none';}}
function renderNeg(){if(!NEG_LIST.length){$('neg-empty').style.display='flex';$('neg-content').style.display='none';const hdr=$('neg-hdr');if(hdr)hdr.style.display='none';return;}$('neg-empty').style.display='none';$('neg-content').style.display='block';const hdr=$('neg-hdr');if(hdr)hdr.style.display='flex';filterNeg();}
let _negSortCol='movimentou',_negSortAsc=false,_negPg=1;const NEG_PP=60;
function sortNeg(col){if(_negSortCol===col)_negSortAsc=!_negSortAsc;else{_negSortCol=col;_negSortAsc=true;}document.querySelectorAll('#body-neg-head th').forEach(th=>{const c=th.dataset.col;th.textContent=th.dataset.label+(c===col?(_negSortAsc?' ↑':' ↓'):' ↕');});_negPg=1;filterNeg();}
function filterNeg(){const q=($('neg-search')?.value||'').trim().toLowerCase();const st=$('neg-filter-status')?.value||'';const vend=$('neg-filter-vendedor')?.value||'';NEG_FILTERED=NEG_MATCHES.filter(m=>{const cnpjD=(m.cnpj||'').replace(/^0+/,'');if(q&&!cnpjD.includes(q)&&!(m.cnpj||'').includes(q)&&!(m.nome||'').toLowerCase().includes(q)&&!(m.tabela||'').toLowerCase().includes(q))return false;if(st==='movimentou'&&!m.movimentou)return false;if(st==='sem_mov'&&m.movimentou)return false;if(vend&&m.vendedor!==vend)return false;return true;});
NEG_FILTERED.sort((a,b)=>{let av=a[_negSortCol],bv=b[_negSortCol];if(_negSortCol==='movimentou'){av=a.movimentou?1:0;bv=b.movimentou?1:0;}if(_negSortCol==='totalFrete'||_negSortCol==='awbs.length'){av=_negSortCol==='awbs.length'?a.awbs.length:a.totalFrete;bv=_negSortCol==='awbs.length'?b.awbs.length:b.totalFrete;}const numA=parseFloat(av),numB=parseFloat(bv);const cmp=!isNaN(numA)&&!isNaN(numB)?numA-numB:String(av||'').localeCompare(String(bv||''),'pt-BR');return _negSortAsc?cmp:-cmp;});
const negCount=$('neg-list-count');if(negCount)negCount.textContent=`${NEG_FILTERED.length} de ${NEG_MATCHES.length}`;
const vendorSel=$('neg-filter-vendedor');if(vendorSel){const cur=vendorSel.value;const vendors=[...new Set(NEG_MATCHES.map(m=>m.vendedor).filter(Boolean))].sort();vendorSel.innerHTML='<option value="">Todos Vendedores</option>'+vendors.map(v=>`<option value="${v}"${v===cur?' selected':''}>${v}</option>`).join('');}
renderNegPg();}
function renderNegPg(){const tbody=$('body-neg');if(!tbody)return;tbody.innerHTML='';const start=(_negPg-1)*NEG_PP;NEG_FILTERED.slice(start,start+NEG_PP).forEach(m=>{const filiais=[...new Set(m.awbs.map(a=>a.filial))].join(', ')||'—';const cnpjDigits=(m.cnpj||'').replace(/^0+/,'')||m.cnpj;const hasAwbs=m.awbs.length>0;const tr=document.createElement('tr');tr.style.cursor=hasAwbs?'pointer':'default';if(hasAwbs){tr.classList.add('row-clickable');tr.onclick=()=>openNegAwbs(m);}const lastFmt=m.lastDate?new Date(m.lastDate+'T12:00:00').toLocaleDateString('pt-BR'):'—';const nomeDisplay=m.nome?`<span title="${m.nome.replace(/"/g,'&quot;')}" style="display:block;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${m.nome}</span>`:'<span style="color:var(--text3)">—</span>';const tabelaDisplay=m.tabela?`<span title="${m.tabela.replace(/"/g,'&quot;')}" style="display:block;max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${m.tabela}</span>`:'—';const contatoRaw=[m.contato,m.obs].filter(Boolean).join(' · ')||'—';const contatoDisplay=`<span title="${contatoRaw.replace(/"/g,'&quot;')}" style="display:block;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${contatoRaw}</span>`;tr.innerHTML=`<td class="mono" style="font-size:11px;color:var(--cyan);white-space:nowrap">${cnpjDigits}</td><td style="font-weight:${m.nome?'600':'400'};padding-right:8px">${nomeDisplay}</td><td style="white-space:nowrap">${m.movimentou?'<span style="background:rgba(239,68,68,.13);color:var(--danger);padding:2px 7px;border-radius:6px;font-size:11px;font-weight:700">🔔 Mov.</span>':'<span style="color:var(--text3);font-size:11px">Aguard.</span>'}</td><td style="font-size:11px;color:var(--text2);white-space:nowrap">${filiais}</td><td style="font-size:11px;color:var(--text2);white-space:nowrap">${m.dtNeg||'—'}</td><td style="font-size:11px;color:var(--text2)">${tabelaDisplay}</td><td style="font-size:11px;color:var(--text2);white-space:nowrap">${m.vendedor||'—'}</td><td style="white-space:nowrap;color:${m.lastDate?'var(--accent)':'var(--text3)'}">${lastFmt}</td><td class="mono" style="text-align:center">${m.awbs.length||'—'}</td><td class="mono" style="color:${m.totalFrete>0?'var(--success)':'var(--text3)'};white-space:nowrap">${m.totalFrete>0?R$(m.totalFrete):'—'}</td><td style="color:var(--text2);font-size:11px">${contatoDisplay}</td>`;tbody.appendChild(tr);});renderPag('pg-neg',NEG_FILTERED.length,_negPg,NEG_PP,p=>{_negPg=p;renderNegPg();});}
function openNegAwbs(m){const existing=$('modal-neg-awbs');if(existing)existing.remove();const cnpjDigits=(m.cnpj||'').replace(/^0+/,'')||m.cnpj;const sorted=[...m.awbs].sort((a,b)=>b.data.localeCompare(a.data));const rows=sorted.map(a=>`<tr><td class="mono" style="font-size:11px">${a.awb}</td><td>${new Date(a.data+'T12:00:00').toLocaleDateString('pt-BR')}</td><td><span style="background:rgba(249,115,22,.12);color:var(--accent);padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;font-family:var(--mono)">${a.filial}</span></td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${a.remetente}</td><td><span class="tag ${a.source}">${a.source}</span></td><td class="mono" style="color:var(--accent);font-weight:700">${R$(a.valor_frete)}</td><td style="color:var(--text3);font-size:11px">${a.cidade}/${a.uf}</td></tr>`).join('');const modal=document.createElement('div');modal.id='modal-neg-awbs';modal.style.cssText='position:fixed;inset:0;background:rgba(8,14,26,.88);z-index:700;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(4px)';modal.innerHTML=`<div style="background:var(--card);border:1px solid var(--border2);border-radius:16px;width:min(98vw,860px);max-height:86vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.7)"><div style="padding:16px 20px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-start"><div><div style="font-weight:800;font-size:15px">${m.nome||cnpjDigits}</div><div style="font-size:11px;color:var(--text3);margin-top:3px">CNPJ ${cnpjDigits} · ${sorted.length} embarque(s) · Total ${R$(m.totalFrete)}</div></div><button onclick="document.getElementById('modal-neg-awbs').remove();document.body.style.overflow=''" style="background:none;border:none;color:var(--text2);font-size:20px;cursor:pointer;line-height:1;padding:4px 8px">×</button></div><div style="overflow-y:auto;flex:1"><table style="width:100%"><thead><tr><th>AWB</th><th>Data</th><th>Filial</th><th>Remetente</th><th>Tipo</th><th>Frete</th><th>Destino</th></tr></thead><tbody>${rows}</tbody></table></div></div>`;modal.addEventListener('click',e=>{if(e.target===modal){modal.remove();document.body.style.overflow='';}});document.body.appendChild(modal);document.body.style.overflow='hidden';}

// ══════════════════════════════════════════════════════════
//  AWB SORT & EXPORT
// ══════════════════════════════════════════════════════════
let _awbSortCol=-1,_awbSortAsc=true;
function sortAWB(col){if(_awbSortCol===col)_awbSortAsc=!_awbSortAsc;else{_awbSortCol=col;_awbSortAsc=true;}const fields=['awb','data',null,'remetente',null,null,null,'peso',null,'valor_frete'];const field=fields[col];if(!field)return;awbAll.sort((a,b)=>{const av=a[field],bv=b[field];const an=typeof av==='number'?av:parseFloat(av)||0;const bn=typeof bv==='number'?bv:parseFloat(bv)||0;const cmp=!isNaN(an)&&!isNaN(bn)?an-bn:String(av||'').localeCompare(String(bv||''),'pt-BR');return _awbSortAsc?cmp:-cmp;});renderAWBPg();}
function exportAWBCSV(){let csv='AWB,Data,Filial,Remetente,Destinatário,Cidade,UF,Volumes,Peso,Merc.,Frete,Tipo,Modal,Origem,Período\n';awbAll.forEach(a=>{csv+=`"${a.awb}","${a.data}","${a.filial}","${a.remetente}","${a.destinatario}","${a.cidade}","${a.uf}",${a.volumes},${a.peso},${a.valor_mercantil},${a.valor_frete},"${a.tipo_frete}","${a.modal}","${a.source}","${a.periodo}"\n`;});const el=document.createElement('a');el.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));el.download='awb_export.csv';el.click();}
function exportCSV(filial){const{cur,prev}=_data;const dmC=dmap(cur,filial,null),dmP=dmap(prev,filial,null);let csv=`Data,Dia,Carteira,Extra,Total,AWBs,Volumes,${_prevY}_equiv,Delta%\n`;_wdC.forEach((dt,i)=>{const r=dmC[dt];const dp=_wdP[i];const v25=dp?(dmP[dp]?.total||0):0;const tot=r?.total||0;const dp2=v25>0?((tot-v25)/v25*100).toFixed(1):'';csv+=`${dt},${dn(dt)},${r?.carteira||0},${r?.extra||0},${tot},${r?.awbs||0},${r?.volumes||0},${v25},${dp2}\n`;});const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=`${filial}_${_curY}-${padM(_mo)}.csv`;a.click();}

function updatePeriodChip(curYear,month,prevYear){$('period-chip').innerHTML=`${MONTHS[month]} ${curYear}<small>vs ${prevYear} · ${STATE.curIdx+1}/${STATE.periods.length}</small>`;}

// ══════════════════════════════════════════════════════════
//  PAGINATION
// ══════════════════════════════════════════════════════════
function renderPag(cid,total,cur,pp,onClick){const el=$(cid);el.innerHTML='';const tp=Math.ceil(total/pp);if(tp<=1)return;const mk=(lbl,p,active,dis)=>{const b=document.createElement('button');b.className='pg-btn'+(active?' active':'');b.textContent=lbl;b.disabled=dis;if(!dis)b.onclick=()=>onClick(p);el.appendChild(b);};mk('‹',cur-1,false,cur===1);let pgs=[];if(tp<=7)for(let i=1;i<=tp;i++)pgs.push(i);else{pgs=[1];if(cur>3)pgs.push('…');for(let i=Math.max(2,cur-1);i<=Math.min(tp-1,cur+1);i++)pgs.push(i);if(cur<tp-2)pgs.push('…');pgs.push(tp);}pgs.forEach(p=>mk(p,p,p===cur,p==='…'));mk('›',cur+1,false,cur===tp);const info=document.createElement('span');info.className='pg-info';info.textContent=`${cur}/${tp}`;el.appendChild(info);}

// ══════════════════════════════════════════════════════════
//  CHARTS HELPERS
// ══════════════════════════════════════════════════════════
const CHARTS={};function dch(k){if(CHARTS[k]){CHARTS[k].destroy();delete CHARTS[k];}}function mkChart(id,type,data,opts){return new Chart($(id),{type,data,options:opts});}
function co({stacked=false}={}){return{responsive:true,plugins:{legend:{labels:{color:'#7c93b0',font:{size:11}}},tooltip:{callbacks:{label:c=>' '+R$(c.raw||0)}}},scales:{x:{stacked,grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10}}},y:{stacked,grid:{color:'rgba(30,51,82,.5)'},ticks:{color:'#3d5270',font:{size:10},callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}};}

// ══════════════════════════════════════════════════════════
//  TABLE SORT
// ══════════════════════════════════════════════════════════
const sortSt={};function srt(id,col){const t=$(id);if(!t)return;const rows=Array.from(t.querySelector('tbody').rows);const key=id+'_'+col;const asc=sortSt[key]!=='asc';sortSt[key]=asc?'asc':'desc';rows.sort((a,b)=>{const av=a.cells[col]?.textContent.replace(/[R$\s,.%+]/g,'')||'';const bv=b.cells[col]?.textContent.replace(/[R$\s,.%+]/g,'')||'';const an=parseFloat(av),bn=parseFloat(bv);if(!isNaN(an)&&!isNaN(bn))return asc?an-bn:bn-an;return asc?av.localeCompare(bv,'pt-BR'):bv.localeCompare(av,'pt-BR');});rows.forEach(r=>t.querySelector('tbody').appendChild(r));}

// ══════════════════════════════════════════════════════════
//  TABS
// ══════════════════════════════════════════════════════════
function showPane(name,el){$$('.pane').forEach(p=>p.classList.remove('active'));$$('.tab').forEach(t=>t.classList.remove('active'));const paneEl=$('pane-'+name);if(paneEl)paneEl.classList.add('active');if(el)el.classList.add('active');renderPane(name);}

// ══════════════════════════════════════════════════════════
//  UTILS
// ══════════════════════════════════════════════════════════
function set(id,val,style){const el=$(id);if(!el)return;el.textContent=val;if(style)el.style.cssText=style;}
function bb(id,txt,cls){const el=$(id);if(!el)return;el.textContent=txt;el.className='badge '+cls;}
function showLoading(txt,sub='',pct=null){$('loading').classList.remove('hidden');$('loading-txt').textContent=txt;$('loading-sub').textContent=sub||'';const wrap=$('loading-bar-wrap'),fill=$('loading-bar-fill');if(pct!==null){wrap.style.display='block';fill.style.width=pct+'%';}else{wrap.style.display='none';fill.style.width='0%';}}
function hideLoading(){$('loading').classList.add('hidden');const fill=$('loading-bar-fill');if(fill)fill.style.width='0%';const wrap=$('loading-bar-wrap');if(wrap)wrap.style.display='none';}
// ══════════════════════════════════════════════════════════
//  COMISSÃO
// ══════════════════════════════════════════════════════════
function getCommRates(){
  const pctFOB=1.3;
  const pctCIF=0.2;
  const pctMy=100;
  return{pctFOB,pctCIF,pctMy,myFOB:pctFOB*pctMy/100,myCIF:pctCIF*pctMy/100};
}

function calcComm(awbs,filial){
  const r=getCommRates();
  const src=filial?awbs.filter(a=>a.filial===filial):awbs;
  const fobBase=src.filter(a=>(a.tipo_frete||'').toUpperCase()==='FOB').reduce((s,a)=>s+a.valor_frete,0);
  const cifBase=src.filter(a=>(a.tipo_frete||'').toUpperCase()==='CIF').reduce((s,a)=>s+a.valor_frete,0);
  const outros=src.filter(a=>!['FOB','CIF'].includes((a.tipo_frete||'').toUpperCase())).reduce((s,a)=>s+a.valor_frete,0);
  const myComm=fobBase*r.pctFOB/100+cifBase*r.pctCIF/100;
  return{fobBase,cifBase,outros,myComm};
}

function calcCommSC(sc,filial){
  if(!sc||!sc.length)return{fobBase:0,cifBase:0,myComm:0};
  const r=getCommRates();
  const src=(filial?sc.filter(a=>a.filial===filial):sc);
  const fobBase=src.filter(a=>(a.tipo_frete||'').toUpperCase()==='FOB').reduce((s,a)=>s+a.valor_frete,0);
  const cifBase=src.filter(a=>(a.tipo_frete||'').toUpperCase()==='CIF').reduce((s,a)=>s+a.valor_frete,0);
  const myComm=fobBase*r.pctFOB/100+cifBase*r.pctCIF/100;
  return{fobBase,cifBase,myComm};
}

function renderComissao(){
  if(!_data){return;}
  const r=getCommRates();
  if($('comm-period'))$('comm-period').textContent=`${MONTHS[_mo]} ${_curY}`;
  if($('comm-rates-effective'))$('comm-rates-effective').innerHTML=`Taxa efetiva sua: <strong style="color:var(--gold);font-family:var(--mono)">${r.myFOB.toFixed(4)}% FOB</strong> · <strong style="color:var(--cyan);font-family:var(--mono)">${r.myCIF.toFixed(4)}% CIF</strong>`;

  // Acumula todos os períodos do ano
  const allCur=[];STATE.periods.forEach(p=>{const d=STATE.cache[p];if(d)allCur.push(...d.cur);});
  const{sc}=_data; // SC só do mês atual

  // Banner principal — comissão total do ano
  const commAno=calcComm(allCur,null);
  const commMes=calcComm(_data.cur,null);
  const commSCMes=calcCommSC(sc,null);
  const today=new Date().toISOString().split('T')[0];
  const elapsed=_wdC.filter(d=>d<=today).length;
  const projMes=elapsed>0?(commMes.myComm/elapsed)*_wdC.length:0;

  const bannerGrid=$('comm-banner-grid');
  if(bannerGrid){
    bannerGrid.innerHTML=`
      <div class="kpi" style="border-top:3px solid var(--gold)">
        <div class="kpi-lbl">💰 Minha Comissão — Ano</div>
        <div class="kpi-val" style="color:var(--gold)">${R$(commAno.myComm)}</div>
      </div>
      <div class="kpi" style="border-top:3px solid var(--cyan)">
        <div class="kpi-lbl">💰 Minha Comissão — Mês</div>
        <div class="kpi-val" style="color:var(--cyan)">${R$(commMes.myComm)}</div>
        <div class="kpi-sub">Projeção mês: ${R$(projMes)}</div>
      </div>
      <div class="kpi" style="border-top:3px solid var(--accent)">
        <div class="kpi-lbl">FOB faturado — Mês</div>
        <div class="kpi-val" style="color:var(--accent)">${R$(commMes.fobBase)}</div>
        <div class="kpi-sub">Ano: ${R$(commAno.fobBase)}</div>
      </div>
      <div class="kpi" style="border-top:3px solid var(--cyan)">
        <div class="kpi-lbl">CIF faturado — Mês</div>
        <div class="kpi-val" style="color:var(--cyan)">${R$(commMes.cifBase)}</div>
        <div class="kpi-sub">Ano: ${R$(commAno.cifBase)}</div>
      </div>
      ${commSCMes.myComm>0?`
      <div class="kpi" style="border-top:3px solid var(--sc);background:var(--sc-bg)">
        <div class="kpi-lbl">⚠ SC Potencial — Minha Comissão</div>
        <div class="kpi-val" style="color:var(--sc)">${R$(commSCMes.myComm)}</div>
        <div class="kpi-sub">Se todo SC migrar: ${R$(commMes.myComm+commSCMes.myComm)}</div>
      </div>`:''}`;
  }

  // SC section
  const scSection=$('comm-sc-section');
  const scGrid=$('comm-sc-potential-grid');
  if(sc&&sc.length>0&&scSection&&scGrid){
    scSection.style.display='block';
    scGrid.innerHTML='';
    STATE.filiais.forEach((filial,fi)=>{
      const col=FILIAL_COLORS[fi%FILIAL_COLORS.length];
      const scF=calcCommSC(sc,filial);
      if(scF.myComm<=0)return;
      scGrid.innerHTML+=`
        <div style="background:var(--card);border:1px solid var(--sc-border);border-radius:10px;padding:12px 14px">
          <div style="font-size:11px;font-weight:700;color:${col.color};margin-bottom:8px">${filial}</div>
          <div style="font-size:14px;font-weight:800;font-family:var(--mono);color:var(--sc);margin-bottom:4px">${R$(scF.myComm)}</div>
          <div style="font-size:10px;color:var(--text3)">FOB SC: ${R$(scF.fobBase)} · CIF SC: ${R$(scF.cifBase)}</div>
        </div>`;
    });
  }else if(scSection)scSection.style.display='none';

  // Tabela por filial
  const tbody=$('body-comm-filial');
  if(tbody){
    tbody.innerHTML='';
    let totFOB=0,totCIF=0,totMy=0,totSCMy=0,totSCBase=0;
    STATE.filiais.forEach((filial,fi)=>{
      const col=FILIAL_COLORS[fi%FILIAL_COLORS.length];
      const c=calcComm(allCur,filial);
      const cSC=calcCommSC(sc,filial);
      totFOB+=c.fobBase;totCIF+=c.cifBase;totMy+=c.myComm;totSCMy+=cSC.myComm;totSCBase+=cSC.fobBase+cSC.cifBase;
      const tr=document.createElement('tr');
      tr.innerHTML=`
        <td><span style="background:rgba(249,115,22,.12);color:${col.color};padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700">${filial}</span></td>
        <td style="text-align:right;font-family:var(--mono)">${R$(c.fobBase)}</td>
        <td style="text-align:right;font-family:var(--mono)">${R$(c.cifBase)}</td>
        <td style="text-align:right;font-family:var(--mono);font-weight:800;color:var(--gold)">${R$(c.myComm)}</td>
        <td style="text-align:right;font-family:var(--mono);color:var(--sc)">${cSC.fobBase+cSC.cifBase>0?R$(cSC.fobBase+cSC.cifBase):'—'}</td>
        <td style="text-align:right;font-family:var(--mono);color:var(--warning)">${cSC.myComm>0?R$(cSC.myComm):'—'}</td>`;
      tbody.appendChild(tr);
    });
    // Total row
    const trTot=document.createElement('tr');
    trTot.style.cssText='font-weight:800;background:rgba(245,158,11,.05);border-top:2px solid var(--gold-border)';
    trTot.innerHTML=`
      <td style="color:var(--gold)">TOTAL ANO</td>
      <td style="text-align:right;font-family:var(--mono)">${R$(totFOB)}</td>
      <td style="text-align:right;font-family:var(--mono)">${R$(totCIF)}</td>
      <td style="text-align:right;font-family:var(--mono);color:var(--gold);font-size:15px">${R$(totMy)}</td>
      <td style="text-align:right;font-family:var(--mono);color:var(--sc)">${totSCBase>0?R$(totSCBase):'—'}</td>
      <td style="text-align:right;font-family:var(--mono);color:var(--warning)">${totSCMy>0?R$(totSCMy):'—'}</td>`;
    tbody.appendChild(trTot);
  }

  // Gráfico mensal
  dch('chCommMonthly');
  if($('ch-comm-monthly')){
    const labels=[];const myData=[];const scData=[];
    STATE.periods.forEach(p=>{
      const{year,month}=decodePeriod(p);const d=STATE.cache[p];if(!d)return;
      labels.push(MONTHS[month].substring(0,3)+' '+year);
      const c=calcComm(d.cur,null);myData.push(c.myComm);
      const cSC=month===_mo?calcCommSC(d.sc,null).myComm:0;scData.push(cSC);
    });
    CHARTS.chCommMonthly=mkChart('ch-comm-monthly','bar',{labels,datasets:[
      {label:'Comissão Comissionada',data:myData,backgroundColor:'rgba(245,158,11,.8)',borderRadius:4,stack:'a'},
      {label:'Potencial SC',data:scData,backgroundColor:'rgba(167,139,250,.6)',borderRadius:4,stack:'a'}
    ]},co({stacked:true}));
  }

  // Gráfico por tipo — mês atual
  dch('chCommTipo');
  if($('ch-comm-tipo')){
    const c=calcComm(_data.cur,null);
    const cSC=calcCommSC(sc,null);
    CHARTS.chCommTipo=mkChart('ch-comm-tipo','doughnut',{
      labels:['FOB (comissionado)','CIF (comissionado)','FOB SC potencial','CIF SC potencial'],
      datasets:[{
        data:[c.fobBase*r.pctFOB/100,c.cifBase*r.pctCIF/100,
              cSC?cSC.fobBase*r.pctFOB/100*r.pctMy/100:0,
              cSC?cSC.cifBase*r.pctCIF/100*r.pctMy/100:0],
        backgroundColor:['rgba(249,115,22,.8)','rgba(6,182,212,.8)','rgba(167,139,250,.5)','rgba(167,139,250,.3)'],
        borderWidth:1
      }]
    },{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{position:'right',labels:{color:'#7c93b0',font:{size:10},boxWidth:10}},tooltip:{callbacks:{label:c=>' '+R$(c.raw)}}}});
  }

  // Tabela diária do mês atual
  if($('comm-mes-atual-lbl'))$('comm-mes-atual-lbl').textContent=`${MONTHS[_mo]} ${_curY}`;
  const dmCurAll=dmap(_data.cur,null,_data.sc);
  const dmScAll=_data.sc?dmap([],null,_data.sc):{};
  const tBodyD=$('body-comm-daily');
  const tFoot=$('foot-comm-daily');
  if(tBodyD){
    tBodyD.innerHTML='';
    let totFOBD=0,totCIFD=0,totGerD=0,totMyD=0,totSCD=0,totSCMyD=0;
    _wdC.forEach((dt,i)=>{
      const dayAWBs=_data.cur.filter(a=>a.data===dt);
      if(dayAWBs.length===0&&dt>today)return;
      const c=calcComm(dayAWBs,null);
      const scDay=_data.sc?_data.sc.filter(a=>a.data===dt):[];
      const scDayComm=scDay.reduce((s,a)=>{
        const rr=getCommRates();
        const isFOB=(a.tipo_frete||'').toUpperCase()==='FOB';
        return s+(isFOB?a.valor_frete*rr.pctFOB/100*rr.pctMy/100:a.valor_frete*rr.pctCIF/100*rr.pctMy/100);
      },0);
      const scDayBase=scDay.reduce((s,a)=>s+a.valor_frete,0);
      totFOBD+=c.fobBase;totCIFD+=c.cifBase;totMyD+=c.myComm;totSCD+=scDayBase;totSCMyD+=scDayComm;
      const isToday=dt===today;const isFuture=dt>today;
      const tr=document.createElement('tr');
      tr.style.background=isToday?'rgba(245,158,11,.06)':isFuture?'rgba(255,255,255,.01)':'';
      tr.innerHTML=`
        <td class="mono"><strong>${fd(dt)}</strong>${isToday?' <span style="background:var(--gold);color:#000;padding:1px 5px;border-radius:3px;font-size:9px;font-weight:800">HOJE</span>':''}</td>
        <td style="color:var(--text2)">${dn(dt)}</td>
        <td style="text-align:right;font-family:var(--mono);color:${c.fobBase>0?'var(--accent)':'var(--text3)'}">${c.fobBase>0?R$(c.fobBase):'—'}</td>
        <td style="text-align:right;font-family:var(--mono);color:${c.cifBase>0?'var(--cyan)':'var(--text3)'}">${c.cifBase>0?R$(c.cifBase):'—'}</td>
        <td style="text-align:right;font-family:var(--mono);font-weight:800;color:${c.myComm>0?'var(--gold)':'var(--text3)'}">${c.myComm>0?R$(c.myComm):'—'}</td>
        <td style="text-align:right;font-family:var(--mono);color:var(--sc)">${scDayBase>0?R$(scDayBase):'—'}</td>
        <td style="text-align:right;font-family:var(--mono);color:var(--warning)">${scDayComm>0?R$(scDayComm):'—'}</td>`;
      tBodyD.appendChild(tr);
    });
    if(tFoot)tFoot.innerHTML=`<tr style="font-weight:800;background:rgba(245,158,11,.07);border-top:2px solid var(--gold-border)">
      <td colspan="2" style="color:var(--gold)">TOTAL MÊS</td>
      <td style="text-align:right;font-family:var(--mono)">${R$(totFOBD)}</td>
      <td style="text-align:right;font-family:var(--mono)">${R$(totCIFD)}</td>
      <td style="text-align:right;font-family:var(--mono);color:var(--gold);font-size:14px">${R$(totMyD)}</td>
      <td style="text-align:right;font-family:var(--mono);color:var(--sc)">${totSCD>0?R$(totSCD):'—'}</td>
      <td style="text-align:right;font-family:var(--mono);color:var(--warning)">${totSCMyD>0?R$(totSCMyD):'—'}</td>
    </tr>`;
    if($('comm-mes-total-lbl'))$('comm-mes-total-lbl').textContent=`Total mês comissão minha: ${R$(totMyD)}${totSCMyD>0?' + '+R$(totSCMyD)+' SC potencial':''}`;
  }
}

function exportCommCSV(){
  const r=getCommRates();
  const allCur=[];STATE.periods.forEach(p=>{const d=STATE.cache[p];if(d)allCur.push(...d.cur);});
  const{sc}=_data;
  let csv=`Filial,FOB Base,CIF Base,Minha Comissão,SC a Recuperar,Comissão Potencial SC\n`;
  STATE.filiais.forEach(filial=>{
    const c=calcComm(allCur,filial);const cSC=calcCommSC(sc,filial);
    csv+=`"${filial}",${c.fobBase},${c.cifBase},${c.myComm},${cSC.fobBase+cSC.cifBase},${cSC.myComm}\n`;
  });
  csv+=`\nTaxas utilizadas: FOB ${r.pctFOB}% | CIF ${r.pctCIF}%\n`;
  csv+=`Taxa efetiva sua: FOB ${r.myFOB.toFixed(4)}% | CIF ${r.myCIF.toFixed(4)}%\n`;
  const el=document.createElement('a');el.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));el.download=`comissao_${_curY}.csv`;el.click();
}

function toast(msg,type='ok'){const t=$('toastEl');t.textContent=(type==='ok'?'OK: ':'Erro: ')+msg;t.className=`toast ${type} show`;clearTimeout(t._t);t._t=setTimeout(()=>t.classList.remove('show'),4500);}
// ── 1. RASTREAR FILIAIS COM SC ─────────────────────────────────
if (!STATE.scFiliais) STATE.scFiliais = new Set();

// ── 2. parseXLS — versão corrigida ────────────────────────────
parseXLS = async function(handle, filial, source, year) {
  try {
    const buf = await (await handle.getFile()).arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });

    // Header: busca até linha 12 para cobrir planilhas SC com cabeçalho deslocado
    // Reconhece tanto 'AWB' quanto 'CTRC' (Braspress SC) e combinação Remetente+Data
    let hi = 0;
    for (let i = 0; i < Math.min(12, rows.length); i++) {
      const hasAWB  = rows[i].some(c => /\bAWB\b/i.test(String(c)));
      const hasCTRC = rows[i].some(c => /\bCTRC\b/i.test(String(c)));
      const hasRem  = rows[i].some(c => /^Remetente$/i.test(String(c).trim()));
      const hasData = rows[i].some(c => /^Data$/i.test(String(c).trim()));
      if (hasAWB || hasCTRC || (hasRem && hasData)) { hi = i; break; }
    }

    const hdr = rows[hi].map(h => String(h).trim());
    const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
    const normHdr = hdr.map(norm);
    const ci = (...ns) => {
      for (const n of ns) {
        const nk = norm(n);
        const exact = normHdr.findIndex(h => h === nk);
        if (exact >= 0) return exact;
      }
      for (const n of ns) {
        const nk = norm(n);
        const partial = normHdr.findIndex(h => h.includes(nk));
        if (partial >= 0) return partial;
      }
      return -1;
    };

    const I = {
      // CTRC = AWB nas planilhas SC da Braspress
      awb:     ci('Nro. AWB','Numero AWB','Nr. AWB','Num. AWB','AWB','CTRC','Nro','Nr.','Numero'),
      data:    ci('Data Emissão','Data Emissao','Dt. Emissão','Dt. Emissao','Dt Emissao','Data de Emissão','Emitido','Emissão','Data','Dt.'),
      rem:     ci('Remetente','Nome Remetente','Razão Social Rem','Razao Social Rem','Remet'),
      dest:    ci('Destinatário','Destinatario','Nome Destinatário','Razão Social Dest','Razao Social Dest','Dest'),
      cid:     ci('Cidade Destino','Cidade Dest','Municipio','Município','Cidade'),
      uf:      ci('UF Destino','UF Dest','Estado Destino','Estado','UF'),
      // 'Volume' (SC Braspress) antes de 'Volumes' (demais)
      vol:     ci('Volume','Volumes','Qtd Volumes','Qtd. Volumes','Quantidade','Vol','Qtd'),
      peso:    ci('Peso Real','Peso Bruto','Peso Kg','Peso (kg)','Peso'),
      merc:    ci('Valor Mercadoria','Valor Mercantil','Valor da Mercadoria','Vl Mercantil','Vlr Mercantil','Vl. Merc.','Mercantil','Merc'),
      fret:    ci('Valor Frete','Valor do Frete','Valor Frete Total','Vl. do Frete','Vl Frete','Vlr Frete','Vl. Frete','Total Frete','Frete'),
      tipo:    ci('Tipo Frete','Tipo de Frete','Tp. Frete','Tp Frete','Frete Tipo'),
      modal:   ci('Modal de Transporte','Modalidade Transporte','Modalidade','Modal'),
      ramo:    ci('Ramo de Atividade','Ramo Atividade','Segmento Atividade','Ramo','Segmento'),
      cnpjRem: ci('CNPJ Remetente','CNPJ/CPF Remetente','CNPJ Rem','CNPJ do Remetente','CNPJ Remet.'),
      cnpjDest:ci('CNPJ Destinatário','CNPJ Destinatario','CNPJ/CPF Destinatário','CNPJ/CPF Destinatario','CNPJ Dest'),
      cnpj:    ci('CNPJ Remetente','CNPJ/CPF','C.N.P.J.','C.N.P.J','CNPJ','Cnpj'),
    };

    // Fallback: detectar coluna de data por tipo de valor na primeira linha de dados
    if (I.data < 0) {
      for (let col = 0; col < hdr.length; col++) {
        const sample = rows[hi + 1]?.[col];
        if (!sample && sample !== 0) continue;
        if (sample instanceof Date) { I.data = col; break; }
        if (typeof sample === 'number' && sample > 40000 && sample < 55000) { I.data = col; break; }
        if (typeof sample === 'string' && /\d{2}\/\d{2}\/\d{4}/.test(sample)) { I.data = col; break; }
        if (typeof sample === 'string' && /\d{4}-\d{2}-\d{2}/.test(sample)) { I.data = col; break; }
      }
    }

    const recs = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      const awbVal = r[I.awb];
      if (!awbVal && awbVal !== 0) continue;

      // ── ANTI-LIXO: linha de totais/rodapé ──────────────────
      // Planilhas SC Braspress terminam com uma linha de totais onde
      // col 0 = número pequeno (contagem de registros, ex: 9).
      // CTRCs reais são >= 100000. Descartamos números pequenos em source=sc.
      if (source === 'sc' && typeof awbVal === 'number' && awbVal < 100000) continue;

      // Linha sem data válida = rodapé ou linha em branco
      const dataVal = I.data >= 0 ? r[I.data] : '';
      if (!dataVal && dataVal !== 0) continue;

      // ── PARSE DE DATA ───────────────────────────────────────
      let dt = dataVal;
      if (dt instanceof Date) {
        const y = dt.getFullYear(), m = String(dt.getMonth()+1).padStart(2,'0'), d = String(dt.getDate()).padStart(2,'0');
        dt = `${y}-${m}-${d}`;
      } else if (typeof dt === 'number' && dt > 0) {
        try {
          const msUtc = Math.round((dt - 25569) * 86400 * 1000);
          const jsDate = new Date(msUtc);
          if (!isNaN(jsDate.getTime())) {
            const y = jsDate.getUTCFullYear(), m = String(jsDate.getUTCMonth()+1).padStart(2,'0'), d = String(jsDate.getUTCDate()).padStart(2,'0');
            dt = `${y}-${m}-${d}`;
          } else continue;
        } catch (e) { continue; }
      } else {
        dt = String(dt).trim().split(' ')[0].split('T')[0];
        if (!dt) continue;
        if (dt.includes('/')) {
          const p = dt.split('/').map(v => v.trim());
          if (p.length === 3) {
            let [a, b, c] = p;
            a = a.padStart(2,'0'); b = b.padStart(2,'0');
            if (c.length === 2) c = '20' + c;
            dt = `${c}-${b}-${a}`;
          }
        }
        const tmp = new Date(dt + 'T12:00:00');
        if (isNaN(tmp.getTime())) continue;
        dt = tmp.toISOString().slice(0, 10);
      }

      // Ano deve ser razoável
      const dtYear = parseInt(dt.slice(0, 4));
      if (dtYear < 2015 || dtYear > 2040) continue;

      const cnpjRem  = I.cnpjRem  >= 0 ? (String(r[I.cnpjRem]  || '').replace(/\D/g,'').replace(/^0+/,'') || '') :
                       I.cnpj     >= 0 ? (String(r[I.cnpj]     || '').replace(/\D/g,'').replace(/^0+/,'') || '') : '';
      const cnpjDest = I.cnpjDest >= 0 ? (String(r[I.cnpjDest] || '').replace(/\D/g,'').replace(/^0+/,'') || '') : '';

      recs.push({
        awb: String(awbVal).trim(),
        data: dt,
        filial, source, year,
        remetente:      String(r[I.rem]  || '').trim().substring(0, 45),
        destinatario:   String(r[I.dest] || '').substring(0, 45),
        cidade:         String(r[I.cid]  || ''),
        uf:             String(r[I.uf]   || ''),
        volumes:        parseInt(r[I.vol]) || 0,
        peso:           pNum(r[I.peso]),
        valor_mercantil:pNum(r[I.merc]),
        valor_frete:    pNum(r[I.fret]),
        tipo_frete:     String(r[I.tipo]  || ''),
        modal:          normModal(r[I.modal]),
        ramo:           String(r[I.ramo]  || '').substring(0, 40),
        cnpj: cnpjRem, cnpj_remetente: cnpjRem, cnpj_destinatario: cnpjDest,
        periodo: STATE.periods[STATE.curIdx] || ''
      });
    }
    return recs;
  } catch (e) { return []; }
};

// ── 3. loadPeriodData — registrar filiais com SC ──────────────
loadPeriodData = async function(curYear, prevYear, month) {
  const mo = padM(month);
  const curDir  = await getMonthDir(curYear, mo);
  const prevDir = await getMonthDir(prevYear, mo);
  const allCur  = curDir  ? await readDirAWBs(curDir,  curYear) : [];
  const prev    = prevDir ? await readDirAWBs(prevDir, prevYear) : await readAnnualYearAWBs(prevYear, month);

  const cur = allCur.filter(a => a.source !== 'sc');
  const sc  = allCur.filter(a => a.source === 'sc');

  // Registrar globalmente quais filiais têm arquivo SC
  if (!STATE.scFiliais) STATE.scFiliais = new Set();
  sc.forEach(a => STATE.scFiliais.add(a.filial));

  // Marcar SC que já migraram para Carteira/Extra

  return { cur, prev, sc };
};

// ── 4. Helpers de visibilidade SC ─────────────────────────────
function filialHasSC(filial) {
  return !!(STATE.scFiliais && STATE.scFiliais.has(filial));
}
function anyFilialHasSC() {
  return !!(STATE.scFiliais && STATE.scFiliais.size > 0);
}

// ── 5. Patch renderFilial — ocultar/mostrar seções SC ─────────
// Guarda a função original e envolve com lógica condicional SC

console.log('[FRT] Carregado. Parser CTRC, freshness, pagadores por filial.');

// ══════════════════════════════════════════════════════════
//  ONDA 4 — PAGADORES NOVOS/RETORNO: ABA DEDICADA POR FILIAL
// ══════════════════════════════════════════════════════════

// Estado do painel de pagadores
const PAYER_STATE={};

// Injeta sub-aba Pagadores + container no pane de cada filial
function injectPayersTabIntoFilialPane(f){
  const pfx=f.toLowerCase();
  const dataEl=$(pfx+'-data');if(!dataEl)return;
  if($(`payer-panel-${pfx}`))return; // já injetado

  // Sub-abas (Faturamento | Pagadores)
  const subtabsDiv=document.createElement('div');
  subtabsDiv.className='subtabs';
  subtabsDiv.style.marginBottom='0';
  subtabsDiv.innerHTML=`
    <div class="subtab active" id="stab-fat-${pfx}" onclick="switchFilialTab('fat','${pfx}')">Faturamento</div>
    <div class="subtab" id="stab-pay-${pfx}" onclick="switchFilialTab('pay','${pfx}')">Pagadores</div>`;

  // Wrapper para faturamento (conteúdo original)
  const fatWrap=document.createElement('div');
  fatWrap.id=`fat-panel-${pfx}`;
  // Mover conteúdo original para dentro do wrapper
  while(dataEl.firstChild)fatWrap.appendChild(dataEl.firstChild);

  // Container do painel de pagadores
  const payPanel=document.createElement('div');
  payPanel.id=`payer-panel-${pfx}`;
  payPanel.style.display='none';
  payPanel.innerHTML=buildPayerPanelHTML(f,pfx);

  dataEl.appendChild(subtabsDiv);
  dataEl.appendChild(fatWrap);
  dataEl.appendChild(payPanel);
}

function buildPayerPanelHTML(f,pfx){
  return `
  <div style="margin-top:18px">
    <!-- KPI summary row -->
    <div class="payer-day-grid" id="pay-kpi-${pfx}"></div>
    <!-- Day selector -->
    <div class="tbl-card">
      <div class="tbl-top" style="gap:10px;flex-wrap:wrap">
        <div class="tbl-top-ttl">${ICONS.users} Pagadores por Dia — ${f}</div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <select class="inp" id="pay-filter-cat-${pfx}" onchange="renderPayerDayList('${pfx}')" style="font-size:12px">
            <option value="">Todos os dias</option>
            <option value="new">Com novos</option>
            <option value="ret">Com retorno</option>
          </select>
          <input class="inp" id="pay-search-${pfx}" placeholder="filtrar cliente/CNPJ..." oninput="filterPayerDetail('${pfx}')" style="font-size:12px;min-width:160px">
          <button class="btn btn-s" onclick="exportPayerCSV('${pfx}')" style="font-size:11px">${ICONS.download} CSV</button>
        </div>
      </div>
      <div class="payer-day-head">
        <div>Data</div><div>Pagador</div><div>AWBs</div><div>Frete</div><div>Status</div><div>Visto antes</div>
      </div>
      <div id="pay-day-list-${pfx}" style="max-height:480px;overflow-y:auto"></div>
    </div>
  </div>`;
}

function switchFilialTab(tab,pfx){
  const fatPanel=$(`fat-panel-${pfx}`);
  const payPanel=$(`payer-panel-${pfx}`);
  const stabFat=$(`stab-fat-${pfx}`);
  const stabPay=$(`stab-pay-${pfx}`);
  if(!fatPanel||!payPanel)return;
  const isFat=tab==='fat';
  fatPanel.style.display=isFat?'':'none';
  payPanel.style.display=isFat?'none':'';
  stabFat?.classList.toggle('active',isFat);
  stabPay?.classList.toggle('active',!isFat);
  if(!isFat&&!PAYER_STATE[pfx]?.built)buildPayerPanelData(pfx);
}

// Constrói os dados de pagadores para um pfx de filial
function buildPayerPanelData(pfx){
  const filial=pfx.toUpperCase();
  if(!_data||!Object.keys(ALL_REM_DATES).length)rebuildAllAWBIndex();

  // Coleta AWBs de todos os períodos para essa filial
  const allFilialAWBs=[];
  Object.values(STATE.cache).forEach(d=>{
    allFilialAWBs.push(...d.cur.filter(a=>a.filial===filial));
  });

  // Por dia útil do mês atual
  const byDay={};
  _wdC.forEach(dt=>{
    const dayAWBs=allFilialAWBs.filter(a=>a.data===dt);
    if(!dayAWBs.length)return;
    const clients={};
    dayAWBs.forEach(a=>{
      const payer=freightPayerInfo(a);
      const k=normRem(payer.nome);
      if(!k)return;
      if(!clients[k])clients[k]={pagador:payer.nome,cnpj:payer.cnpj||'',tipoFrete:payer.tipo,papel:payer.papel,awbs:[],totalFrete:0};
      clients[k].awbs.push(a.awb);
      clients[k].totalFrete+=a.valor_frete;
    });
    const entries=Object.entries(clients).map(([k,g])=>{
      const cat=clientCategory(k,dt);
      const dates=ALL_REM_DATES[k]||[];
      const before=dates.filter(d=>d<dt);
      const lastDate=before.length?before[before.length-1]:null;
      return{...g,cat,lastDate,normKey:k,date:dt};
    });
    byDay[dt]={entries,newCount:entries.filter(e=>e.cat==='new').length,retCount:entries.filter(e=>e.cat==='ret').length,totalFrete:dayAWBs.reduce((s,a)=>s+a.valor_frete,0),awbCount:dayAWBs.length};
  });

  PAYER_STATE[pfx]={byDay,built:true,filial};

  // KPIs
  const totNew=Object.values(byDay).reduce((s,d)=>s+d.newCount,0);
  const totRet=Object.values(byDay).reduce((s,d)=>s+d.retCount,0);
  const daysWithNew=Object.values(byDay).filter(d=>d.newCount>0).length;
  const kpiEl=$(`pay-kpi-${pfx}`);
  if(kpiEl)kpiEl.innerHTML=`
    <div class="payer-stat new"><div class="payer-stat-lbl">${ICONS.users} Clientes Novos<br><span style="font-size:9px;font-weight:400">nunca transportaram</span></div><div class="payer-stat-val" style="color:var(--success)">${totNew}</div><div class="payer-stat-sub">${daysWithNew} dia(s) com novos</div></div>
    <div class="payer-stat ret"><div class="payer-stat-lbl">${ICONS.trend} Retornaram<br><span style="font-size:9px;font-weight:400">ausentes +90 dias</span></div><div class="payer-stat-val" style="color:var(--accent)">${totRet}</div><div class="payer-stat-sub">${Object.values(byDay).filter(d=>d.retCount>0).length} dia(s) com retorno</div></div>
    <div class="payer-stat"><div class="payer-stat-lbl">${ICONS.calendar} Dias analisados</div><div class="payer-stat-val">${Object.keys(byDay).length}</div><div class="payer-stat-sub">dias úteis com NFs</div></div>
    <div class="payer-stat"><div class="payer-stat-lbl">${ICONS.chart} Total pagadores/dia</div><div class="payer-stat-val">${Object.values(byDay).reduce((s,d)=>s+d.entries.length,0)}</div><div class="payer-stat-sub">entradas únicas</div></div>`;

  renderPayerDayList(pfx);
}

function renderPayerDayList(pfx){
  const ps=PAYER_STATE[pfx];if(!ps)return;
  const catFilter=$(`pay-filter-cat-${pfx}`)?.value||'';
  const container=$(`pay-day-list-${pfx}`);if(!container)return;
  container.innerHTML='';
  const today=new Date().toISOString().slice(0,10);

  const dates=Object.keys(ps.byDay).filter(dt=>{
    if(dt>today)return false;
    const d=ps.byDay[dt];
    if(catFilter==='new'&&d.newCount===0)return false;
    if(catFilter==='ret'&&d.retCount===0)return false;
    return true;
  }).sort((a,b)=>b.localeCompare(a));

  if(!dates.length){container.innerHTML='<div style="padding:32px;text-align:center;color:var(--text3)">Nenhum dia com pagadores neste filtro.</div>';return;}

  dates.forEach(dt=>{
    const d=ps.byDay[dt];
    // Day group header (clicável para expandir)
    const hdr=document.createElement('div');
    hdr.className='payer-day-row tot';
    hdr.style.cssText='cursor:pointer;user-select:none';
    const isToday=dt===today;
    hdr.innerHTML=`
      <div class="mono" style="font-weight:800;color:var(--text)">${fd(dt)}${isToday?'<span style="margin-left:6px;background:var(--gold);color:#000;padding:1px 5px;border-radius:3px;font-size:9px">HOJE</span>':''}<br><span style="font-size:10px;color:var(--text3);font-weight:400">${dn(dt)}</span></div>
      <div style="font-size:11px;color:var(--text2)">${d.awbCount} AWBs · ${d.entries.length} pagadores<br><span style="color:var(--success);font-weight:700">${d.newCount} novos</span> · <span style="color:var(--accent);font-weight:700">${d.retCount} retorno</span></div>
      <div class="mono" style="color:var(--text2)">${d.awbCount}</div>
      <div class="mono" style="color:var(--accent);font-weight:800">${R$(d.totalFrete)}</div>
      <div></div><div style="font-size:11px;color:var(--text3)">clique para ver</div>`;
    const detailDiv=document.createElement('div');
    detailDiv.style.display='none';
    detailDiv.dataset.dt=dt;
    hdr.onclick=()=>{
      const open=detailDiv.style.display==='none';
      detailDiv.style.display=open?'':'none';
      hdr.style.background=open?'rgba(249,115,22,.06)':'';
      if(open&&!detailDiv.dataset.rendered)renderPayerDayDetail(detailDiv,d,pfx);
    };
    container.appendChild(hdr);
    container.appendChild(detailDiv);
  });
}

function renderPayerDayDetail(container,dayData,pfx){
  container.dataset.rendered='1';
  const searchQ=($(`pay-search-${pfx}`)?.value||'').trim().toLowerCase();
  const entries=[...dayData.entries].filter(e=>{
    if(!searchQ)return true;
    return (e.pagador||'').toLowerCase().includes(searchQ)||(e.cnpj||'').includes(searchQ);
  });
  entries.sort((a,b)=>(a.cat==='new'?0:a.cat==='ret'?1:2)-(b.cat==='new'?0:b.cat==='ret'?1:2));
  container.innerHTML=entries.map((e,i)=>{
    const isNew=e.cat==='new';const isRet=e.cat==='ret';
    const statusHtml=isNew?`<span style="background:rgba(16,185,129,.18);color:var(--success);padding:2px 9px;border-radius:6px;font-size:11px;font-weight:700">Novo</span>`:
      isRet?`<span style="background:rgba(249,115,22,.18);color:var(--accent);padding:2px 9px;border-radius:6px;font-size:11px;font-weight:700">Retornou</span>`:
      '<span style="color:var(--text3);font-size:11px">Recorrente</span>';
    const lastFmt=e.lastDate?new Date(e.lastDate+'T12:00:00').toLocaleDateString('pt-BR'):'1ª vez';
    const cnpjDisplay=e.cnpj?e.cnpj.replace(/^0+/,'')||e.cnpj:'—';
    const bg=i%2===0?'':'rgba(255,255,255,.012)';
    return `<div class="payer-day-row" style="background:${bg}">
      <div class="mono" style="font-size:11px;color:var(--cyan)">${cnpjDisplay}</div>
      <div><div style="font-weight:700;font-size:12px">${e.pagador||'—'}</div><div style="font-size:10px;color:var(--text3)">${e.tipoFrete||'N/D'} · ${e.papel||''}</div></div>
      <div class="mono" style="text-align:center">${e.awbs.length}</div>
      <div class="mono" style="color:var(--accent);font-weight:800">${R$(e.totalFrete)}</div>
      <div>${statusHtml}</div>
      <div style="font-size:11px;color:var(--text2)">${lastFmt}</div>
    </div>`;
  }).join('');
  if(!entries.length)container.innerHTML='<div style="padding:14px 16px;color:var(--text3);font-size:12px">Nenhum resultado para o filtro.</div>';
}

function filterPayerDetail(pfx){
  // Re-render open detail divs
  const container=$(`pay-day-list-${pfx}`);if(!container)return;
  container.querySelectorAll('[data-rendered]').forEach(div=>{
    delete div.dataset.rendered;
    const dt=div.dataset.dt;
    const ps=PAYER_STATE[pfx];if(!ps||!ps.byDay[dt])return;
    renderPayerDayDetail(div,ps.byDay[dt],pfx);
  });
}

function exportPayerCSV(pfx){
  const ps=PAYER_STATE[pfx];if(!ps)return;
  let csv='Data,Dia,CNPJ,Pagador,Tipo Frete,Papel,AWBs,Frete Total,Status,Ultima vez\n';
  Object.entries(ps.byDay).sort((a,b)=>b[0].localeCompare(a[0])).forEach(([dt,d])=>{
    d.entries.forEach(e=>{
      const status=e.cat==='new'?'Novo':e.cat==='ret'?'Retornou':'Recorrente';
      csv+=`"${dt}","${dn(dt)}","${e.cnpj||''}","${e.pagador}","${e.tipoFrete||''}","${e.papel||''}",${e.awbs.length},${e.totalFrete.toFixed(2)},"${status}","${e.lastDate||''}"\n`;
    });
  });
  const el=document.createElement('a');
  el.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));
  el.download=`pagadores_${pfx}_${_curY}-${padM(_mo)}.csv`;
  el.click();
}

// Hook renderFilial para injetar sub-abas na primeira renderização
const _origRenderFilial=renderFilial;
renderFilial=function(filial){
  _origRenderFilial(filial);
  const pfx=filial.toLowerCase();
  // Injeta sub-abas se ainda não estiver injetado
  if(!$(`payer-panel-${pfx}`)){
    injectPayersTabIntoFilialPane(filial);
  }
  // Resetar estado (dados podem ter mudado com novo mês)
  if(PAYER_STATE[pfx])PAYER_STATE[pfx].built=false;
};
