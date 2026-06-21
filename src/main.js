'use strict';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import dataXlsxUrl from '../data.xlsx?url';

/* ══════════════════════════════════════════════════
   CONFIGURATION
   ══════════════════════════════════════════════════
   The dashboard reads data from data.xlsx, located in the project
   root right next to the src/ folder — bundled inside this project,
   no external network call, no CORS, no API key, no Google account
   permissions needed. Vite resolves the import below at build time
   and copies the file into the deployed output automatically.

   TO UPDATE THE DATA:
   1. Open data.xlsx (in the project root) in Excel / Google Sheets / LibreOffice
   2. Replace its contents with your real export (keep the same
      column headers — see README.md for the expected columns)
   3. Save the file back to data.xlsx in the project root
   4. Run `npm run build` and redeploy

   The dashboard re-reads this file on every page load and every
   REFRESH_MS milliseconds, so once redeployed, the new numbers
   appear automatically without further code changes.
══════════════════════════════════════════════════ */
const REFRESH_MS = 60000; // re-read the file every 60 seconds

/* ── STATE ─────────────────────────────────────── */
let activeTab = 'overall';
let activeMonth = 'all';
let liveData = null; // populated from data.xlsx when available
const MONTHS = ['Jan','Feb','March','April','May','June','July','August','Sept','Oct','Nov','Dec'];
const MLBLS  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const PRODS  = ['Education','Endowment','Family Protection'];
const PC     = ['#2563eb','#1d4ed8','#93c5fd']; // education=blue, endowment=dark blue, fp=light
const PIE_COLORS = PC; // alias used by pie/donut charts
const CH = {};

/* Map raw sheet rows to dashboard DB format.
   IMPORTANT: Column names below must match your actual sheet headers.
   Common columns expected: REGION, PREMIUM, PRODUCT, Month, Premiumpayingterm,
   SALESBRANCH, SC_NAME, SC_VALUE_TARGET, SC_VOL_TARGET */
function rowsToDB(rows) {
  if (!rows || !rows.length) return null;

  // Detect available columns
  const sample = rows[0];
  const cols = Object.keys(sample);
  const find = (...candidates) => candidates.find(c => cols.includes(c)) || null;

  const COL_REGION   = find('REGION','Region','region','CLUSTER');
  const COL_PREM     = find('PREMIUM','Premium','premium','PREMIUM_COLLECTED','premium_collected');
  const COL_PRODUCT  = find('PRODUCT','Product','product');
  const COL_MONTH    = find('Month','MONTH','month');
  const COL_TERM     = find('Premiumpayingterm','TERM','premium_paying_term','PPT');
  const COL_BRANCH   = find('SALESBRANCH','salesbranch','Branch','BRANCH','branch');
  const COL_SC       = find('SALES_CONSULTANT','SC','Sales_Consultant','SALES CONSULTANT');

  if (!COL_REGION && !COL_BRANCH) return null; // sheet columns don't match — use fallback

  const clusterData = {};
  const branchMap   = {};
  const prodTotals  = { Education:{q:0,v:0}, Endowment:{q:0,v:0}, 'Family Protection':{q:0,v:0} };
  const monthlyPrem = {};
  const monthlyVol  = {};
  const termDist    = {lt10:0,gte10:0};

  MONTHS.forEach(m => { monthlyPrem[m]=0; monthlyVol[m]=0;
    ['Education','Endowment','Family Protection'].forEach(p => {
      if (!prodTotals[p][m]) prodTotals[p][m]={v:0,q:0};
    });
  });

  rows.forEach(row => {
    const region  = COL_REGION  ? String(row[COL_REGION]  ?? '').trim() : '';
    const prem    = parseFloat(COL_PREM ? row[COL_PREM] : 0) || 0;
    const product = COL_PRODUCT ? String(row[COL_PRODUCT] ?? '').trim() : '';
    const month   = COL_MONTH   ? String(row[COL_MONTH]   ?? '').trim() : '';
    const term    = parseFloat(COL_TERM ? row[COL_TERM] : 0) || 0;
    const branch  = COL_BRANCH  ? String(row[COL_BRANCH]  ?? '').trim() : '';

    // Cluster aggregation
    if (region && region !== 'Error' && DB.clusterTargets[region]) {
      if (!clusterData[region]) {
        clusterData[region] = {};
        MONTHS.forEach(m => clusterData[region][m]={v:0,q:0});
      }
      if (month && clusterData[region][month] !== undefined) {
        clusterData[region][month].v += prem;
        clusterData[region][month].q += 1;
      }
    }

    // Monthly totals
    if (month && monthlyPrem[month] !== undefined) {
      monthlyPrem[month] += prem;
      monthlyVol[month]  += 1;
    }

    // Product totals
    const normProd = PRODS.find(p => product.toLowerCase().includes(p.toLowerCase()));
    if (normProd && month && prodTotals[normProd][month] !== undefined) {
      prodTotals[normProd][month].v += prem;
      prodTotals[normProd][month].q += 1;
      prodTotals[normProd].v += prem;
      prodTotals[normProd].q += 1;
    }

    // Term distribution
    if (term >= 10) termDist.gte10++;
    else if (term > 0) termDist.lt10++;

    // Branch aggregation
    if (branch) {
      if (!branchMap[branch]) branchMap[branch] = {name:branch,v:0,q:0,lt10:0,gte10:0,products:{Education:{v:0,q:0},Endowment:{v:0,q:0},'Family Protection':{v:0,q:0}}};
      branchMap[branch].v += prem;
      branchMap[branch].q += 1;
      if (term >= 10) branchMap[branch].gte10++; else if (term > 0) branchMap[branch].lt10++;
      if (normProd) { branchMap[branch].products[normProd].v += prem; branchMap[branch].products[normProd].q += 1; }
    }
  });

  return {
    clusterData: Object.keys(DB.clusterTargets).reduce((acc,k) => {
      acc[k] = clusterData[k] || Object.fromEntries(MONTHS.map(m=>[m,{v:0,q:0}]));
      return acc;
    }, {}),
    monthlyPrem, monthlyVol, prodTotals, termDist,
    branchData: Object.values(branchMap).sort((a,b)=>b.v-a.v),
    totalRecords: rows.length,
    // Keep existing targets & maps (not in transmittal sheet)
    clusterTargets: DB.clusterTargets,
    rmTargets: DB.rmTargets, rmData: DB.rmData, rmTlMap: DB.rmTlMap,
    tlTargets: DB.tlTargets, tlData: DB.tlData, tlScMap: DB.tlScMap,
    scTargets: DB.scTargets, scData: DB.scData,
  };
}

/* ══════════════════════════════════════════════════
   SYNC — loads the Excel workbook bundled at ../data.xlsx
   (project root, next to src/). This file ships inside the deployed
   app itself, so there is no external network call, no CORS, and no
   third-party dependency. To update the data: replace data.xlsx in
   the project root with your export, rebuild (npm run build), and
   redeploy. The dashboard re-reads the file automatically on every
   page load and on the refresh interval.
══════════════════════════════════════════════════ */
const DATA_FILE_URL = dataXlsxUrl;

async function fetchSheet() {
  const dot = document.getElementById('live-dot');
  const lbl = document.getElementById('sync-lbl');
  const banner = document.getElementById('sync-banner');

  if (dot) dot.className = 'live-dot syncing';
  if (lbl) lbl.textContent = 'Loading…';

  try {
    const res = await fetch(DATA_FILE_URL + '?t=' + Date.now(), { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status + ' — data.xlsx could not be loaded');
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (!rows || !rows.length) throw new Error('Workbook has no data rows');
    const built = rowsToDB(rows);
    if (!built) throw new Error('Column mapping failed. Sheet columns: ' + Object.keys(rows[0]).join(', '));
    Object.assign(DB, built);
    if (banner) banner.classList.add('hidden');
    if (dot) dot.className = 'live-dot';
    if (lbl) lbl.textContent = 'Live · ' + new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    showToast('Loaded from data.xlsx');
    render();
  } catch (e) {
    console.error('[Excel]', e.message);
    if (dot) dot.className = 'live-dot error';
    if (lbl) lbl.textContent = 'Load error';
    if (banner) {
      banner.style.cssText = 'background:#fef2f2;border:1.5px solid #fca5a5;border-radius:8px;padding:10px 16px;font-size:11.5px;color:#991b1b;display:block';
      banner.classList.remove('hidden');
      banner.innerHTML = '\u26a0 Could not load data.xlsx: ' + e.message
        + ' <a href="#" id="retry-link" style="margin-left:8px;color:#1d4ed8;font-weight:600">Retry</a>';
      const retryLink = document.getElementById('retry-link');
      if (retryLink) retryLink.addEventListener('click', (ev) => { ev.preventDefault(); manualRefresh(); });
    }
    setTimeout(() => {
      if (dot) dot.className = 'live-dot';
      if (lbl) lbl.textContent = 'Live';
    }, 6000);
  }
}

function manualRefresh() { fetchSheet(); }
window.manualRefresh = manualRefresh;

function mkChart(id, type, labels, datasets, opts={}) {
  if (CH[id]) { CH[id].destroy(); delete CH[id]; }
  const ctx = document.getElementById(id);
  if (!ctx) return;
  const dfOpts = {
    responsive:true, maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtV(c.raw)}}},
    scales:{x:{grid:{display:false},ticks:{font:{size:10},color:'#64748b'}},
            y:{grid:{color:'#f1f5f9'},ticks:{font:{size:10},color:'#64748b',callback:v=>fmtV(v)}}}
  };
  CH[id] = new Chart(ctx, {type, data:{labels,datasets}, options:Object.assign({},dfOpts,opts)});
}

/* ── FORMATTERS ────────────────────────────────── */
function fmtV(n) {
  if (n>=1e9) return 'KSh '+(n/1e9).toFixed(1)+'B';
  if (n>=1e6) return 'KSh '+(n/1e6).toFixed(1)+'M';
  if (n>=1e3) return 'KSh '+(n/1e3).toFixed(0)+'K';
  return 'KSh '+n.toFixed(0);
}
function fmtN(n) { return n>=1e6?(n/1e6).toFixed(1)+'M':n>=1e3?(n/1e3).toFixed(0)+'K':String(n); }
function fmtNum(n) { return Number(n).toLocaleString(); }
function pct(a,b) { return b>0?Math.round(a/b*100):0; }
function wt(vp,qp) { return Math.round((vp+qp)/2); }
function pillCls(p) { return p>=80?'G':p>=50?'A':'R'; }
function initials(name) { return name.split(' ').map(w=>w[0]).join('').toUpperCase().slice(0,2); }

function getMs() {
  return activeMonth === 'all' ? MONTHS.slice(0, 4) : [activeMonth];
}

/* ── TABLE ROW BUILDERS ────────────────────────── */
function mkRow(name, va, vt, qa, qt, isBold, clickFn) {
  const vp=pct(va,vt), qp=pct(qa,qt), wp=wt(vp,qp);
  const cls = isBold ? 'bold' : '';
  const rowClass = clickFn ? 'tl-row-clickable' : '';
  const rowClick = clickFn ? `onclick="${clickFn}"` : '';
  return `<tr class="${rowClass}" ${rowClick}>
    <td class="${cls}">${name}</td>
    <td class="r">${fmtV(va)}</td><td class="r">${fmtV(vt)}</td>
    <td class="r"><span class="pill ${pillCls(vp)}">${vp}%</span></td>
    <td class="r">${qa}</td><td class="r">${qt}</td>
    <td class="r"><span class="pill ${pillCls(qp)}">${qp}%</span></td>
    <td class="r"><span class="pill ${pillCls(wp)}">${wp}%</span></td>
  </tr>`;
}
function mkRowNoBar(name, va, vt, qa, qt, clickFn) {
  return mkRow(name, va, vt, qa, qt, false, clickFn);
}
function mkFoot(label, rows, hasBarcol) {
  const tva=rows.reduce((s,r)=>s+r.va,0), tvt=rows.reduce((s,r)=>s+r.vt,0);
  const tqa=rows.reduce((s,r)=>s+r.qa,0), tqt=rows.reduce((s,r)=>s+r.qt,0);
  const vp=pct(tva,tvt), qp=pct(tqa,tqt), wp=wt(vp,qp);
  return `<tr><td>${label}</td><td class="r">${fmtV(tva)}</td><td class="r">${fmtV(tvt)}</td>
    <td class="r"><span class="pill ${pillCls(vp)}">${vp}%</span></td>
    <td class="r">${tqa}</td><td class="r">${tqt}</td>
    <td class="r"><span class="pill ${pillCls(qp)}">${qp}%</span></td>
    <td class="r"><span class="pill ${pillCls(wp)}">${wp}%</span></td>
    ${hasBarcol?'<td></td>':''}</tr>`;
}

/* ── DONUT HELPER ──────────────────────────────── */
function mkDonut(svgId, legId, nId, lId, slices) {
  const svg = document.getElementById(svgId);
  const leg = document.getElementById(legId);
  if (!svg || !leg) return;
  const total = slices.reduce((s,x)=>s+x.v,0);
  let off=0; const r=40, cx=50, cy=50, circ=2*Math.PI*r;
  svg.innerHTML = slices.map(s=>{
    const frac=s.v/total, dash=circ*frac, gap=circ*(1-frac);
    const el=`<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${s.c}" stroke-width="12" stroke-dasharray="${dash} ${circ}" stroke-dashoffset="${-off}" stroke-linecap="round"/>`;
    off+=circ*frac; return el;
  }).join('');
  if (nId) document.getElementById(nId).textContent = total>=1000?Math.round(total/1000)+'K':total;
  if (lId) document.getElementById(lId).textContent = 'total';
  leg.innerHTML = slices.map(s=>`<div class="leg-row"><div class="leg-dot" style="background:${s.c}"></div><span>${s.label}</span><span class="leg-pct">${Math.round(s.v/total*100)}%</span></div>`).join('');
}

/* ══ PAGE: OVERALL ═══════════════════════════════ */
function renderOverall() {
  const ms = getMs();
  let tva=0,tvt=0,tqa=0,tqt=0;
  Object.keys(DB.clusterTargets).forEach(c=>{
    tva+=ms.reduce((s,m)=>(s+(DB.clusterData[c]||{})[m]?.v||0),0);
    tvt+=DB.clusterTargets[c].v*(ms.length/12);
    tqa+=ms.reduce((s,m)=>(s+(DB.clusterData[c]||{})[m]?.q||0),0);
    tqt+=DB.clusterTargets[c].q*(ms.length/12);
  });
  const vp=pct(tva,tvt),qp=pct(tqa,tqt),wp=wt(vp,qp);
  document.getElementById('ov-val').textContent=fmtV(tva);
  document.getElementById('ov-vsub').textContent=`${vp}% of target`;
  document.getElementById('ov-vol').textContent=tqa;
  document.getElementById('ov-qsub').textContent=`${qp}% of target`;
  document.getElementById('ov-wt').textContent=wp+'%';
  document.getElementById('ov-rec').textContent=DB.totalRecords.toLocaleString();
  setTimeout(()=>{
    document.getElementById('ov-vbar').style.width=Math.min(100,vp)+'%';
    document.getElementById('ov-qbar').style.width=Math.min(100,qp)+'%';
    document.getElementById('ov-wbar').style.width=Math.min(100,wp)+'%';
  },100);
  const TQA=ms.reduce((s,m)=>s+(DB.monthlyVol[m]||0),0);
  document.getElementById('ov-meta').textContent=TQA+' records'+(activeMonth!=='all'?' · '+activeMonth:'');
  const clRows = Object.keys(DB.clusterTargets).map(c=>{
    const tgt=DB.clusterTargets[c];
    let va=0,qa=0;
    ms.forEach(m=>{const d=(DB.clusterData[c]||{})[m]||{v:0,q:0};va+=d.v;qa+=d.q;});
    return {name:c,va,vt:tgt.v,qa,qt:tgt.q};
  });
  document.getElementById('ov-tbody').innerHTML=clRows.map(r=>mkRow(r.name,r.va,r.vt,r.qa,r.qt,false,'')).join('');
  document.getElementById('ov-tfoot').innerHTML=mkFoot('Total',clRows,true);
  const mLabels=MONTHS.slice(0,4);
  const prodDs=PRODS.map((p,i)=>({label:p,backgroundColor:PC[i],borderRadius:3,data:mLabels.map(m=>DB.prodTotals[p][m]?.q||0)}));
  mkChart('c-prod','bar',mLabels,prodDs,{plugins:{legend:{display:false}},scales:{x:{stacked:true,grid:{display:false},ticks:{font:{size:10}}},y:{stacked:true,grid:{color:'#f1f5f9'},ticks:{font:{size:10},callback:v=>v}}}});
  mkChart('c-prem','bar',mLabels,[{label:'Premium',backgroundColor:['#dc2626','#ef4444','#f87171','#fca5a5'],borderRadius:4,data:mLabels.map(m=>DB.monthlyPrem[m]||0)}]);
  mkChart('c-vol','bar',mLabels,[{label:'Volume',backgroundColor:'#7c3aed',borderRadius:4,data:mLabels.map(m=>DB.monthlyVol[m]||0)}],{scales:{y:{ticks:{callback:v=>v,font:{size:10}}}}});
  mkDonut('dn-svg','dn-leg','dn-n','dn-l',[
    {v:DB.termDist.gte10,c:'#dc2626',label:'≥10 years'},
    {v:DB.termDist.lt10,c:'#fca5a5',label:'<10 years'}
  ]);
}

/* ══ PAGE: REGIONAL MANAGERS ════════════════════ */
function renderRM() {
  const ms = getMs();
  const rmOrder = Object.keys(DB.rmTargets);
  const rows = rmOrder.map(name=>{
    const tgt=DB.rmTargets[name]; let va=0,qa=0;
    ms.forEach(m=>{const d=(DB.rmData[name]||{})[m]||{v:0,q:0};va+=d.v;qa+=d.q;});
    return {name,va,vt:tgt.v,qa,qt:tgt.q};
  });
  document.getElementById('rm-tbody').innerHTML=rows.map(r=>mkRow(r.name,r.va,r.vt,r.qa,r.qt,false,'')).join('');
  document.getElementById('rm-tfoot').innerHTML=mkFoot('Total',rows,true);
  const TQA=rows.reduce((s,r)=>s+r.qa,0);
  document.getElementById('rm-meta').textContent=(activeMonth==='all'?'YTD':'Month: '+activeMonth)+' · '+TQA+' records';
  const allMs=MONTHS.slice(0,4);
  const prodDs=PRODS.map((p,i)=>({label:p,backgroundColor:PC[i],borderRadius:3,data:allMs.map(m=>DB.prodTotals[p][m]?.q||0)}));
  mkChart('c-rm-prod','bar',allMs,prodDs,{plugins:{legend:{display:false}},scales:{x:{stacked:true,grid:{display:false},ticks:{font:{size:10}}},y:{stacked:true,grid:{color:'#f1f5f9'},ticks:{font:{size:10},callback:v=>v}}}});
  mkChart('c-rm-prem','bar',allMs,[{label:'Premium',backgroundColor:['#dc2626','#ef4444','#f87171','#fca5a5'],borderRadius:4,data:allMs.map(m=>DB.monthlyPrem[m]||0)}]);
  mkDonut('dn-svg2','dn-leg2','dn-n2','dn-l2',[
    {v:DB.termDist.gte10,c:'#dc2626',label:'≥10 years'},
    {v:DB.termDist.lt10,c:'#fca5a5',label:'<10 years'}
  ]);
}

/* ══ PAGE: TEAM LEADERS ════════════════════════ */
function getTLRm(tlName) {
  for (const [rm, tls] of Object.entries(DB.rmTlMap)) {
    if (tls.includes(tlName)) return rm;
  }
  return 'Unassigned';
}

function renderTL() {
  const tlOrder = Object.keys(DB.tlTargets);
  const mtdMs = activeMonth === 'all' ? ['Jan'] : [activeMonth];
  const ytdMs = MONTHS.slice(0, 4);
  const allMs = MONTHS.slice(0, 4);

  function tlRows(ms) {
    return tlOrder.map(name => {
      const tgt = DB.tlTargets[name]; let v=0, q=0;
      ms.forEach(m => { const d=(DB.tlData[name]||{})[m]||{v:0,q:0}; v+=d.v; q+=d.q; });
      return {name, va:v, vt:tgt.v, qa:q, qt:tgt.q};
    });
  }

  const mtdRows = tlRows(mtdMs);
  const ytdRows = tlRows(ytdMs);
  const totalVA = ytdRows.reduce((s,r)=>s+r.va,0);
  const totalVT = ytdRows.reduce((s,r)=>s+r.vt,0);
  const totalQA = ytdRows.reduce((s,r)=>s+r.qa,0);
  const totalQT = ytdRows.reduce((s,r)=>s+r.qt,0);
  const overallWt = wt(pct(totalVA,totalVT), pct(totalQA,totalQT));
  const beating = ytdRows.filter(r=>pct(r.va,r.vt)>=80).length;

  document.getElementById('tl-summary-kpis').innerHTML = `
    <div class="tl-kpi"><div class="tl-kpi-label">Team Leaders</div><div class="tl-kpi-value">${tlOrder.length}</div><div class="tl-kpi-sub">Active team leads</div></div>
    <div class="tl-kpi"><div class="tl-kpi-label">YTD Value</div><div class="tl-kpi-value">${fmtV(totalVA)}</div><div class="tl-kpi-sub">${pct(totalVA,totalVT)}% of target</div></div>
    <div class="tl-kpi"><div class="tl-kpi-label">YTD Volume</div><div class="tl-kpi-value">${totalQA}</div><div class="tl-kpi-sub">${pct(totalQA,totalQT)}% of target</div></div>
    <div class="tl-kpi"><div class="tl-kpi-label">On Track (≥80%)</div><div class="tl-kpi-value">${beating} / ${tlOrder.length}</div><div class="tl-kpi-sub">Weighted avg: ${overallWt}%</div></div>
  `;

  const cardColors = ['#dc2626','#7c3aed','#2563eb','#16a34a','#b45309','#0891b2','#9333ea','#be185d','#1d4ed8','#15803d'];
  document.getElementById('tl-cards-grid').innerHTML = ytdRows.map((r, i) => {
    const vp=pct(r.va,r.vt), qp=pct(r.qa,r.qt), wp=wt(vp,qp);
    const color = cardColors[i % cardColors.length];
    const rm = getTLRm(r.name);
    const scCount = (DB.tlScMap[r.name] || []).length;
    const barColor = wp>=80?'#16a34a':wp>=50?'#b45309':'#dc2626';
    return `<div class="tl-card" onclick="openTLDetail('${r.name.replace(/'/g,"\\'")}')">
      <div class="tl-card-avatar" style="background:${color}1a;color:${color}">${initials(r.name)}</div>
      <div class="tl-card-name">${r.name}</div>
      <div class="tl-card-rm">RM: ${rm}</div>
      <div class="tl-card-stats">
        <div class="tl-stat-row"><span class="tl-stat-label">Value YTD</span><span class="tl-stat-val">${fmtV(r.va)}</span></div>
        <div class="tl-stat-row"><span class="tl-stat-label">Volume</span><span class="tl-stat-val">${r.qa} / ${r.qt}</span></div>
        <div class="tl-stat-row"><span class="tl-stat-label">Weighted %</span><span class="tl-stat-val" style="color:${barColor}">${wp}%</span></div>
        <div class="tl-stat-row"><span class="tl-stat-label">Team Size</span><span class="tl-stat-val">${scCount} SC${scCount!==1?'s':''}</span></div>
      </div>
      <div class="tl-card-bar"><div class="tl-card-bar-fill" style="width:${Math.min(100,wp)}%;background:${barColor}"></div></div>
    </div>`;
  }).join('');

  document.getElementById('tl-mtd-tbody').innerHTML = mtdRows.map(r=>mkRowNoBar(r.name,r.va,r.vt,r.qa,r.qt,`openTLDetail('${r.name.replace(/'/g,"\\'")}')` )).join('');
  document.getElementById('tl-mtd-tfoot').innerHTML = mkFoot('Total', mtdRows, false);
  document.getElementById('tl-ytd-tbody').innerHTML = ytdRows.map(r=>mkRowNoBar(r.name,r.va,r.vt,r.qa,r.qt,`openTLDetail('${r.name.replace(/'/g,"\\'")}')` )).join('');
  document.getElementById('tl-ytd-tfoot').innerHTML = mkFoot('Total', ytdRows, false);

  mkChart('c-tl-vol','bar',allMs,[{label:'Volume',backgroundColor:'#dc2626',borderRadius:4,data:allMs.map(m=>DB.monthlyVol[m]||0)}],{scales:{y:{ticks:{callback:v=>v,font:{size:10}}}}});
  mkChart('c-tl-prem','bar',allMs,[{label:'Premium',backgroundColor:['#dc2626','#f97316','#f59e0b','#b91c1c'],borderRadius:4,data:allMs.map(m=>DB.monthlyPrem[m]||0)}]);
}

function openTLDetail(tlName) {
  document.getElementById('tl-list-view').style.display = 'none';
  document.getElementById('tl-detail-view').style.display = 'block';
  const ytdMs = MONTHS.slice(0, 4);
  const tgt = DB.tlTargets[tlName] || {v:0, q:0};
  let ytdV=0, ytdQ=0;
  ytdMs.forEach(m => { const d=(DB.tlData[tlName]||{})[m]||{v:0,q:0}; ytdV+=d.v; ytdQ+=d.q; });
  const rm = getTLRm(tlName);
  const scs = DB.tlScMap[tlName] || [];
  const vp=pct(ytdV,tgt.v), qp=pct(ytdQ,tgt.q), wp=wt(vp,qp);
  const heroColor = wp>=80?'#16a34a':wp>=50?'#b45309':'#dc2626';
  document.getElementById('tl-detail-name-crumb').textContent = tlName;
  document.getElementById('tl-detail-avatar').textContent = initials(tlName);
  document.getElementById('tl-detail-name').textContent = tlName;
  document.getElementById('tl-detail-rm').textContent = rm;
  document.getElementById('tl-detail-sc-count').textContent = scs.length;
  document.getElementById('tl-detail-br-count').textContent = '5+';
  document.getElementById('tl-hero-kpis').innerHTML = `
    <div class="tl-hero-kpi"><div class="tl-hero-kpi-label">YTD Value</div><div class="tl-hero-kpi-value" style="color:var(--red)">${fmtV(ytdV)}</div><div class="tl-hero-kpi-sub">${vp}% of ${fmtV(tgt.v)} target</div></div>
    <div class="tl-hero-kpi"><div class="tl-hero-kpi-label">YTD Volume</div><div class="tl-hero-kpi-value" style="color:#7c3aed">${ytdQ}</div><div class="tl-hero-kpi-sub">${qp}% of ${tgt.q} target</div></div>
    <div class="tl-hero-kpi"><div class="tl-hero-kpi-label">Weighted %</div><div class="tl-hero-kpi-value" style="color:${heroColor}">${wp}%</div><div class="tl-hero-kpi-sub">${wp>=80?'On Track':wp>=50?'At Risk':'Needs Attention'}</div></div>
  `;
  document.getElementById('tl-perf-cards').innerHTML = [
    {label:'Value Actual', value:fmtV(ytdV), pct:vp, target:fmtV(tgt.v), color:'#dc2626'},
    {label:'Volume Actual', value:ytdQ, pct:qp, target:tgt.q+' policies', color:'#7c3aed'},
    {label:'Weighted Score', value:wp+'%', pct:wp, target:'Target: 100%', color:heroColor},
  ].map(m=>`<div class="tl-perf-card">
    <div class="tl-perf-card-label">${m.label}</div>
    <div class="tl-perf-card-value">${m.value}</div>
    <div class="tl-perf-card-pct">${m.pct}% · ${m.target}</div>
    <div class="tl-perf-bar"><div class="tl-perf-bar-fill" style="width:${Math.min(100,m.pct)}%;background:${m.color}"></div></div>
  </div>`).join('');
  document.getElementById('tl-monthly-tbody').innerHTML = ytdMs.map(m=>{
    const d=(DB.tlData[tlName]||{})[m]||{v:0,q:0};
    return d.v>0?`<tr><td>${m}</td><td class="r">${fmtV(d.v)}</td><td class="r">${d.q}</td></tr>`:'';
  }).join('')||`<tr><td colspan="3" style="text-align:center;color:#94a3b8;padding:10px">No data</td></tr>`;
  const scColors=['#1d4ed8','#7c3aed','#0891b2','#15803d','#b45309'];
  const scVolData = scs.map(sc=>{let v=0,q=0;ytdMs.forEach(m=>{const d=(DB.scData[sc]||{})[m]||{v:0,q:0};v+=d.v;q+=d.q;});return {sc,v,q};});
  document.getElementById('tl-sc-list').innerHTML = scVolData.map((item,i)=>{
    const scTgt=DB.scTargets[item.sc]||{v:0,q:0};
    const scVp=pct(item.v,scTgt.v);
    const pc=scVp>=80?'#16a34a':scVp>=50?'#b45309':'#dc2626';
    return `<div class="tl-sc-item"><div class="tl-sc-avatar" style="background:${scColors[i%scColors.length]}1a;color:${scColors[i%scColors.length]}">${initials(item.sc)}</div><div class="tl-sc-info"><div class="tl-sc-name">${item.sc}</div><div class="tl-sc-sub">Vol: ${item.q} · ${fmtV(item.v)}</div></div><div class="tl-sc-pct" style="color:${pc}">${scVp}%</div></div>`;
  }).join('')||'<div class="tl-empty">No sales consultants mapped</div>';
  const topBranches=DB.branchData.slice().sort((a,b)=>b.v-a.v).slice(0,5);
  document.getElementById('tl-branch-list').innerHTML = topBranches.map(b=>`<div class="tl-branch-item"><div class="tl-branch-name"><div class="tl-branch-dot"></div>${b.name}</div><div class="tl-branch-right"><div class="tl-branch-val">${fmtV(b.v)}</div><div class="tl-branch-vol">${b.q} policies</div></div></div>`).join('');
  const volData=ytdMs.map(m=>(DB.tlData[tlName]||{})[m]?.q||0);
  const valData=ytdMs.map(m=>(DB.tlData[tlName]||{})[m]?.v||0);
  ['c-tl-detail-vol','c-tl-detail-val','c-tl-detail-sc'].forEach(id=>{if(CH[id]){CH[id].destroy();delete CH[id];}});
  const ctx1=document.getElementById('c-tl-detail-vol');
  if(ctx1) CH['c-tl-detail-vol']=new Chart(ctx1,{type:'bar',data:{labels:ytdMs,datasets:[{label:'Volume',backgroundColor:'#dc2626',borderRadius:4,data:volData}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{display:false},ticks:{font:{size:10}}},y:{grid:{color:'#f1f5f9'},ticks:{font:{size:10},callback:v=>v}}}}});
  const ctx2=document.getElementById('c-tl-detail-val');
  if(ctx2) CH['c-tl-detail-val']=new Chart(ctx2,{type:'line',data:{labels:ytdMs,datasets:[{label:'Value',borderColor:'#7c3aed',backgroundColor:'#7c3aed20',borderWidth:2,fill:true,tension:.4,pointRadius:4,pointBackgroundColor:'#7c3aed',data:valData}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtV(c.raw)}}},scales:{x:{grid:{display:false},ticks:{font:{size:10}}},y:{grid:{color:'#f1f5f9'},ticks:{font:{size:10},callback:v=>fmtV(v)}}}}});
  const ctx3=document.getElementById('c-tl-detail-sc');
  if(ctx3&&scVolData.length>0) CH['c-tl-detail-sc']=new Chart(ctx3,{type:'doughnut',data:{labels:scVolData.map(s=>s.sc),datasets:[{data:scVolData.map(s=>s.q||1),backgroundColor:scColors.slice(0,scVolData.length),borderWidth:2,borderColor:'#fff',hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10,padding:8}}}}});
  window.scrollTo({top:0,behavior:'smooth'});
}

function closeTLDetail() {
  document.getElementById('tl-detail-view').style.display = 'none';
  document.getElementById('tl-list-view').style.display = 'block';
}

/* ══ PAGE: SALES CONSULTANTS (REDESIGNED) ══════ */
function renderSC() {
  const ms = getMs();
  const scOrder = Object.keys(DB.scTargets);
  const ytdMs = MONTHS.slice(0, 4);

  // YTD rows
  const ytdRows = scOrder.map(name=>{
    const tgt=DB.scTargets[name]; let v=0,q=0;
    ytdMs.forEach(m=>{const d=(DB.scData[name]||{})[m]||{v:0,q:0};v+=d.v;q+=d.q;});
    return {name,va:v,vt:tgt.v,qa:q,qt:tgt.q};
  }).filter(r=>r.va>0||r.qa>0);

  const totalVA=ytdRows.reduce((s,r)=>s+r.va,0);
  const totalQA=ytdRows.reduce((s,r)=>s+r.qa,0);
  const totalVT=ytdRows.reduce((s,r)=>s+r.vt,0);
  const totalQT=ytdRows.reduce((s,r)=>s+r.qt,0);
  const overallVp=pct(totalVA,totalVT), overallQp=pct(totalQA,totalQT);
  document.getElementById('sc-ytd-meta').textContent = `${ytdRows.length} consultants · ${totalQA} policies · YTD`;

  document.getElementById('sc-ytd-tbody').innerHTML = ytdRows
    .sort((a,b)=>b.va-a.va)
    .map(r=>mkRowNoBar(r.name,r.va,r.vt,r.qa,r.qt,'')).join('');
  document.getElementById('sc-ytd-tfoot').innerHTML = mkFoot('Total',ytdRows,false);

  // ── Premium term horizontal bars ──
  const gte10 = DB.termDist.gte10;
  const lt10  = DB.termDist.lt10;
  const total = gte10 + lt10;
  const pGte = total>0?Math.round(gte10/total*100):0;
  const pLt  = total>0?Math.round(lt10/total*100):0;
  // Animate bars after paint
  setTimeout(()=>{
    const g = document.getElementById('sc-term-gte10-bar');
    const l = document.getElementById('sc-term-lt10-bar');
    if(g){ g.style.width=pGte+'%'; g.textContent=gte10; }
    if(l){ l.style.width=pLt+'%';  l.textContent=lt10; }
    const gp=document.getElementById('sc-term-gte10-pct');
    const lp=document.getElementById('sc-term-lt10-pct');
    if(gp) gp.textContent=pGte+'%';
    if(lp) lp.textContent=pLt+'%';
    const tp=document.getElementById('sc-term-totalpct');
    if(tp) tp.textContent=`Total: ${total} policies`;
  },150);

  // ── Product pie chart ──
  const prodData = PRODS.map(p=>DB.prodTotals[p].q||0);
  const prodTotal = prodData.reduce((s,v)=>s+v,0);
  if(CH['c-sc-prod-pie']){CH['c-sc-prod-pie'].destroy();delete CH['c-sc-prod-pie'];}
  const ctxPie = document.getElementById('c-sc-prod-pie');
  if(ctxPie){
    CH['c-sc-prod-pie'] = new Chart(ctxPie,{
      type:'pie',
      data:{
        labels: PRODS.map((p,i)=>{
          const cnt=prodData[i];
          const pct2=prodTotal>0?((cnt/prodTotal)*100).toFixed(2):0;
          return `${cnt} (${pct2}%)`;
        }),
        datasets:[{data:prodData,backgroundColor:PIE_COLORS,borderWidth:2,borderColor:'#fff',hoverOffset:6}]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        plugins:{
          legend:{
            position:'right',
            labels:{
              font:{size:11},boxWidth:12,padding:12,
              generateLabels: chart=>{
                return PRODS.map((p,i)=>({
                  text: p,
                  fillStyle: PIE_COLORS[i],
                  strokeStyle:'#fff',
                  lineWidth:2,
                  index:i
                }));
              }
            }
          },
          tooltip:{callbacks:{label:ctx=>{
            const cnt=prodData[ctx.dataIndex];
            const pct2=prodTotal>0?((cnt/prodTotal)*100).toFixed(2):0;
            return `${PRODS[ctx.dataIndex]}: ${cnt} (${pct2}%)`;
          }}}
        }
      }
    });
  }

  // ── Branch table ──
  const branches = (DB.branchData||[]).slice().sort((a,b)=>b.v-a.v);
  const bodyEl = document.getElementById('sc-branch-body');
  if(bodyEl){
    bodyEl.innerHTML = branches.map(b=>`
      <div class="sc-branch-row">
        <div class="sc-branch-name">${b.name}</div>
        <div class="sc-branch-val">${fmtNum(Math.round(b.v))}</div>
        <div class="sc-branch-val">${b.q}</div>
      </div>`).join('');
  }
  const tv=branches.reduce((s,b)=>s+b.v,0);
  const tq=branches.reduce((s,b)=>s+b.q,0);
  const tfv=document.getElementById('sc-branch-total-v');
  const tfq=document.getElementById('sc-branch-total-q');
  if(tfv) tfv.textContent=fmtNum(Math.round(tv));
  if(tfq) tfq.textContent=tq;
}

/* ══ PAGE: BRANCH PRODUCTION ════════════════════ */
function renderBP() {
  const sorted = DB.branchData.slice().sort((a,b)=>b.v-a.v);
  document.getElementById('bp-tbody').innerHTML=sorted.map(b=>`<tr>
    <td class="bold">${b.name}</td>
    <td class="r">${fmtV(b.v)}</td>
    <td class="r">${b.q}</td>
    <td class="r">${b.products.Education.q>0?b.products.Education.q:'-'}</td>
    <td class="r">${b.products.Endowment.q>0?b.products.Endowment.q:'-'}</td>
    <td class="r">${b.products['Family Protection'].q>0?b.products['Family Protection'].q:'-'}</td>
  </tr>`).join('');
  const tv=sorted.reduce((s,b)=>s+b.v,0), tq=sorted.reduce((s,b)=>s+b.q,0);
  document.getElementById('bp-tfoot').innerHTML=`<tr><td>Total (${sorted.length} branches)</td><td class="r">${fmtV(tv)}</td><td class="r">${tq}</td><td class="r">—</td><td class="r">—</td><td class="r">—</td></tr>`;
  mkChart('c-bp-term','bar',['≥10 yrs','<10 yrs'],[{backgroundColor:['#2563eb','#93c5fd'],borderRadius:6,borderWidth:0,data:[DB.termDist.gte10,DB.termDist.lt10]}],{indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{grid:{display:false},ticks:{font:{size:10}}},y:{grid:{display:false},ticks:{font:{size:10}}}}});
  mkDonut('dn-bp','dn-bp-leg','dn-bp-n','dn-bp-l',PRODS.map((p,i)=>({v:DB.prodTotals[p].q,c:PIE_COLORS[i]||PC[i],label:p})));
}

/* ══ ROUTING ════════════════════════════════════ */
function render() {
  if (activeTab==='overall') renderOverall();
  else if (activeTab==='rm') renderRM();
  else if (activeTab==='tl') renderTL();
  else if (activeTab==='sc') renderSC();
  else if (activeTab==='bp') renderBP();
}

/* ══ TAB CLICKS ═════════════════════════════════ */
document.querySelectorAll('.tab').forEach(t => {
  t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));
    t.classList.add('active');
    activeTab = t.dataset.tab;
    document.getElementById('page-'+activeTab).classList.add('active');
    if (activeTab !== 'tl') closeTLDetail();
    render();
  });
});

/* ══ MONTH FILTER ════════════════════════════════ */
document.querySelectorAll('.mb').forEach(b => {
  b.addEventListener('click', () => {
    document.querySelectorAll('.mb').forEach(x=>x.classList.remove('on'));
    b.classList.add('on');
    activeMonth = b.dataset.m;
    closeTLDetail();
    render();
  });
});

/* ══ TOAST ════════════════════════════════════════ */
function showToast(msg) {
  const t=document.getElementById('toast');
  const m=document.getElementById('toast-msg');
  if(m) m.textContent=msg;
  t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),2500);
}

/* ══ DATE / INIT ════════════════════════════════ */
function setDate() {
  const d = new Date();
  document.getElementById('hdr-date').textContent = d.toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric'});
}
setDate();

/* ══ BOOT ════════════════════════════════════════ */
render(); // render with static data immediately

// Then try to fetch live data
fetchSheet();

// Poll every REFRESH_MS
setInterval(fetchSheet, REFRESH_MS);

/* ══════════════════════════════════════════════════
   GLOBAL EXPOSURE
   These functions are invoked via inline onclick="" attributes
   generated through innerHTML, so they must be attached to window
   for Vite's module-scoped build to find them.
══════════════════════════════════════════════════ */
window.openTLDetail = openTLDetail;
window.closeTLDetail = closeTLDetail;
