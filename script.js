/* ══════════════════════════════════════════
   CONFIGURAÇÃO — URLs do Google Sheets
   Publicar cada aba: Arquivo → Publicar na web → CSV
══════════════════════════════════════════ */
const BASE = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTj3u5RIJkXWM4DJP_vYx5VjJMC_PpHdk6C62daBJJE0hk1NJhB86UdahFIDB7AEUxZiJ5OEiVu5c_u/pub?single=true&output=csv&gid=';

const SHEETS = {
  SEMANAL:    BASE + '0',
  EZ_TICKETS: BASE + '351627180',
  METAS:      BASE + '956110426'
  // EZ_RESUMO, EZ_CLASS, EZ_HEATMAP, EZ_PERF removidas — tudo calculado de EZ_TICKETS
};

/* ── DADOS EM MEMÓRIA ── */
let SEMANAL_RAW  = {};
let EZ_TICKETS   = [];
let METAS_RAW    = [];

const SC = {green:'#1E7A42', yellow:'#966A00', red:'#B82418', gray:'#9BA8B0'};
let SEM = [];

/* ── PARSER CSV GENÉRICO ── */
function parseCSV(text, sep=',') {
  const lines = text.trim().split(/\r?\n/);
  const start = lines[0].startsWith('sep=') ? 1 : 0;
  const headers = lines[start].split(sep).map(h => h.replace(/^"|"$/g,'').trim());
  return lines.slice(start+1).filter(l=>l.trim()).map(line => {
    const vals = [];
    let cur = '', inQ = false;
    for (let c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === sep && !inQ) { vals.push(cur.trim()); cur = ''; }
      else { cur += c; }
    }
    vals.push(cur.trim());
    const obj = {};
    headers.forEach((h,i) => { obj[h] = vals[i] !== undefined ? vals[i].replace(/^"|"$/g,'') : ''; });
    return obj;
  });
}

/* ── PARSER NUMÉRICO pt-BR ── */
function parseNum(s) {
  if (s === null || s === undefined) return 0;
  s = String(s).trim().replace(/^R\$\s*/, '').replace(/%\s*$/, '').replace(/\s/g,'');
  if (!s) return 0;
  // Remover pontos de milhar (ponto seguido de exatamente 3 dígitos)
  s = s.replace(/\.(?=\d{3}(\D|$))/g, '');
  // Vírgula decimal → ponto
  s = s.replace(',', '.');
  return parseFloat(s) || 0;
}

/* ── PARSER SEMANAL ── */
function parseSemanal(text) {
  const lines = text.trim().split(/\r?\n/);

  // Detectar sep= e calcular skip exato
  let skip = 0;
  if (lines[0] && lines[0].startsWith('sep=')) skip = 1;
  // Após sep (ou início): linha vazia, título, cabeçalho-meses, cabeçalho-semanas = +4
  skip += 4;

  const dataLines = lines.slice(skip).filter(l => l.trim());
  const result = {};
  let currentInd = '';
  const TIPO_MAP = { 'Meta':'meta', 'Resultado':'res', 'Ano anterior':'anoAnt' };

  dataLines.forEach(line => {
    const vals = [];
    let cur = '', inQ = false;
    for (let c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { vals.push(cur.trim().replace(/^"|"$/g,'')); cur = ''; }
      else { cur += c; }
    }
    vals.push(cur.trim().replace(/^"|"$/g,''));

    const col0 = vals[0] ? vals[0].trim() : '';
    const col1 = vals[1] ? vals[1].trim() : '';

    if (col0) currentInd = col0;
    if (!currentInd) return;

    const tipo = TIPO_MAP[col1];
    if (!tipo) return;

    if (!result[currentInd]) result[currentInd] = {};
    for (let mes = 1; mes <= 12; mes++) {
      if (!result[currentInd][mes]) result[currentInd][mes] = {};
      for (let sem = 1; sem <= 4; sem++) {
        const col = 2 + (mes-1)*4 + (sem-1);
        if (!result[currentInd][mes][sem]) result[currentInd][mes][sem] = {meta:0, res:0, anoAnt:0};
        if (col < vals.length) {
          const v = parseNum(vals[col] || '');
          result[currentInd][mes][sem][tipo] = v;
        }
      }
    }
  });

  console.log('[Semanal] indicadores:', Object.keys(result));
  if (Object.keys(result).length > 0) {
    const s = Object.keys(result)[0];
    console.log('[Semanal] ex:', s, '| mes3 S1:', result[s]?.[3]?.[1]);
  }
  return result;
}

/* ── FILTRO: semanas por mês+semana ── */
function getWeeksForFilter(mes, sem) {
  // sem=0 → mês todo (4 semanas); sem=1..4 → semana específica
  if (!sem || sem === 0) return [1,2,3,4].map(s => ({mes, sem: s}));
  return [{mes, sem}];
}

/* ── FILTRO: range de datas para EZ tickets ── */
function getDateRangeForFilter(mes, sem) {
  const year = new Date().getFullYear();
  const pad  = n => String(n).padStart(2,'0');
  const last = new Date(year, mes, 0).getDate();
  if (!sem || sem === 0) return { de: `${year}-${pad(mes)}-01`, ate: `${year}-${pad(mes)}-${pad(last)}` };
  const dayStart = (sem-1)*7 + 1;
  const dayEnd   = sem === 4 ? last : sem*7;
  return { de: `${year}-${pad(mes)}-${pad(dayStart)}`, ate: `${year}-${pad(mes)}-${pad(dayEnd)}` };
}

/* ── SOMA SEMANAL ── */
function sumSemanal(indicador, weeks) {
  let meta = 0, res = 0, anoAnt = 0;
  weeks.forEach(({mes, sem}) => {
    const d = SEMANAL_RAW[indicador]?.[mes]?.[sem];
    if (d) { meta += d.meta||0; res += d.res||0; anoAnt += d.anoAnt||0; }
  });
  return {meta, res, anoAnt};
}

/* ── PARSER EZ TICKETS ── */
function processEZTickets(rows) {
  return rows.map(r => {
    const d = r['Data'] || '';
    let dataStr = '';
    if (d) {
      const [dia, mes, ano] = d.split('/');
      dataStr = `${ano}-${(mes||'').padStart(2,'0')}-${(dia||'').padStart(2,'0')}`;
    }
    return {
      DataStr: dataStr,
      Hora: parseInt(r['Hora']) || 0,
      Agente: r['Agente'] || '',
      Status: r['Status'] || '',
      Finalizado: r['Finalizado'] === '1' || r['Finalizado'] === 1,
      TPI_min: parseFloat(r['TPI_min']) || 0,
      TMA_min: parseFloat(r['TMA_min']) || 0,
      Classificacao: r['Classificacao_Principal'] || '',
      Ativo: r['Ativo'] || ''
    };
  }).filter(r => r.DataStr);
}



/* ── LOADING ── */
function setLoading(on) {
  const msg = document.getElementById('loading-msg');
  if (msg) msg.style.display = on ? 'flex' : 'none';
}

/* ── CARREGAMENTO PRINCIPAL ── */

/* ── PARSER ABA METAS ── */
// Estrutura: Agente | Indicador | Janeiro Meta | Janeiro Real | Fevereiro Meta | Fevereiro Real | ...
function processMetasSheet(rows) {
  const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                 'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const result = [];
  rows.forEach(r => {
    const agente    = (r['Agente']    || '').trim();
    const indicador = (r['Indicador'] || '').trim();
    if (!agente || !indicador || agente === 'Agente') return;
    MESES.forEach((m, i) => {
      const mes  = i + 1;
      const meta = parseNum(r[m + ' Meta'] || '0');
      const real = parseNum(r[m + ' Real'] || '0');
      result.push({ agente, indicador, mes, meta, real });
    });
  });
  console.log('[Metas] registros:', result.length, '| agentes:', [...new Set(result.map(d=>d.agente))]);
  return result;
}

/* ── HELPER: filtrar METAS_RAW ── */
function getMetasData(mes) {
  return METAS_RAW.filter(d => d.mes === mes);
}


/* ══ RENDER ABA METAS ══ */
let metasRendered = false;
function renderMetas() {
  // Sempre re-renderiza ao trocar mês
  const el = document.getElementById('metas-main');
  if (!el) return;

  const mes = parseInt(document.getElementById('f-mes')?.value) || (new Date().getMonth()+1);
  const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                 'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const nomeMes = MESES[mes-1];

  // Se não há dados ainda → skeleton de espera
  if (!METAS_RAW.length) {
    el.innerHTML = `
    <div style="padding:48px 32px;text-align:center;">
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:28px;font-weight:700;
        color:var(--gold);letter-spacing:1px;margin-bottom:12px;">METAS</div>
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:14px;color:var(--txt-faint);
        letter-spacing:1px;margin-bottom:32px;">Conecte a planilha de metas para visualizar os dados</div>
      <div style="display:inline-block;padding:12px 24px;border:1px solid rgba(180,165,140,0.3);
        border-radius:6px;font-family:'Barlow Condensed',sans-serif;font-size:13px;
        color:var(--txt-faint);letter-spacing:1px;">
        Cole o link CSV em <strong style="color:var(--gold);">SHEETS.METAS</strong> no script.js
      </div>
    </div>`;
    return;
  }

  const dados = getMetasData(mes);
  const indicadores = [...new Set(dados.map(d => d.indicador))];
  const agentes     = [...new Set(dados.filter(d => d.agente !== 'Time').map(d => d.agente))];
  const timeRow     = (ind) => dados.find(d => d.agente === 'Time' && d.indicador === ind) || {meta:0,real:0};

  const SC = {green:'#1E7A42', yellow:'#966A00', red:'#B82418', gray:'#9BA8B0'};
  function cor(real, meta) {
    if (!meta) return SC.gray;
    const p = real/meta*100;
    return p >= 100 ? SC.green : p >= 70 ? SC.yellow : SC.red;
  }
  function badge(real, meta) {
    if (!meta) return '—';
    return Math.round(real/meta*100) + '%';
  }
  function fmtV(ind, v) {
    if (ind === 'Faturamento' || ind === 'Receita') return 'R$ ' + Math.round(v).toLocaleString('pt-BR');
    if (ind === 'Litros' || ind === 'Litros vendidos') return Math.round(v).toLocaleString('pt-BR') + ' L';
    if (ind === 'Conversão') return v.toFixed(1) + '%';
    return Math.round(v).toLocaleString('pt-BR');
  }

  // ── Seção 1: Time ──
  let teamHTML = indicadores.map(ind => {
    const t   = timeRow(ind);
    const pct = t.meta ? Math.min(t.real/t.meta*100, 100) : 0;
    const c   = cor(t.real, t.meta);
    const bg  = c === SC.green ? 'rgba(30,122,66,0.12)' : c === SC.yellow ? 'rgba(150,106,0,0.12)' : 'rgba(184,36,24,0.10)';
    return `
    <div class="card line-l2" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header">
          <div class="c-title pill-l2">${ind}</div>
          <div class="c-sub">${nomeMes} · Time</div>
        </div>
        <div style="margin-top:14px;">
          <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px;">
            <div style="font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:700;color:var(--txt);">${fmtV(ind, t.real)}</div>
            <div style="font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:600;color:${c};">${badge(t.real,t.meta)}</div>
          </div>
          <div style="height:6px;background:rgba(180,165,140,0.15);border-radius:3px;overflow:hidden;">
            <div style="height:100%;width:${pct.toFixed(1)}%;background:${c};border-radius:3px;transition:width 0.8s ease;"></div>
          </div>
          <div style="display:flex;justify-content:space-between;margin-top:5px;">
            <span style="font-family:'Barlow Condensed',sans-serif;font-size:11px;color:var(--txt-faint);">meta: ${fmtV(ind, t.meta)}</span>
            <span style="font-family:'Barlow Condensed',sans-serif;font-size:11px;color:var(--txt-faint);">${t.real >= t.meta ? '✓ atingida' : 'falta ' + fmtV(ind, Math.max(0, t.meta - t.real))}</span>
          </div>
        </div>
      </div>
    </div>`;
  }).join('');

  // ── Seção 2: Agentes ──
  let agentesHTML = agentes.map(ag => {
    const rows = dados.filter(d => d.agente === ag);
    const cards = rows.map(r => {
      const pct = r.meta ? Math.min(r.real/r.meta*100,100) : 0;
      const c   = cor(r.real, r.meta);
      return `
        <div style="margin-bottom:14px;">
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:600;
            letter-spacing:0.8px;color:var(--txt-faint);text-transform:uppercase;margin-bottom:4px;">${r.indicador}</div>
          <div style="display:flex;align-items:center;gap:10px;">
            <div style="flex:1;height:8px;background:rgba(180,165,140,0.15);border-radius:4px;overflow:hidden;">
              <div style="height:100%;width:${pct.toFixed(1)}%;background:${c};border-radius:4px;transition:width 0.8s ease;"></div>
            </div>
            <span style="font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700;
              color:${c};min-width:38px;text-align:right;">${badge(r.real,r.meta)}</span>
          </div>
          <div style="display:flex;justify-content:space-between;margin-top:3px;">
            <span style="font-family:'Barlow Condensed',sans-serif;font-size:10px;color:var(--txt-faint);">${fmtV(r.indicador,r.real)} de ${fmtV(r.indicador,r.meta)}</span>
          </div>
        </div>`;
    }).join('');

    // Calcular % médio do agente para ranking
    const medPct = rows.reduce((s,r) => s + (r.meta ? r.real/r.meta*100 : 0), 0) / Math.max(rows.length, 1);
    const cAg    = cor(medPct, 100);

    return `
    <div class="card line-l3" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header">
          <div class="c-title pill-l3">${ag}</div>
          <div class="c-sub" style="color:${cAg};font-weight:600;">${Math.round(medPct)}% da meta</div>
        </div>
        <div style="margin-top:14px;">${cards}</div>
      </div>
    </div>`;
  }).join('');

  // ── Seção 3: Ranking ──
  const ranking = agentes.map(ag => {
    const rows  = dados.filter(d => d.agente === ag);
    const med   = rows.reduce((s,r) => s + (r.meta ? r.real/r.meta*100 : 0), 0) / Math.max(rows.length,1);
    return { ag, med };
  }).sort((a,b) => b.med - a.med);

  const medals = ['🥇','🥈','🥉'];
  const rankHTML = ranking.map((r,i) => {
    const c   = cor(r.med, 100);
    const bar = Math.min(r.med, 100).toFixed(1);
    return `
    <div style="display:flex;align-items:center;gap:14px;padding:10px 0;
      border-bottom:1px solid rgba(180,165,140,0.12);">
      <div style="font-size:20px;width:28px;text-align:center;">${medals[i]||''}</div>
      <div style="flex:1;">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:700;
          color:var(--txt);">${r.ag}</div>
        <div style="height:5px;background:rgba(180,165,140,0.15);border-radius:3px;margin-top:4px;overflow:hidden;">
          <div style="height:100%;width:${bar}%;background:${c};border-radius:3px;"></div>
        </div>
      </div>
      <div style="font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:700;
        color:${c};min-width:48px;text-align:right;">${Math.round(r.med)}%</div>
    </div>`;
  }).join('');

  el.innerHTML = `
  <!-- METAS TEAM -->
  <div class="row" style="grid-template-columns:repeat(${Math.min(indicadores.length,3)},1fr);">
    ${teamHTML}
  </div>

  <!-- METAS INDIVIDUAIS + RANKING -->
  <div class="row" style="grid-template-columns:${agentes.length > 0 ? '2fr 1fr' : '1fr'};">
    <div style="display:grid;grid-template-columns:repeat(${Math.min(agentes.length,3)},1fr);gap:16px;">
      ${agentesHTML}
    </div>
    <div class="card line-l3" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header">
          <div class="c-title pill-l3">Ranking ${nomeMes}</div>
          <div class="c-sub">% médio de atingimento</div>
        </div>
        <div style="margin-top:12px;">${rankHTML}</div>
      </div>
    </div>
  </div>`;
}


/* ══ DARK MODE ══ */
(function(){
  const DARK_KEY = 'germania-dark';
  const root = document.documentElement;

  function applyDark(on) {
    root.setAttribute('data-theme', on ? 'dark' : 'light');
    const btn = document.getElementById('btn-dark');
    if (btn) btn.textContent = on ? 'modo claro' : 'modo escuro';
    try { localStorage.setItem(DARK_KEY, on ? '1' : '0'); } catch(e){}
  }

  window.toggleDark = function() {
    const isDark = root.getAttribute('data-theme') === 'dark';
    applyDark(!isDark);
  };

  // Inicializar
  const saved = (() => { try { return localStorage.getItem(DARK_KEY); } catch(e){ return null; } })();
  applyDark(saved === '1');
})();

async function loadData() {
  setLoading(true);
  try {
    await Promise.all([
      fetch(SHEETS.SEMANAL).then(r=>r.text()).then(t=>{ SEMANAL_RAW = parseSemanal(t); }),
      fetch(SHEETS.EZ_TICKETS).then(r=>r.text()).then(t=>{ EZ_TICKETS = processEZTickets(parseCSV(t)); }),
      fetch(SHEETS.METAS).then(r=>r.text()).then(t=>{ METAS_RAW = processMetasSheet(parseCSV(t)); }),
    ]);
  } catch(e) { console.warn('Erro ao carregar dados:', e); }
  console.log('[EZ_TICKETS]', EZ_TICKETS.length, 'tickets | meses:', [...new Set(EZ_TICKETS.map(d=>d.DataStr.slice(0,7)))]);
  setLoading(false);
  go();
}

/* ── UTILITÁRIOS VISUAIS ── */
const CR=34,CCX=46,CCY=46,CSZ=92,CCIRC=2*Math.PI*CR;
function circ(elId,pct,color,txt){
  const el=document.getElementById(elId);if(!el)return;
  const cp=Math.min(Math.max(pct||0,0),100),off=CCIRC-(cp/100)*CCIRC;
  const uid='g'+elId;
  const cLight=color==='#9BA8B0'?'#D0D8DC':color==='#1E7A42'?'#6FD49A':color==='#966A00'?'#F5C050':color==='#B82418'?'#F08878':'#CCCCCC';
  el.innerHTML=`<svg width="${CSZ}" height="${CSZ}" viewBox="0 0 ${CSZ} ${CSZ}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="${uid}" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stop-color="${cLight}" stop-opacity="1"/>
        <stop offset="100%" stop-color="${color}" stop-opacity="1"/>
      </linearGradient>
      <filter id="${uid}sh">
        <feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="rgba(0,0,0,0.15)"/>
      </filter>
    </defs>
    <circle cx="${CCX}" cy="${CCY}" r="${CR+7}" fill="rgba(0,0,0,0.07)"/>
    <circle cx="${CCX}" cy="${CCY}" r="${CR+5}" class="circ-bg-fill" fill="white" filter="url(#${uid}sh)"/>
    <circle cx="${CCX}" cy="${CCY}" r="${CR}" fill="none" stroke="rgba(180,165,140,0.20)" stroke-width="8"/>
    <circle cx="${CCX}" cy="${CCY}" r="${CR}" fill="none"
      stroke="url(#${uid})" stroke-width="8" stroke-linecap="round"
      stroke-dasharray="${CCIRC.toFixed(2)}" stroke-dashoffset="${CCIRC.toFixed(2)}"
      transform="rotate(-90 ${CCX} ${CCY})"
      style="transition:stroke-dashoffset 0.9s cubic-bezier(0.4,0,0.2,1);"/>
    <text x="${CCX}" y="${CCY+5}"
      font-family="Barlow Condensed,sans-serif" font-size="13" font-weight="400"
      text-anchor="middle" fill="${color}">${txt}</text>
  </svg>`;
  requestAnimationFrame(()=>requestAnimationFrame(()=>{
    const f=el.querySelector('circle[transform]');
    if(f)f.style.strokeDashoffset=off.toFixed(2);
  }));
}

function st(r,m){if(!m)return'gray';const p=(r/m)*100;return p>=100?'green':p>=70?'yellow':'red';}
function pl(r,m){return m?Math.round((r/m)*100)+'%':'—';}
function setS(id,s){const c=document.getElementById(id);if(c&&!c.classList.contains('line-l1')&&!c.classList.contains('line-l2'))c.setAttribute('data-s',s);}
function setTip(id,t){const c=document.getElementById(id);if(c)c.textContent=t;}
function hkpi(bId,vId,tId,val,mr,fn){
  const b=document.getElementById(bId),v=document.getElementById(vId);
  if(!v||!b)return;
  v.innerHTML=fn(val);
  const barId=bId.replace('hk-','hb-');
  const bar=document.getElementById(barId);
  if(bar&&mr){
    const pct=Math.min((val/mr)*100,100);
    const s=pct>=100?'green':pct>=70?'yellow':'red';
    bar.className='hk-bar '+s;
    requestAnimationFrame(()=>requestAnimationFrame(()=>{bar.style.width=pct.toFixed(1)+'%';}));
    const pctEl=document.getElementById(barId.replace('hb-','hp-'));
    if(pctEl)pctEl.textContent=pct.toFixed(0)+'%';
  }
  b.className='hk neu';
}
function dias(de,ate){return Math.max(1,Math.round((new Date(ate)-new Date(de))/864e5)+1);}
function fmt(n){return n.toLocaleString('pt-BR');}
function fR(n){return'R$'+Math.round(n).toLocaleString('pt-BR');}
function fL(n){return Math.round(n)+'L';}

/* ── SPARKLINE ── */
function spark(id,vals,labs,fmtFn,highlightIdx=-1){
  const wrap=document.getElementById(id);if(!wrap)return;
  wrap.innerHTML='';
  const W=wrap.offsetWidth||220,H=wrap.offsetHeight||120;
  const pL=8,pR=12,pT=20,pB=20,uW=W-pL-pR,uH=H-pT-pB,n=vals.length;
  const valid=vals.filter(v=>v>0);if(!valid.length)return;
  const mn=Math.min(...valid)*0.88,mx=Math.max(...valid)*1.08,rng=mx-mn||1;
  const xs=vals.map((_,i)=>pL+(i/(n-1))*uW);
  const ys=vals.map(v=>pT+uH-((v-mn)/rng)*uH);
  const pts=xs.map((x,i)=>`${x.toFixed(1)},${ys[i].toFixed(1)}`).join(' ');
  const area=`M${xs[0].toFixed(1)},${ys[0].toFixed(1)} ${xs.map((x,i)=>`L${x.toFixed(1)},${ys[i].toFixed(1)}`).join(' ')} L${xs[n-1].toFixed(1)},${(H-pB).toFixed(1)} L${xs[0].toFixed(1)},${(H-pB).toFixed(1)} Z`;
  const dotColors=vals.map((v,i)=>i===0?'#9BA8B0':v>=vals[i-1]?'#1E7A42':'#B82418');
  const vLbls=vals.map((v,i)=>{
    const a=i===0?'start':i===n-1?'end':'middle';
    return `<text x="${xs[i].toFixed(1)}" y="${(ys[i]-7).toFixed(1)}" text-anchor="${a}"
      font-family="Barlow Condensed,sans-serif" font-size="11" font-weight="600"
      fill="${dotColors[i]}">${fmtFn(v)}</text>`;
  }).join('');
  const sLbls=labs.map((l,i)=>{
    const a=i===0?'start':i===n-1?'end':'middle';
    return `<text x="${xs[i].toFixed(1)}" y="${(H-4).toFixed(1)}" text-anchor="${a}"
      font-family="Barlow Condensed,sans-serif" font-size="10" font-weight="400"
      fill="#A89870">${l}</text>`;
  }).join('');
  const dots=vals.map((v,i)=>{
    const isHL=highlightIdx>=0&&i===highlightIdx;
    const r=isHL?7:4;
    const stroke=isHL?'#FFA62C':'white';
    const sw=isHL?2.5:2;
    const op=highlightIdx>=0&&!isHL?'0.35':'1';
    return `<circle class="sd"
      cx="${xs[i].toFixed(1)}" cy="${ys[i].toFixed(1)}" r="${r}" fill="${dotColors[i]}"
      stroke="${stroke}" stroke-width="${sw}" opacity="${op}"
      style="cursor:pointer;transition:r 0.15s;"
      data-v="${fmtFn(v)}" data-l="${labs[i]}"/>`;
  }).join('');
  const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
  svg.setAttribute('viewBox',`0 0 ${W} ${H}`);
  svg.setAttribute('width','100%');svg.setAttribute('height','100%');
  svg.style.display='block';
  svg.innerHTML=`<path d="${area}" fill="rgba(180,160,120,0.08)"/>
    <polyline points="${pts}" fill="none" stroke="#C4B89A" stroke-width="1.8"
      stroke-linejoin="round" stroke-linecap="round"/>
    ${vLbls}${dots}${sLbls}`;
  const oldTip=document.querySelector('.sp-tip[data-id="'+id+'"]');
  if(oldTip)oldTip.remove();
  const tip=document.createElement('div');
  tip.className='sp-tip';tip.dataset.id=id;
  wrap.style.position='relative';
  document.body.appendChild(tip);wrap.appendChild(svg);
  svg.querySelectorAll('.sd').forEach(d=>{
    d.addEventListener('mouseenter',(e)=>{
      d.setAttribute('r','6');
      tip.textContent=d.dataset.l+': '+d.dataset.v;
      tip.style.left=(e.clientX+12)+'px';tip.style.top=(e.clientY-32)+'px';tip.style.opacity='1';
    });
    d.addEventListener('mousemove',(e)=>{ tip.style.left=(e.clientX+12)+'px';tip.style.top=(e.clientY-32)+'px'; });
    d.addEventListener('mouseleave',()=>{d.setAttribute('r','4');tip.style.opacity='0';});
  });
}

/* ── VISÃO GERAL ── */
function go(){
  const mes   = parseInt(document.getElementById('f-mes')?.value) || (new Date().getMonth()+1);
  const sem   = parseInt(document.getElementById('f-sem')?.value) || 0;
  const weeks = getWeeksForFilter(mes, sem);

  const alc = sumSemanal('Alcance', weeks);
  const at  = sumSemanal('Engajamento / Atendimento', weeks);
  const orc = sumSemanal('Orçamentos', weeks);
  const ped = sumSemanal('Pedidos', weeks);
  const lit = sumSemanal('Litros vendidos', weeks);
  const fat = sumSemanal('Faturamento', weeks);

  const tAlc=alc.res,mrAlc=alc.meta;
  const tAt=at.res,mrAt=at.meta;
  const tOrc=orc.res,mrO=orc.meta;
  const tPed=ped.res,mrP=ped.meta;
  const tLit=lit.res,mrL=lit.meta;
  const tRec=fat.res,mrR=fat.meta;

  const tmP=tPed?tRec/tPed:0, tmL=tPed?tLit/tPed:0, rL=tLit?tRec/tLit:0;
  const nSem=weeks.length||1;
  const aL=tLit/nSem, aR=tRec/nSem;
  const sAlc=st(tAlc,mrAlc),sAt=st(tAt,mrAt);
  const sO=st(tOrc,mrO),sL=st(tLit,mrL),sR=st(tRec,mrR),sP=st(tPed,mrP);

  document.getElementById('v-alc').textContent=fmt(Math.round(tAlc));
  document.getElementById('v-at').textContent=fmt(Math.round(tAt));
  document.getElementById('v-orc').textContent=fmt(Math.round(tOrc));
  document.getElementById('v-ped').textContent=fmt(Math.round(tPed));
  document.getElementById('v-lit').innerHTML=fmt(Math.round(tLit))+'<span class="u"> L</span>';
  document.getElementById('v-rec').textContent='R$ '+fmt(Math.round(tRec));
  document.getElementById('v-tp').textContent='R$ '+fmt(Math.round(tmP));
  document.getElementById('v-tl').innerHTML=tmL.toFixed(1)+'<span class="u"> L</span>';
  document.getElementById('v-rl').textContent='R$ '+rL.toFixed(2).replace('.',',');

  const convAlcAt=tAlc?(tAt/tAlc*100).toFixed(1)+'% de conversão':'—';
  const convAtOrc=tAt?(tOrc/tAt*100).toFixed(1)+'% de conversão':'—';
  const convOrcPed=tOrc?(tPed/tOrc*100).toFixed(1)+'% de conversão':'—';
  document.getElementById('bz-alc').textContent=convAlcAt;
  document.getElementById('bz-at').textContent=convAtOrc;
  document.getElementById('bz-orc').textContent=convAtOrc;
  document.getElementById('bz-ped').textContent=convOrcPed;
  document.getElementById('bz-lit').textContent=fmt(Math.round(aL))+' L/semana em média';
  document.getElementById('bz-rec').textContent='R$ '+fmt(Math.round(aR))+'/semana em média';
  document.getElementById('bz-tp').textContent='~R$ '+fmt(Math.round(tmP))+' por evento';
  document.getElementById('bz-tl').textContent='~'+tmL.toFixed(1)+' L por evento';
  document.getElementById('bz-rl').textContent='~R$ '+rL.toFixed(2).replace('.',',')+' por litro vendido';

  // Evolução vs ano anterior
  function fmtEvo(res, ant) {
    if (!ant) return '—';
    const pct = ((res - ant) / ant * 100);
    const arrow = pct >= 0 ? '↑' : '↓';
    const color = pct >= 0 ? '#1E7A42' : '#B82418';
    const el = document.createElement ? null : null; // inline via textContent + style
    return { txt: arrow+' '+Math.abs(Math.round(pct))+'% vs 2025', color };
  }
  function setEvo(id, res, ant) {
    const el = document.getElementById(id+'-evo');
    if (!el) return;
    if (!ant) { el.textContent = '—'; el.style.color = '#9BA8B0'; return; }
    const pct = (res - ant) / ant * 100;
    el.textContent = (pct >= 0 ? '↑ ' : '↓ ') + Math.abs(Math.round(pct)) + '% vs 2025';
    el.style.color = pct >= 0 ? '#1E7A42' : '#B82418';
  }
  setEvo('bz-alc', tAlc, alc.anoAnt);
  setEvo('bz-at',  tAt,  at.anoAnt);
  setEvo('bz-orc', tOrc, orc.anoAnt);
  setEvo('bz-ped', tPed, ped.anoAnt);
  setEvo('bz-lit', tLit, lit.anoAnt);
  setEvo('bz-rec', tRec, fat.anoAnt);

  setS('c-orc',sO);setS('c-ped',sP);setS('c-lit',sL);setS('c-rec',sR);
  setTip('ct-alc',fmt(Math.round(tAt))+' atend de '+fmt(Math.round(tAlc))+' alcance');
  setTip('ct-at',fmt(Math.round(tOrc))+' orc de '+fmt(Math.round(tAt))+' atend');
  setTip('ct-orc','Meta: '+fmt(Math.round(mrO))+' · Real: '+fmt(Math.round(tOrc)));
  setTip('ct-ped','Meta: '+fmt(Math.round(mrP))+' · Real: '+fmt(Math.round(tPed)));
  setTip('ct-lit','Meta: '+fmt(Math.round(mrL))+'L · Real: '+fmt(Math.round(tLit))+'L');
  setTip('ct-rec','Meta: R$ '+fmt(Math.round(mrR))+' · Real: R$ '+fmt(Math.round(tRec)));

  circ('ci-alc',mrAlc?(tAlc/mrAlc*100):0,SC[sAlc],pl(Math.round(tAlc),Math.round(mrAlc)));
  circ('ci-at',mrAt?(tAt/mrAt*100):0,SC[sAt],pl(Math.round(tAt),Math.round(mrAt)));
  circ('ci-orc',mrO?(tOrc/mrO*100):0,SC[sO],pl(tOrc,mrO));
  circ('ci-ped',mrP?(tPed/mrP*100):0,SC[sP],pl(tPed,mrP));
  circ('ci-lit',mrL?(tLit/mrL*100):0,SC[sL],pl(Math.round(tLit),Math.round(mrL)));
  circ('ci-rec',mrR?(tRec/mrR*100):0,SC[sR],pl(Math.round(tRec),Math.round(mrR)));

  // Sparklines — sempre mostra as 4 semanas do mês selecionado para contexto
  SEM=['S1','S2','S3','S4'];
  const spP=[],spL=[],spR=[];
  [1,2,3,4].forEach(s=>{
    const fv=SEMANAL_RAW['Faturamento']?.[mes]?.[s]?.res||0;
    const pv=SEMANAL_RAW['Pedidos']?.[mes]?.[s]?.res||0;
    const lv=SEMANAL_RAW['Litros vendidos']?.[mes]?.[s]?.res||0;
    spP.push(pv?fv/pv:0); spL.push(pv?lv/pv:0); spR.push(lv?fv/lv:0);
  });
  const hlIdx = sem > 0 ? sem - 1 : -1;
  requestAnimationFrame(()=>{ spark('sp-tp',spP,SEM,fR,hlIdx); spark('sp-tl',spL,SEM,fL,hlIdx); spark('sp-rl',spR,SEM,v=>'R$'+v.toFixed(2).replace('.',','),hlIdx); });

  hkpi('hk-lit','hv-lit','ht-lit',tLit,mrL,v=>fmt(Math.round(v))+'<span class="u"> L</span>');
  hkpi('hk-rec','hv-rec','ht-rec',tRec,mrR,v=>'R$'+Math.round(v/1000)+'k');
  const cvPct=tOrc?(tPed/tOrc*100):0;
  document.getElementById('hv-cv').textContent=tOrc?cvPct.toFixed(1)+'%':'—';
  const cvBar=document.getElementById('hb-cv');
  if(cvBar){
    const s=cvPct>=50?'green':cvPct>=30?'yellow':'red';
    cvBar.className='hk-bar '+s;
    requestAnimationFrame(()=>requestAnimationFrame(()=>{cvBar.style.width=Math.min(cvPct,100).toFixed(1)+'%';}));
    const cvPctEl=document.getElementById('hp-cv');
    if(cvPctEl)cvPctEl.textContent=Math.min(cvPct,100).toFixed(0)+'%';
  }
  const MESES=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const resumoEl=document.getElementById('h-resumo-lbl');
  if(resumoEl)resumoEl.textContent=sem?MESES[mes-1]+' · S'+sem:'Resumo '+MESES[mes-1];

  // Atualiza abas ativas ao filtrar
  ezRendered=false;
  if(document.getElementById('tab-ez')?.classList.contains('active'))renderEZ();
  if(document.getElementById('tab-metas')?.classList.contains('active'))renderMetas();
}

/* ── ABA EZ ── */
let ezRendered=false;
function renderEZ(){
  if(ezRendered)return;
  ezRendered=true;
  const mes  = parseInt(document.getElementById('f-mes')?.value) || (new Date().getMonth()+1);

  const sem  = parseInt(document.getElementById('f-sem')?.value) || 0;
  const resp = document.getElementById('f-resp')?.value || '';
  const {de, ate} = getDateRangeForFilter(mes, sem);

  const data=EZ_TICKETS.filter(d=>{
    if(de&&d.DataStr<de)return false;
    if(ate&&d.DataStr>ate)return false;
    if(resp&&d.Agente!==resp)return false;
    return true;
  });

  const total=data.length;
  const tpiMed=data.reduce((s,d)=>s+(d.TPI_min||0),0)/Math.max(total,1);
  const tmaMed=data.reduce((s,d)=>s+(d.TMA_min||0),0)/Math.max(total,1);
  const ativo=data.filter(d=>d.Ativo==='ATIVO').length;
  const recep=data.filter(d=>d.Ativo==='RECEPTIVO').length;

  // Classificações — 100% de EZ_TICKETS
  const classCount={};
  data.forEach(d=>{ const c=d.Classificacao||'Sem Classificação'; classCount[c]=(classCount[c]||0)+1; });
  const classSort=Object.entries(classCount).sort((a,b)=>b[1]-a[1]).slice(0,7)
    .map(([label,count])=>({label,count,pct:total?count/total:0}));

  // Performance por agente — 100% de EZ_TICKETS
  const agentesUniq=[...new Set(EZ_TICKETS.map(d=>d.Agente))].filter(Boolean).sort();
  const perf=agentesUniq.map(a=>{
    const ag=data.filter(d=>d.Agente===a);
    if(!ag.length)return null;
    const fin=ag.filter(d=>d.Status==='Finalizado').length;
    const tpi=ag.reduce((s,d)=>s+(d.TPI_min||0),0)/Math.max(ag.length,1);
    const tma=ag.reduce((s,d)=>s+(d.TMA_min||0),0)/Math.max(ag.length,1);
    const cc={};ag.forEach(d=>{const c=d.Classificacao||'Sem class.';cc[c]=(cc[c]||0)+1;});
    const topClass=Object.entries(cc).sort((a,b)=>b[1]-a[1])[0]?.[0]||'—';
    return{nome:a,tickets:ag.length,fin,tpiMin:tpi,tmaMin:tma,topClass};
  }).filter(Boolean);

  function fmtMin(min){
    min=Math.round(min);
    if(min<60)return min+'min';
    return Math.floor(min/60)+'h '+String(min%60).padStart(2,'0')+'min';
  }

  const classColors=['#3D6490','#2E6644','#6B4E10','#8B3A8B','#C8941A','#9BA8B0','#B85C38'];
  const totalLabel=total.toLocaleString('pt-BR');
  const bezLabel=ativo+' ativos · '+recep+' receptivos';
  const tpiLabel=fmtMin(tpiMed);
  const tmaLabel=fmtMin(tmaMed);

  const html=`
  <div class="row">
    <div class="card line-l1" data-s="none">
      <div class="card-ab">
        <div class="c-header"><div class="c-title pill-l1">Total de Tickets</div><div class="c-sub">Protocolos com atendimento humano</div></div>
        <div class="c-center"><div class="c-val-block">
          <div class="ez-kpi-val">${totalLabel}</div>
          <span class="bezel neu">${bezLabel}</span>
        </div></div>
      </div>
    </div>
    <div class="card line-l1" data-s="none">
      <div class="card-ab">
        <div class="c-header"><div class="c-title pill-l1">TPI Médio</div><div class="c-sub">Tempo para primeira interação · equipe</div></div>
        <div class="c-center"><div class="c-val-block">
          <div class="ez-kpi-val" style="font-size:38px;">${tpiLabel}</div>
          <span class="bezel neu">tempo de resposta inicial</span>
        </div></div>
      </div>
    </div>
    <div class="card line-l1" data-s="none">
      <div class="card-ab">
        <div class="c-header"><div class="c-title pill-l1">TMA Médio</div><div class="c-sub">Tempo médio de atendimento · equipe</div></div>
        <div class="c-center"><div class="c-val-block">
          <div class="ez-kpi-val" style="font-size:38px;">${tmaLabel}</div>
          <span class="bezel neu">duração média por ticket</span>
        </div></div>
      </div>
    </div>
  </div>

  <div class="row" style="grid-template-columns:1fr 1fr;">
    <div class="card line-l2" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l2">Classificação dos Tickets</div><div class="c-sub">Distribuição por tipo de resultado</div></div>
        <div style="margin-top:10px;width:100%;">
          ${classSort.map(({label,count,pct},i)=>`
            <div class="ez-bar-row" style="margin-bottom:8px;font-size:13px;">
              <div class="ez-bar-label" style="font-size:13px;">${label}</div>
              <div class="ez-bar-track" style="height:8px;"><div class="ez-bar-fill" style="width:${(pct*100).toFixed(1)}%;background:${classColors[i%classColors.length]};height:8px;border-radius:4px;"></div></div>
              <div class="ez-bar-pct" style="font-size:13px;">${Math.round(pct*100)}%</div>
            </div>`).join('')}
        </div>
      </div>
    </div>
    <div class="card line-l2" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l2">Picos de Demanda</div><div class="c-sub">Mapa de calor · dia da semana × hora do dia</div></div>
        <div id="ez-heatmap" style="margin-top:8px;overflow:hidden;"></div>
      </div>
    </div>
  </div>

  <div class="row" style="grid-template-columns:1fr;">
    <div class="card line-l3" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l3">Performance por Agente</div><div class="c-sub">Consolidado do período filtrado</div></div>
        <div style="margin-top:12px;overflow-x:auto;">
          <table class="ez-table">
            <thead><tr>
              <th>Agente</th><th>Tickets</th><th>Finalizados</th><th>% Finalizado</th>
              <th>TPI Médio</th><th>TMA Médio</th><th>Classificação Mais Frequente</th>
            </tr></thead>
            <tbody>
              ${perf.map(p=>`<tr>
                <td class="agent">${p.nome}</td>
                <td class="num">${p.tickets}</td>
                <td class="num">${p.fin}</td>
                <td>${p.tickets?Math.round((p.fin/p.tickets)*100)+'%':'—'}</td>
                <td>${p.tpiMed||fmtMin(p.tpiMin||0)}</td>
                <td>${p.tmaMed||fmtMin(p.tmaMin||0)}</td>
                <td><span class="ez-badge">${p.topClass}</span></td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>`;

  document.getElementById('ez-main').innerHTML=html;
  buildHeatmapFromTickets(data);
}

/* ══ HEATMAP ══ */
function buildHeatmapFromTickets(data){
  const DAYS=['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  const matrix=Array.from({length:7},()=>new Array(24).fill(0));
  data.forEach(d=>{
    const hora=d.Hora;
    if(hora<0||hora>23)return;
    const dt=new Date(d.DataStr+'T12:00:00');
    if(isNaN(dt))return;
    matrix[dt.getDay()][hora]++;
  });
  renderHeatmapMatrix(DAYS,matrix);
}

function renderHeatmapMatrix(DAYS,matrix){
  const el=document.getElementById('ez-heatmap');
  if(!el)return;
  const maxVal=Math.max(...matrix.flat(),1);
  function getColor(val){
    if(val===0)return'rgba(200,185,160,0.10)';
    const t=val/maxVal;
    if(t<=0.33){const p=t/0.33;return`rgb(${Math.round(241+(255-241)*p)},${Math.round(227+(166-227)*p)},${Math.round(206+(44-206)*p)})`;}
    else{const p=(t-0.33)/0.67;return`rgb(${Math.round(255+(201-255)*p)},${Math.round(166+(43-166)*p)},${Math.round(44+(30-44)*p)})`;}
  }
  function textColor(val){return(val/maxVal)>0.45?'#FFF8F0':'#8B7040';}
  const cellW=22,cellH=28,leftPad=32,topPad=22,bottomPad=22;
  const svgW=leftPad+24*cellW+4,svgH=topPad+7*cellH+bottomPad;
  let inner='';
  for(let h=0;h<24;h++){
    if(h%3===0)inner+=`<text x="${leftPad+h*cellW+cellW/2}" y="${topPad-5}"
      text-anchor="middle" font-family="Barlow Condensed,sans-serif" font-size="9" fill="#A89870">${String(h).padStart(2,'0')}h</text>`;
  }
  for(let d=0;d<7;d++){
    const y=topPad+d*cellH;
    inner+=`<text x="${leftPad-4}" y="${y+cellH/2+4}" text-anchor="end"
      font-family="Barlow Condensed,sans-serif" font-size="9" font-weight="600" fill="#A89870">${DAYS[d]}</text>`;
    for(let h=0;h<24;h++){
      const x=leftPad+h*cellW,val=matrix[d][h];
      inner+=`<rect x="${x+1}" y="${y+1}" width="${cellW-2}" height="${cellH-2}" rx="2" fill="${getColor(val)}"
        data-val="${val}" data-day="${DAYS[d]}" data-hour="${String(h).padStart(2,'0')}h"/>`;
      if(val>0)inner+=`<text x="${x+cellW/2}" y="${y+cellH/2+4}" text-anchor="middle"
        font-family="Barlow Condensed,sans-serif" font-size="10" font-weight="600"
        fill="${textColor(val)}" pointer-events="none">${val}</text>`;
    }
  }
  inner+=`<text x="${leftPad-4}" y="${topPad+7*cellH+bottomPad-5}" text-anchor="end"
    font-family="Barlow Condensed,sans-serif" font-size="9" fill="rgba(168,152,112,0.6)">total</text>`;
  for(let h=0;h<24;h++){
    const tot=matrix.reduce((s,row)=>s+row[h],0);
    if(tot>0)inner+=`<text x="${leftPad+h*cellW+cellW/2}" y="${topPad+7*cellH+bottomPad-5}" text-anchor="middle"
      font-family="Barlow Condensed,sans-serif" font-size="9" fill="rgba(168,152,112,0.7)">${tot}</text>`;
  }
  el.innerHTML=`<svg viewBox="0 0 ${svgW} ${svgH}" width="100%" preserveAspectRatio="xMidYMid meet"
    xmlns="http://www.w3.org/2000/svg" style="display:block;">${inner}</svg>`;
  const oldTip=document.querySelector('.sp-tip[data-id="heatmap"]');
  if(oldTip)oldTip.remove();
  const tip=document.createElement('div');
  tip.className='sp-tip';tip.dataset.id='heatmap';
  document.body.appendChild(tip);
  el.querySelectorAll('rect').forEach(r=>{
    r.style.cursor='default';
    r.addEventListener('mouseenter',e=>{
      const val=r.dataset.val;if(val==='0')return;
      tip.textContent=`${r.dataset.day} ${r.dataset.hour}: ${val} ticket${val!='1'?'s':''}`;
      tip.style.left=(e.clientX+12)+'px';tip.style.top=(e.clientY-32)+'px';tip.style.opacity='1';
    });
    r.addEventListener('mousemove',e=>{tip.style.left=(e.clientX+12)+'px';tip.style.top=(e.clientY-32)+'px';});
    r.addEventListener('mouseleave',()=>{tip.style.opacity='0';});
  });
}

/* ── CONTROLES ── */
function reset(){
  const today=new Date();
  const m=document.getElementById('f-mes'),s=document.getElementById('f-sem'),r=document.getElementById('f-resp');
  if(m)m.value=String(today.getMonth()+1);
  if(s)s.value='0';
  if(r)r.value='';
  document.querySelectorAll('.btn-sh').forEach(b=>b.classList.remove('active'));
  go();
}

function setShortcut(type){
  const today=new Date();
  let mes=today.getMonth()+1, sem=0;
  if(type==='mes-passado'){ mes=today.getMonth()||12; }
  const m=document.getElementById('f-mes'),s=document.getElementById('f-sem');
  if(m)m.value=String(mes);
  if(s)s.value=String(sem);
  document.querySelectorAll('.btn-sh').forEach(b=>b.classList.remove('active'));
  event.target.classList.add('active');
  go();
}

function setTab(el){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  const tabName=el.textContent.trim();
  document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
  // Mostrar filtro Agente só na aba EZ
  const respGroup=document.getElementById('f-resp-group');
  if(respGroup)respGroup.style.display=tabName==='Atendimento EZ'?'flex':'none';
  if(tabName==='Visão Geral')document.getElementById('tab-visao').classList.add('active');
  else if(tabName==='Atendimento EZ'){document.getElementById('tab-ez').classList.add('active');renderEZ();}
  else if(tabName==='Metas'){document.getElementById('tab-metas').classList.add('active');renderMetas();}
}

/* ── FILTROS PADRÃO — mês atual, todo o mês ── */
(function(){
  const today=new Date();
  const m=document.getElementById('f-mes'),s=document.getElementById('f-sem');
  if(m)m.value=String(today.getMonth()+1);
  if(s)s.value='0';
})();

const _upd=document.getElementById('upd-date');
if(_upd){const _d=new Date();_upd.textContent=String(_d.getDate()).padStart(2,'0')+'/'+String(_d.getMonth()+1).padStart(2,'0')+'/'+_d.getFullYear();}

loadData();
