/* ══════════════════════════════════════════
   CONFIGURAÇÃO — URLs do Google Sheets
   Publicar cada aba: Arquivo → Publicar na web → CSV
══════════════════════════════════════════ */
const SHEETS_URL_RD  = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vS69WfTMMD5rXrETJATYKq-jBHIE414MjLwEgRdgB_o-ncM5bmWIF_5URzpXeWwY4j49IyflTSrDifw/pub?gid=390115291&single=true&output=csv'; // Cole aqui a URL da aba RD
const SHEETS_URL_EZ  = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vS69WfTMMD5rXrETJATYKq-jBHIE414MjLwEgRdgB_o-ncM5bmWIF_5URzpXeWwY4j49IyflTSrDifw/pub?gid=880773931&single=true&output=csv'; // Cole aqui a URL da aba EZ
const SHEETS_URL_ALC  = ''; // Cole aqui a URL da aba Alcance
const SHEETS_URL_METAS = ''; // Cole aqui a URL da aba Metas

/* Dados em memória — preenchidos pelo fetch */
let RD  = [];
let EZ_PROTO = [];

/* ── PARSER CSV ── */
function parseCSV(text, sep=',') {
  const lines = text.trim().split(/\r?\n/);
  // Detectar e pular linha "sep=..."
  const start = lines[0].startsWith('sep=') ? 1 : 0;
  const headers = lines[start].split(sep).map(h => h.replace(/^"|"$/g,'').trim());
  return lines.slice(start+1).filter(l=>l.trim()).map(line => {
    // Parser simples respeitando campos com vírgula entre aspas
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

/* ── PARSER EZ BRUTO → EZ_PROTO ── */
function processEZ(rows) {
  const agentesValidos = ['Taciana', 'Lídia', 'Raiza'];
  const RESP_MAP = {
    'Taciana Daniela da silva campos': 'Taciana',
    'Lídia Felix': 'Lídia',
    'raiza.huguenin': 'Raiza'
  };

  function hmsToSec(s) {
    if (!s || s === '-' || s === '0') return 0;
    const p = String(s).trim().split(':');
    if (p.length === 3) return +p[0]*3600 + +p[1]*60 + +p[2];
    if (p.length === 2) return +p[0]*60 + +p[1];
    return 0;
  }

  // Filtrar humano e converter datas
  const humanos = rows
    .filter(r => (r['Tipo']||'').trim() === 'Humano')
    .map(r => ({
      ...r,
      _date: new Date(r['Criado em'].replace(/(\d{2})\/(\d{2})\/(\d{4})/, '$3-$2-$1')),
      _tpi: hmsToSec(r['Tempo para Primeira Interação']),
      _tma: hmsToSec(r['Tempo de Atendimento']),
      _agente: RESP_MAP[r['Nome do Agente']] || r['Nome do Agente']
    }))
    .sort((a,b) => a._date - b._date);

  // Agrupar por protocolo
  const protos = {};
  humanos.forEach(r => {
    const p = r['Protocolo'];
    if (!protos[p]) protos[p] = [];
    protos[p].push(r);
  });

  return Object.entries(protos).map(([proto, linhas]) => {
    const primeira = linhas[0];
    const ultima = [...linhas].reverse().find(r =>
      ['Vendas','Sua Casa Nosso Bar','-'].includes(r['Nome do Departamento'])
    ) || linhas[linhas.length-1];

    const agente = RESP_MAP[ultima['Nome do Agente']] || ultima['Nome do Agente'];
    if (!agentesValidos.includes(agente)) return null;

    const tma = linhas.reduce((s,r) => s + r._tma, 0);
    const d = primeira._date;
    const dataStr = d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` : '';

    return {
      Protocolo: proto,
      TPI_sec: primeira._tpi,
      TMA_sec: tma,
      DataStr: dataStr,
      Hour: d ? d.getHours() : -1,
      Ativo: primeira['Ativo'] || '',
      Agente: agente,
      Classificacao: ultima['Classificações'] || '',
      Status: ultima['Status'] || '',
      'Avaliação CSAT': ultima['Avaliação CSAT'] || ''
    };
  }).filter(Boolean);
}

/* ── PARSER RD BRUTO ── */
function processRD(rows) {
  const RESP_MAP = {
    'Taciana Daniela da silva campos': 'Taciana',
    'Lídia Felix': 'Lídia',
    'raiza.huguenin': 'Raiza',
    'Mirian Costa': 'Mirian',
    'Marciano luis da silva': 'Marciano',
    'Gabriel Costa': 'Gabriel'
  };
  return rows
    .filter(r => (r['Funil de vendas']||'').trim() === '*Vendas Varejo B2C')
    .map(r => {
      const d = r['Data de criação'];
      let dataStr = '';
      if (d) {
        const [dia,mes,ano] = d.split('/');
        dataStr = `${ano}-${mes?.padStart(2,'0')}-${dia?.padStart(2,'0')}`;
      }
      const sem = dataStr ? (() => {
        const dt = new Date(dataStr);
        const jan1 = new Date(dt.getFullYear(),0,1);
        return Math.ceil(((dt - jan1)/86400000 + jan1.getDay()+1)/7);
      })() : 0;
      return {
        Resp: RESP_MAP[r['Responsável']] || 'Outros',
        Estado: r['Estado'] || '',
        Etapa: r['Etapa'] || '',
        'Valor Único': parseFloat(r['Valor Único'])||0,
        Litragem_num: parseFloat(r['Litragem'])||0,
        DataStr: dataStr,
        Semana: sem
      };
    });
}


/* ── METAS DINÂMICAS ── */
function getMetasMes(de, ate) {
  if (!METAS_DATA.length) return METAS_DEFAULT;
  // Usar mês da data inicial do filtro
  const d = new Date(de);
  const mes = d.getMonth() + 1;
  const ano = d.getFullYear();
  const row = METAS_DATA.find(r => +r['Mês'] === mes && +r['Ano'] === ano);
  if (!row) return METAS_DEFAULT;
  return {
    orc: parseFloat(row['Orçamentos']) || 0,
    ped: parseFloat(row['Pedidos'])    || 0,
    lit: parseFloat(row['Litros'])     || 0,
    rec: parseFloat(row['Receita'])    || 0
  };
}

/* ── PARSER ALCANCE ── */
function processALC(rows) {
  // Soma total de leads no período
  return rows.reduce((s, r) => s + (parseFloat(r['Leads']) || 0), 0);
}

/* ── PARSER METAS ── */
function processMetas(rows) {
  return rows.filter(r => r['Mês'] && r['Ano']);
}


/* ── CALCULAR PART E REF DINAMICAMENTE ── */
function calcPart() {
  const vendidas = RD.filter(d => d.Estado === 'Vendida');
  const total = vendidas.length || 1;
  const agentes = ['Taciana','Lídia','Raiza'];
  agentes.forEach(a => {
    PART[a] = vendidas.filter(d => d.Resp === a).length / total;
  });
  // REF = média geral do período completo
  const tRec = vendidas.reduce((s,d) => s + (d['Valor Único']||0), 0);
  const tLit = vendidas.reduce((s,d) => s + (d.Litragem_num||0), 0);
  const tPed = vendidas.length || 1;
  REF.tp = tRec / tPed;
  REF.tl = tLit / tPed;
  REF.rl = tLit ? tRec / tLit : 0;
}

/* ── ESTADO DE LOADING ── */
function setLoading(on) {
  const msg = document.getElementById('loading-msg');
  if (msg) msg.style.display = on ? 'flex' : 'none';
}

/* ── CARREGAMENTO PRINCIPAL ── */
async function loadData() {
  setLoading(true);
  try {
    const promises = [];
    if (SHEETS_URL_RD)    promises.push(fetch(SHEETS_URL_RD).then(r=>r.text()).then(t=>{ RD = processRD(parseCSV(t,',')); calcPart(); }));
    else RD = [];
    if (SHEETS_URL_EZ)    promises.push(fetch(SHEETS_URL_EZ).then(r=>r.text()).then(t=>{ EZ_PROTO = processEZ(parseCSV(t,',')); }));
    else EZ_PROTO = [];
    if (SHEETS_URL_ALC)   promises.push(fetch(SHEETS_URL_ALC).then(r=>r.text()).then(t=>{ ALC = processALC(parseCSV(t,',')); }));
    if (SHEETS_URL_METAS) promises.push(fetch(SHEETS_URL_METAS).then(r=>r.text()).then(t=>{ METAS_DATA = processMetas(parseCSV(t,',')); }));
    await Promise.all(promises);
  } catch(e) {
    console.warn('Erro ao carregar dados:', e);
  }
  console.log('RD carregado:', RD.length, 'registros');
  console.log('EZ_PROTO carregado:', EZ_PROTO.length, 'registros');
  console.log('METAS_DATA carregado:', METAS_DATA.length, 'registros');
  console.log('ALC carregado:', ALC, 'leads');
  setLoading(false);
  go();
}

let PART={}; // calculado dinamicamente do RD;
let METAS_DATA=[]; // [{Mes,Ano,Orçamentos,Pedidos,Litros,Receita}] via Sheets
const METAS_DEFAULT={orc:0,lit:0,rec:0,ped:0};
let REF={tp:0,tl:0,rl:0}; // calculado do RD ao carregar
let ALC=0; // preenchido via Sheets
const ORC_ETAPAS=['Proposta Comercial','Cadastro Pedido','Logística / Entrega','Pós Venda e Fidelização'];
let SEM=[]; // gerado dinamicamente em go()
const SC={green:'#1E7A42',yellow:'#966A00',red:'#B82418',gray:'#9BA8B0'};

/* CÍRCULO — stroke 7px */
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
    <circle cx="${CCX}" cy="${CCY}" r="${CR+5}" fill="white" filter="url(#${uid}sh)"/>
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
  // Barra de progresso
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
function mpr(mt,de,ate,resp){
  let m=mt;if(resp&&PART[resp])m=mt*PART[resp];
  const h=new Date(),i=new Date(de),f=new Date(ate);
  const dp=Math.max(1,Math.round((f-i)/864e5)+1);
  const fe=f<h?f:h,pp=Math.max(1,Math.round((fe-i)/864e5)+1);
  return m*(pp/dp);
}
function dias(de,ate){return Math.max(1,Math.round((new Date(ate)-new Date(de))/864e5)+1);}
function fmt(n){return n.toLocaleString('pt-BR');}
function fR(n){return'R$'+Math.round(n).toLocaleString('pt-BR');}
function fL(n){return Math.round(n)+'L';}

function setShortcut(type){
  const today = new Date();
  const pad = n => String(n).padStart(2,'0');
  const fmt = d => d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
  let de, ate;
  if(type==='hoje'){
    de = ate = fmt(today);
  } else if(type==='7d'){
    const s = new Date(today); s.setDate(today.getDate()-6);
    de = fmt(s); ate = fmt(today);
  } else if(type==='15d'){
    const s = new Date(today); s.setDate(today.getDate()-14);
    de = fmt(s); ate = fmt(today);
  } else if(type==='mes-atual'){
    de = fmt(new Date(today.getFullYear(), today.getMonth(), 1));
    ate = fmt(new Date(today.getFullYear(), today.getMonth()+1, 0));
  } else if(type==='mes-passado'){
    de = fmt(new Date(today.getFullYear(), today.getMonth()-1, 1));
    ate = fmt(new Date(today.getFullYear(), today.getMonth(), 0));
  }
  document.getElementById('f-de').value = de;
  document.getElementById('f-ate').value = ate;
  // highlight ativo
  document.querySelectorAll('.btn-sh').forEach(b=>b.classList.remove('active'));
  event.target.classList.add('active');
  go();
}
function setTab(el){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  const tabName = el.textContent.trim();
  document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
  if(tabName==='Visão Geral') document.getElementById('tab-visao').classList.add('active');
  else if(tabName==='Atendimento EZ'){document.getElementById('tab-ez').classList.add('active');renderEZ();}
  else if(tabName==='Metas') document.getElementById('tab-metas').classList.add('active');
}
/* ── SPARKLINE ── */
function spark(id,vals,labs,fmtFn){
  const wrap=document.getElementById(id);if(!wrap)return;
  wrap.innerHTML='';
  const W=wrap.offsetWidth||220;
  const H=wrap.offsetHeight||120;
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
      fill="#A89870">${l.split('·')[0].trim()}</text>`;
  }).join('');

  const dots=vals.map((v,i)=>`<circle class="sd"
    cx="${xs[i].toFixed(1)}" cy="${ys[i].toFixed(1)}"
    r="4" fill="${dotColors[i]}" stroke="white" stroke-width="2"
    style="cursor:pointer;transition:r 0.15s;"
    data-v="${fmtFn(v)}" data-l="${labs[i].split('·')[0].trim()}"/>`).join('');

  const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
  svg.setAttribute('viewBox',`0 0 ${W} ${H}`);
  svg.setAttribute('width','100%');svg.setAttribute('height','100%');
  svg.style.display='block';
  svg.innerHTML=`
    <path d="${area}" fill="rgba(180,160,120,0.08)"/>
    <polyline points="${pts}" fill="none" stroke="#C4B89A" stroke-width="1.8"
      stroke-linejoin="round" stroke-linecap="round"/>
    ${vLbls}${dots}${sLbls}`;

  // Remover tip anterior se existir
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
      tip.style.left=(e.clientX+12)+'px';
      tip.style.top=(e.clientY-32)+'px';
      tip.style.opacity='1';
    });
    d.addEventListener('mousemove',(e)=>{
      tip.style.left=(e.clientX+12)+'px';
      tip.style.top=(e.clientY-32)+'px';
    });
    d.addEventListener('mouseleave',()=>{d.setAttribute('r','4');tip.style.opacity='0';});
  });
}


/* ══ RENDER ABA EZ ══ */
let ezRendered=false;
function renderEZ(){
  if(ezRendered)return;
  ezRendered=true;

  const de=document.getElementById('f-de').value||'2026-02-01';
  const ate=document.getElementById('f-ate').value||'2026-03-23';
  const resp=document.getElementById('f-resp').value||'';
  
  // Filtrar por data e agente
  let data=EZ_PROTO.filter(d=>{
    if(de&&d.DataStr<de)return false;
    if(ate&&d.DataStr>ate)return false;
    if(resp&&d.Agente!==resp)return false;
    return true;
  });

  const total=data.length;
  const tpiMed=data.reduce((s,d)=>s+(d.TPI_sec||0),0)/Math.max(total,1);
  const tmaMed=data.reduce((s,d)=>s+(d.TMA_sec||0),0)/Math.max(total,1);
  const ativo=data.filter(d=>d.Ativo==='ATIVO').length;
  const recep=data.filter(d=>d.Ativo==='RECEPTIVO').length;

  // Classificações
  const classCount={};
  data.forEach(d=>{
    const c=(d.Classificacao||'Sem classificação').split(',')[0].trim();
    classCount[c]=(classCount[c]||0)+1;
  });
  const classSort=Object.entries(classCount).sort((a,b)=>b[1]-a[1]).slice(0,6);

  // CSAT
  const csatData=data.filter(d=>d['Avaliação CSAT']&&d['Avaliação CSAT']!=='-'&&d['Avaliação CSAT']);
  const csatTotal=csatData.length;

  // Performance por agente
  const agentes=['Taciana','Lídia','Raiza'];
  const perf=agentes.map(a=>{
    const ag=data.filter(d=>d.Agente===a);
    const fin=ag.filter(d=>d.Status==='Finalizado').length;
    const tpi=ag.reduce((s,d)=>s+(d.TPI_sec||0),0)/Math.max(ag.length,1);
    const tma=ag.reduce((s,d)=>s+(d.TMA_sec||0),0)/Math.max(ag.length,1);
    const cc={};ag.forEach(d=>{const c=(d.Classificacao||'Sem class.').split(',')[0].trim();cc[c]=(cc[c]||0)+1;});
    const topClass=Object.entries(cc).sort((a,b)=>b[1]-a[1])[0]?.[0]||'—';
    return {nome:a,tickets:ag.length,fin,tpi,tma,topClass};
  });

  function fmtMin(sec){
    const m=Math.round(sec/60);
    if(m<60)return m+'min';
    return Math.floor(m/60)+'h '+String(m%60).padStart(2,'0')+'min';
  }

  // Cores por classificação
  const classColors=['#3D6490','#2E6644','#6B4E10','#8B3A8B','#C8941A','#9BA8B0'];



  // Montar HTML
  const html=`
  <!-- L1: Total · TPI · TMA -->
  <div class="row">
    <div class="card line-l1" data-s="none">
      <div class="card-ab">
        <div class="c-header"><div class="c-title pill-l1">Total de Tickets</div><div class="c-sub">Protocolos com atendimento humano</div></div>
        <div class="c-center">
          <div class="c-val-block">
            <div class="ez-kpi-val">${total.toLocaleString('pt-BR')}</div>
            <span class="bezel neu">${ativo} ativos · ${recep} receptivos</span>
          </div>
        </div>
      </div>
    </div>
    <div class="card line-l1" data-s="none">
      <div class="card-ab">
        <div class="c-header"><div class="c-title pill-l1">TPI Médio</div><div class="c-sub">Tempo para primeira interação · equipe</div></div>
        <div class="c-center">
          <div class="c-val-block">
            <div class="ez-kpi-val" style="font-size:38px;">${fmtMin(tpiMed)}</div>
            <span class="bezel neu">tempo de resposta inicial</span>
          </div>
        </div>
      </div>
    </div>
    <div class="card line-l1" data-s="none">
      <div class="card-ab">
        <div class="c-header"><div class="c-title pill-l1">TMA Médio</div><div class="c-sub">Tempo médio de atendimento · equipe</div></div>
        <div class="c-center">
          <div class="c-val-block">
            <div class="ez-kpi-val" style="font-size:38px;">${fmtMin(tmaMed)}</div>
            <span class="bezel neu">duração média por ticket</span>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- L2: Classificação dos Tickets — largura total -->
  <div class="row" style="grid-template-columns:1fr;">
    <div class="card line-l2" data-s="none" style="height:auto;min-height:var(--h-ab);">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l2">Classificação dos Tickets</div><div class="c-sub">Distribuição por tipo de resultado</div></div>
        <div style="margin-top:10px;">
          ${classSort.map(([label,count],i)=>`
            <div class="ez-bar-row">
              <div class="ez-bar-label">${label}</div>
              <div class="ez-bar-track"><div class="ez-bar-fill" style="width:${(count/total*100).toFixed(1)}%;background:${classColors[i]};"></div></div>
              <div class="ez-bar-pct">${(count/total*100).toFixed(0)}%</div>
            </div>`).join('')}
        </div>
      </div>
    </div>
  </div>

  <!-- L3: Picos de Demanda — largura total, altura generosa -->
  <div class="row" style="grid-template-columns:1fr;">
    <div class="card line-l3" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l3">Picos de Demanda</div><div class="c-sub">Mapa de calor · dia da semana × hora do dia</div></div>
        <div id="ez-heatmap" style="margin-top:10px;"></div>
      </div>
    </div>
  </div>

  <!-- L4: CSAT — largura total -->
  <div class="row" style="grid-template-columns:1fr;">
    <div class="card line-l4" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l4">CSAT</div><div class="c-sub">Pesquisa de satisfação</div></div>
        <div class="csat-ph" style="margin-top:8px;"><span>— —</span><small>${csatTotal>0?csatTotal+' avaliações':'Disponível ao fim do mês'}</small></div>
      </div>
    </div>
  </div>

  <!-- L5: Performance por Agente — largura total -->
  <div class="row" style="grid-template-columns:1fr;">
    <div class="card line-l3" data-s="none" style="height:auto;">
      <div class="card-ab" style="height:auto;padding-bottom:16px;">
        <div class="c-header"><div class="c-title pill-l3">Performance por Agente</div><div class="c-sub">Consolidado do período filtrado</div></div>
        <div style="margin-top:12px;overflow-x:auto;">
          <table class="ez-table">
            <thead>
              <tr>
                <th>Agente</th>
                <th>Tickets</th>
                <th>Finalizados</th>
                <th>% Finalizado</th>
                <th>TPI Médio</th>
                <th>TMA Médio</th>
                <th>Classificação Mais Frequente</th>
              </tr>
            </thead>
            <tbody>
              ${perf.map(p=>`<tr>
                <td class="agent">${p.nome}</td>
                <td class="num">${p.tickets}</td>
                <td class="num">${p.fin}</td>
                <td>${p.tickets?Math.round(p.fin/p.tickets*100)+'%':'—'}</td>
                <td>${fmtMin(p.tpi)}</td>
                <td>${fmtMin(p.tma)}</td>
                <td><span class="ez-badge">${p.topClass}</span></td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>`;

  document.getElementById('ez-main').innerHTML=html;
  buildHeatmap(data);
}


/* ══ MAPA DE CALOR — PICOS DE DEMANDA ══ */
function buildHeatmap(data) {
  const el = document.getElementById('ez-heatmap');
  if (!el) return;

  const DAYS = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

  // Matriz [diaDaSemana 0-6][hora 0-23] = contagem
  const matrix = Array.from({length: 7}, () => new Array(24).fill(0));

  data.forEach(d => {
    if (d.Hour < 0 || d.Hour > 23) return;
    const dt = new Date(d.DataStr + 'T12:00:00');
    if (isNaN(dt)) return;
    matrix[dt.getDay()][d.Hour]++;
  });

  const maxVal = Math.max(...matrix.flat(), 1);

  // Paleta Germânia: transparente → creme → gold → vermelho
  function getColor(val) {
    if (val === 0) return 'rgba(200,185,160,0.10)';
    const t = val / maxVal;
    if (t <= 0.33) {
      const p = t / 0.33;
      // #F1E3CE → #FFA62C
      return `rgb(${Math.round(241+(255-241)*p)},${Math.round(227+(166-227)*p)},${Math.round(206+(44-206)*p)})`;
    } else {
      const p = (t - 0.33) / 0.67;
      // #FFA62C → #C92B1E
      return `rgb(${Math.round(255+(201-255)*p)},${Math.round(166+(43-166)*p)},${Math.round(44+(30-44)*p)})`;
    }
  }

  function textColor(val) {
    return (val / maxVal) > 0.45 ? '#FFF8F0' : '#8B7040';
  }

  // Dimensões — largura total, células generosas
  const cellW = 20, cellH = 24;
  const leftPad = 32, topPad = 18, bottomPad = 18;
  const svgW = leftPad + 24 * cellW + 4;
  const svgH = topPad + 7 * cellH + bottomPad;

  let inner = '';

  // Labels de hora no topo (a cada 3h)
  for (let h = 0; h < 24; h++) {
    if (h % 3 === 0) {
      inner += `<text x="${leftPad + h*cellW + cellW/2}" y="${topPad - 4}"
        text-anchor="middle" font-family="Barlow Condensed,sans-serif"
        font-size="9" fill="#A89870">${String(h).padStart(2,'0')}h</text>`;
    }
  }

  // Células + labels de dia
  for (let d = 0; d < 7; d++) {
    const y = topPad + d * cellH;
    inner += `<text x="${leftPad - 4}" y="${y + cellH/2 + 4}"
      text-anchor="end" font-family="Barlow Condensed,sans-serif"
      font-size="9" font-weight="600" fill="#A89870">${DAYS[d]}</text>`;

    for (let h = 0; h < 24; h++) {
      const x = leftPad + h * cellW;
      const val = matrix[d][h];
      const fill = getColor(val);
      inner += `<rect x="${x+1}" y="${y+1}" width="${cellW-2}" height="${cellH-2}"
        rx="2" fill="${fill}"
        data-val="${val}" data-day="${DAYS[d]}" data-hour="${String(h).padStart(2,'0')}h"/>`;
      if (val > 0) {
        inner += `<text x="${x + cellW/2}" y="${y + cellH/2 + 4}"
          text-anchor="middle" font-family="Barlow Condensed,sans-serif"
          font-size="10" font-weight="600" fill="${textColor(val)}"
          pointer-events="none">${val}</text>`;
      }
    }
  }

  // Linha de total por hora na base
  inner += `<text x="${leftPad - 4}" y="${topPad + 7*cellH + bottomPad - 4}"
    text-anchor="end" font-family="Barlow Condensed,sans-serif"
    font-size="9" fill="rgba(168,152,112,0.6)">total</text>`;
  for (let h = 0; h < 24; h++) {
    const tot = matrix.reduce((s, row) => s + row[h], 0);
    if (tot > 0) {
      inner += `<text x="${leftPad + h*cellW + cellW/2}" y="${topPad + 7*cellH + bottomPad - 5}"
        text-anchor="middle" font-family="Barlow Condensed,sans-serif"
        font-size="9" fill="rgba(168,152,112,0.7)">${tot}</text>`;
    }
  }

  el.innerHTML = `<svg viewBox="0 0 ${svgW} ${svgH}" width="100%"
    xmlns="http://www.w3.org/2000/svg" style="display:block;overflow:visible;">${inner}</svg>`;

  // Tooltip
  const oldTip = document.querySelector('.sp-tip[data-id="heatmap"]');
  if (oldTip) oldTip.remove();
  const tip = document.createElement('div');
  tip.className = 'sp-tip';
  tip.dataset.id = 'heatmap';
  document.body.appendChild(tip);

  el.querySelectorAll('rect').forEach(r => {
    r.style.cursor = 'default';
    r.addEventListener('mouseenter', e => {
      const val = r.dataset.val;
      if (val === '0') return;
      tip.textContent = `${r.dataset.day} ${r.dataset.hour}: ${val} ticket${val != '1' ? 's' : ''}`;
      tip.style.left = (e.clientX + 12) + 'px';
      tip.style.top  = (e.clientY - 32) + 'px';
      tip.style.opacity = '1';
    });
    r.addEventListener('mousemove', e => {
      tip.style.left = (e.clientX + 12) + 'px';
      tip.style.top  = (e.clientY - 32) + 'px';
    });
    r.addEventListener('mouseleave', () => { tip.style.opacity = '0'; });
  });
}

function go(){
  const de=document.getElementById('f-de').value;
  const ate=document.getElementById('f-ate').value;
  const resp=document.getElementById('f-resp').value;

  const rdB=RD.filter(d=>{if(de&&d.DataStr<de)return false;if(ate&&d.DataStr>ate)return false;if(resp&&d.Resp!==resp)return false;return true;});
  const ez=EZ_PROTO.filter(d=>{if(de&&d.DataStr<de)return false;if(ate&&d.DataStr>ate)return false;if(resp&&d.Agente!==resp)return false;return true;});

  const orcs=rdB.filter(d=>ORC_ETAPAS.includes(d.Etapa)&&d['Valor Único']>0);
  const vend=rdB.filter(d=>d.Estado==='Vendida');
  const tAt=ez.length,tOrc=orcs.length,tPed=vend.length;
  const tLit=vend.reduce((s,d)=>s+(d.Litragem_num||0),0);
  const tRec=vend.reduce((s,d)=>s+(d['Valor Único']||0),0);
  const tmP=tPed?tRec/tPed:0,tmL=tPed?tLit/tPed:0,rL=tLit?tRec/tLit:0;
  const d=dias(de,ate),aL=tLit?tLit/d:0,aR=tRec?tRec/d:0;

  const _M=getMetasMes(de,ate);
  const mrO=mpr(_M.orc,de,ate,resp),mrL=mpr(_M.lit,de,ate,resp);
  const mrR=mpr(_M.rec,de,ate,resp),mrP=mpr(_M.ped,de,ate,resp);
  const sO=st(tOrc,mrO),sL=st(tLit,mrL),sR=st(tRec,mrR),sP=st(tPed,mrP);

  // Valores
  document.getElementById('v-at').textContent=fmt(tAt);
  document.getElementById('v-orc').textContent=fmt(tOrc);
  document.getElementById('v-ped').textContent=fmt(tPed);
  document.getElementById('v-lit').innerHTML=fmt(Math.round(tLit))+'<span class="u"> L</span>';
  document.getElementById('v-rec').textContent='R$ '+fmt(Math.round(tRec));
  document.getElementById('v-tp').textContent='R$ '+fmt(Math.round(tmP));
  document.getElementById('v-tl').innerHTML=tmL.toFixed(1)+'<span class="u"> L</span>';
  document.getElementById('v-rl').textContent='R$ '+rL.toFixed(2).replace('.',',');

  // Bezels — apenas taxa de conversão no formato "X% de conversão"
  const convAlc = ALC ? (tAt/ALC*100).toFixed(1)+'% de conversão' : '—';
  const convAtOrc = tAt ? (tOrc/tAt*100).toFixed(1)+'% de conversão' : '—';
  const convOrcPed = tOrc ? (tPed/tOrc*100).toFixed(1)+'% de conversão' : '—';
  const avgLitDia = fmt(Math.round(aL))+' L/dia em média';
  const avgRecDia = 'R$ '+fmt(Math.round(aR))+'/dia em média';

  document.getElementById('bz-alc').textContent = convAlc;
  document.getElementById('bz-at').textContent  = convAtOrc;
  document.getElementById('bz-orc').textContent = convAtOrc;
  document.getElementById('bz-ped').textContent = convOrcPed;
  document.getElementById('bz-lit').textContent = avgLitDia;
  document.getElementById('bz-rec').textContent = avgRecDia;
  document.getElementById('bz-tp').textContent  = '~R$ '+fmt(Math.round(tmP))+' por evento';
  document.getElementById('bz-tl').textContent  = '~'+tmL.toFixed(1)+' L por evento';
  document.getElementById('bz-rl').textContent  = '~R$ '+rL.toFixed(2).replace('.',',')+' por litro vendido';

  // Status
  setS('c-orc',sO);setS('c-ped',sP);setS('c-lit',sL);setS('c-rec',sR);

  // Tooltips
  setTip('ct-alc',fmt(tAt)+' atend de '+fmt(ALC)+' alcance');
  setTip('ct-at', fmt(tOrc)+' orc de '+fmt(tAt)+' atend');
  setTip('ct-orc','Meta: '+fmt(Math.round(mrO))+' · Real: '+fmt(tOrc));
  setTip('ct-ped','Meta: '+fmt(Math.round(mrP))+' · Real: '+fmt(tPed));
  setTip('ct-lit','Meta: '+fmt(Math.round(mrL))+'L · Real: '+fmt(Math.round(tLit))+'L');
  setTip('ct-rec','Meta: R$ '+fmt(Math.round(mrR))+' · Real: R$ '+fmt(Math.round(tRec)));

  // Círculos
  circ('ci-alc', ALC?(tAt/ALC*100):0,  '#9BA8B0', pl(tAt,ALC));
  circ('ci-at',  tAt?(tOrc/tAt*100):0, '#9BA8B0', tAt?Math.round(tOrc/tAt*100)+'%':'—');
  circ('ci-orc', tOrc?tOrc/mrO*100:0,  SC[sO],    pl(tOrc,mrO));
  circ('ci-ped', tPed?tPed/mrP*100:0,  SC[sP],    pl(tPed,mrP));
  circ('ci-lit', tLit?tLit/mrL*100:0,  SC[sL],    pl(tLit,mrL));
  circ('ci-rec', tRec?tRec/mrR*100:0,  SC[sR],    pl(tRec,mrR));

  // Sparklines
  // Semanas únicas no período
  const semsSet=[...new Set(vend.map(d=>d.Semana))].sort((a,b)=>a-b).slice(-4);
  SEM = semsSet.map((s,i) => {
    // Calcular data de início da semana
    const ano = new Date(de).getFullYear();
    const jan1 = new Date(ano,0,1);
    const d = new Date(jan1.getTime() + (s-1)*7*86400000);
    return 'S'+(i+1)+'·'+String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0');
  });
  const spP=[],spL=[],spR=[];
  semsSet.forEach(s=>{
    const v=vend.filter(d=>d.Semana===s);
    const p=v.length,l=v.reduce((a,d)=>a+(d.Litragem_num||0),0),r=v.reduce((a,d)=>a+(d['Valor Único']||0),0);
    spP.push(p?r/p:0);spL.push(p?l/p:0);spR.push(l?r/l:0);
  });
  requestAnimationFrame(()=>{
    spark('sp-tp',spP,SEM,fR);
    spark('sp-tl',spL,SEM,fL);
    spark('sp-rl',spR,SEM,v=>'R$'+v.toFixed(2).replace('.',','));
  });

  // Header
  hkpi('hk-lit','hv-lit','ht-lit',tLit,mrL,v=>fmt(Math.round(v))+'<span class="u"> L</span>');
  hkpi('hk-rec','hv-rec','ht-rec',tRec,mrR,v=>'R$'+Math.round(v/1000)+'k');
  document.getElementById('hv-cv').textContent=tOrc?(tPed/tOrc*100).toFixed(1)+'%':'—';
  const cvPct=tOrc?(tPed/tOrc*100):0;
  const cvBar=document.getElementById('hb-cv');
  if(cvBar){
    const s=cvPct>=50?'green':cvPct>=30?'yellow':'red';
    cvBar.className='hk-bar '+s;
    requestAnimationFrame(()=>requestAnimationFrame(()=>{cvBar.style.width=Math.min(cvPct,100).toFixed(1)+'%';}));
    const cvPctEl=document.getElementById('hp-cv');
    if(cvPctEl)cvPctEl.textContent=Math.min(cvPct,100).toFixed(0)+'%';
  }
  // h-sub removido
  const MESES=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const deDate=new Date(de+'T12:00:00');
  const ateDate=new Date(ate+'T12:00:00');
  const sameMonth=deDate.getMonth()===ateDate.getMonth()&&deDate.getFullYear()===ateDate.getFullYear();
  const resumoEl=document.getElementById('h-resumo-lbl');
  if(resumoEl)resumoEl.textContent=sameMonth?'Resumo '+MESES[deDate.getMonth()]:'Resumo do Período';
}

function reset(){
  const today=new Date();
  const pad=n=>String(n).padStart(2,'0');
  const y=today.getFullYear(),m=today.getMonth();
  const lastDay=new Date(y,m+1,0).getDate();
  document.getElementById('f-de').value=y+'-'+pad(m+1)+'-01';
  document.getElementById('f-ate').value=y+'-'+pad(m+1)+'-'+pad(lastDay);
  document.getElementById('f-resp').value='';

  document.querySelectorAll('.btn-sh').forEach(b=>b.classList.remove('active'));
  go();
}
/* ── DATAS PADRÃO — mês atual ── */
(function() {
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth()+1).padStart(2,'0');
  const last = new Date(today.getFullYear(), today.getMonth()+1, 0).getDate();
  const de  = document.getElementById('f-de');
  const ate = document.getElementById('f-ate');
  if (de)  de.value  = y+'-'+m+'-01';
  if (ate) ate.value = y+'-'+m+'-'+String(last).padStart(2,'0');
})();

const _upd=document.getElementById('upd-date');if(_upd){const _d=new Date();_upd.textContent=String(_d.getDate()).padStart(2,'0')+'/'+String(_d.getMonth()+1).padStart(2,'0')+'/'+_d.getFullYear();}
loadData();
