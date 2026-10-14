/* ============================================================
   nReport Dashboard — lógica (100% client-side)
   ============================================================ */

'use strict';

/* ---------- estado global ---------- */
let RAW = [];        // linhas cruas (objetos) vindas da planilha
let ROWS = [];       // linhas normalizadas
let CHARTS = {};     // instâncias Chart.js (para destruir ao recarregar)

/* ---------- utilitários de data ---------- */
// Aceita DD/MM/AAAA, DD/MM/AA, AAAA-MM-DD e serial de Excel
function parseDate(v){
  if(v===null||v===undefined||v==='') return null;
  if(v instanceof Date && !isNaN(v)) return v;
  if(typeof v==='number'){ // serial do Excel
    const d=XLSX.SSF ? XLSX.SSF.parse_date_code(v) : null;
    if(d) return new Date(d.y, d.m-1, d.d);
  }
  const s=String(v).trim();
  let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if(m){
    let [_,dd,mm,yy]=m; dd=+dd; mm=+mm; yy=+yy;
    if(yy<100) yy+=2000;
    const d=new Date(yy,mm-1,dd);
    return plausibleDate(d);
  }
  m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m){ return plausibleDate(new Date(+m[1],+m[2]-1,+m[3])); }
  const d=new Date(s); return plausibleDate(d);
}
// Rejeita datas impossíveis: inválidas, no futuro (depois de hoje)
// ou antes de 2019 (antes da plataforma existir). Erros de digitação de ano
// (ex.: 2028) deixam de distorcer gráficos e período.
function plausibleDate(d){
  if(!d || isNaN(d)) return null;
  const hoje=new Date();
  const limiteFuturo=new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()); // hoje 00:00
  const limitePassado=new Date(2019,0,1);
  if(d > limiteFuturo) return null;   // data no futuro = erro de digitação
  if(d < limitePassado) return null;  // antes de 2019 = suspeita
  return d;
}
const ymKey = d => d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}` : null;
const ymLabel = k => {
  if(!k) return '—';
  const [y,m]=k.split('-');
  const nomes=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  return `${nomes[+m-1]}/${y}`;
};
const daysBetween = (a,b)=> (a&&b)? Math.round((b-a)/86400000) : null;
const fmt = n => (n==null||isNaN(n))?'—':n.toLocaleString('pt-BR');

/* ---------- conserto de acentos quebrados (mojibake de CSV Latin-1/UTF-8) ---------- */
// Ex.: "Indicaç�" / "SÃ£o Paul" -> "Indicação" / "São Paulo" (quando possível)
const MOJIBAKE_MAP = {
  'Ã£':'ã','Ã¡':'á','Ã¢':'â','Ã ':'à','Ã©':'é','Ãª':'ê','Ã­':'í',
  'Ã³':'ó','Ã´':'ô','Ãµ':'õ','Ãº':'ú','Ã§':'ç','Ã‰':'É','Ã“':'Ó',
  'Ã‡':'Ç','Ãš':'Ú','Ã':'Á','Ã‚':'Â','Ãƒ':'Ã','Ã”':'Ô','Ã•':'Õ',
  'Ã©':'é','Ã¨':'è'
};
function fixText(s){
  if(s==null) return s;
  let t=String(s);
  // substitui sequências conhecidas de mojibake
  for(const [bad,good] of Object.entries(MOJIBAKE_MAP)) t=t.split(bad).join(good);
  // remove o caractere de substituição (�) usado quando o acento se perdeu
  t=t.replace(/\uFFFD/g,'');
  return t.trim();
}

/* ---------- normalização de UF ---------- */
// As 27 siglas válidas
const UFS = new Set(['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS',
  'MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO']);
// nome do estado (sem acento/caixa) -> UF
const ESTADO_NOME_UF = {
  'acre':'AC','alagoas':'AL','amapa':'AP','amazonas':'AM','bahia':'BA',
  'ceara':'CE','distrito federal':'DF','espirito santo':'ES','goias':'GO',
  'maranhao':'MA','mato grosso':'MT','mato grosso do sul':'MS','minas gerais':'MG',
  'para':'PA','paraiba':'PB','parana':'PR','pernambuco':'PE','piaui':'PI',
  'rio de janeiro':'RJ','rio grande do norte':'RN','rio grande do sul':'RS',
  'rondonia':'RO','roraima':'RR','santa catarina':'SC','sao paulo':'SP',
  'sergipe':'SE','tocantins':'TO','solteiro':null,'solteiro(a)':null
};
// Normaliza o campo "Estado" para uma UF de 2 letras quando possível.
// Trata lixo (cidade no lugar da UF, acento quebrado, texto cortado "SÃ£O PAUL").
function normalizeUF(raw){
  let s=fixText(raw);
  if(!s) return '—';
  const n=norm(s);                 // sem acento, minúsculo
  const up=s.toUpperCase().trim();
  // 1) já é uma UF válida de 2 letras
  if(up.length===2 && UFS.has(up)) return up;
  // 2) nome de estado conhecido (inteiro)
  if(ESTADO_NOME_UF[n]!==undefined) return ESTADO_NOME_UF[n]||'—';
  // 3) começa com "sao paul..." (cobre "sao paulo", "são paul", "sÃ£o paul" já consertado)
  if(/^sao paul/.test(n)) return 'SP';
  if(/^rio de jan/.test(n)) return 'RJ';
  if(/^belo horiz/.test(n)) return 'MG';
  // 4) heurística: pega as 2 primeiras letras se formarem UF (último recurso fraco) -> evita
  // se nada bater, devolve o texto limpo em maiúsculas (fica visível que é "outro")
  return up || '—';
}

/* ---------- detecção flexível de colunas ---------- */
// mapeia o cabeçalho real -> campos canônicos (tolerante a acento/caixa)
const FIELD_ALIASES = {
  nome:['nome','name'],
  email:['email','e-mail'],
  emailConfirmado:['emailconfirmado','email confirmado'],
  celular:['celular','telefone','phone'],
  cidade:['cidade','city'],
  estado:['estado','uf','state'],
  comoConheceu:['comoconheceu','como conheceu','origem','fonte'],
  dataCadastro:['datacadastro','data cadastro','data de cadastro','cadastro'],
  qtdLogin:['qtdlogin','qtd login','logins','quantidade login','qtd de login'],
  ultimoAcesso:['ultimoacesso','ultimo acesso','último acesso'],
  qtdLaudos:['qtdlaudos','qtd laudos','laudos','quantidade laudos','qtd de laudos'],
  dataPrimeiroLaudo:['dataprimeirolaudo','data primeiro laudo','primeiro laudo'],
  tipoAssinatura:['tipoassinatura','tipo assinatura','tipo','assinatura','plano'],
  dataAssinatura:['dataassinatura','data assinatura','data de assinatura'],
  dataCancelamento:['datacancelamento','data cancelamento','data de cancelamento','cancelamento'],
};
const norm = s => String(s||'').trim().toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g,'');

function buildColumnMap(headers){
  const map={}; const missing=[];
  for(const [field,aliases] of Object.entries(FIELD_ALIASES)){
    const hit = headers.find(h=> aliases.includes(norm(h)));
    if(hit) map[field]=hit; else missing.push(field);
  }
  return {map,missing};
}

/* ---------- classificação do tipo (estado atual) ---------- */
// normaliza para: assinante | gratuito | exassinante | cortesia | outro
function classify(tipoRaw){
  const t=norm(tipoRaw);
  if(t.startsWith('ex')) return 'exassinante';   // "ex-assinante"
  if(t.startsWith('assin')) return 'assinante';
  if(t.startsWith('grat')) return 'gratuito';
  if(t.startsWith('cortes')) return 'cortesia';
  return 'outro';
}

/* ---------- heurística de registro de teste/lixo ---------- */
const TEST_EMAIL_DOMAINS = ['mailinator','uorak','tempmail','teste','example',
  'yopmail','guerrilla','sharklasers','niekie','revutap','obxstorm','relumyx',
  'ckuer','gixpos','ptiong','kobace','niprack','hazhab','him6','gopicta','nalwan',
  'jxbav','fidhost','daugr','jalunaki','maazios','gamepec','pamaweb','gpipes',
  'inkmoto','steamoh','viemery','conca','abowned','art2mart','hebase','marvetos',
  'bittnex','fenxz','lidugw','hutdot','etramay','draughtier','certve','bizmud',
  'hopesx','cadebek','rograc','nuitx','prodbits','soppat','agiuse','him6','jucatyo'];
function looksTest(r){
  const nome=norm(r.nome), email=String(r.email||'').toLowerCase();
  if(!nome) return true;
  // nomes claramente de teste
  if(/^(teste|test|aaa+|xxx+|abc|asd|dfg|qwe|fulano|ciclano|beltrano|zzz|ddd|eee|www|kkk|nome)/.test(nome)) return true;
  if(/teste/.test(nome)) return true;
  // nome com caractere repetido longo (ex: "aaaaaaa")
  if(/(.)\1{5,}/.test(nome.replace(/\s/g,''))) return true;
  // domínios descartáveis
  const dom=email.split('@')[1]||'';
  if(TEST_EMAIL_DOMAINS.some(d=>dom.includes(d))) return true;
  // email sem @
  if(email && !email.includes('@')) return true;
  return false;
}

/* ---------- normalização de uma linha ---------- */
function normalizeRow(raw, cmap){
  const g = f => cmap[f]!==undefined ? raw[cmap[f]] : undefined;
  const dCad=parseDate(g('dataCadastro'));
  const dAss=parseDate(g('dataAssinatura'));
  const dCanc=parseDate(g('dataCancelamento'));
  const dLaudo=parseDate(g('dataPrimeiroLaudo'));
  const dAcesso=parseDate(g('ultimoAcesso'));
  const r={
    nome:fixText(g('nome')), email:g('email'),
    estado:normalizeUF(g('estado')),
    cidade:fixText(g('cidade')),
    origem:(fixText(g('comoConheceu'))||'—'),
    qtdLogin:+(g('qtdLogin')||0)||0,
    qtdLaudos:+(g('qtdLaudos')||0)||0,
    dataCadastro:dCad, dataAssinatura:dAss, dataCancelamento:dCanc,
    dataPrimeiroLaudo:dLaudo, ultimoAcesso:dAcesso,
    tipo:classify(g('tipoAssinatura')),
    tipoRaw:String(g('tipoAssinatura')||'').trim(),
    diasAteAssinatura: daysBetween(dCad,dAss),
    diasAtePrimeiroLaudo: daysBetween(dCad,dLaudo),
  };
  r.isTest=looksTest(r);
  return r;
}

/* ============================================================
   LEITURA DE ARQUIVO
   ============================================================ */
function handleFile(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const data=new Uint8Array(e.target.result);
      const wb=XLSX.read(data,{type:'array',cellDates:true,raw:false});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const json=XLSX.utils.sheet_to_json(ws,{defval:'',raw:false});
      if(!json.length){ alert('A planilha parece vazia.'); return; }
      const headers=Object.keys(json[0]);
      const {map,missing}=buildColumnMap(headers);
      if(map.dataCadastro===undefined || map.tipoAssinatura===undefined){
        alert('Não encontrei as colunas essenciais (DataCadastro e TipoAssinatura). Verifique o cabeçalho.');
        return;
      }
      RAW=json;
      ROWS=json.map(r=>normalizeRow(r,map));
      showMappingWarning(missing);
      initApp();
    }catch(err){ console.error(err); alert('Erro ao ler o arquivo: '+err.message); }
  };
  reader.readAsArrayBuffer(file);
}

function showMappingWarning(missing){
  const box=document.getElementById('mapWarn');
  const relevant=missing.filter(m=>['dataPrimeiroLaudo','dataAssinatura','dataCancelamento','qtdLaudos','comoConheceu','estado'].includes(m));
  if(relevant.length){
    box.classList.remove('hidden');
    box.innerHTML='⚠️ Colunas não encontradas (alguns gráficos ficam limitados): <b>'+relevant.join(', ')+'</b>';
  }else box.classList.add('hidden');
}

/* ============================================================
   DADOS FILTRADOS
   ============================================================ */
function currentData(){
  const hideTest=document.getElementById('chkFilterTest').checked;
  return hideTest ? ROWS.filter(r=>!r.isTest) : ROWS;
}

/* ============================================================
   INICIALIZAÇÃO DO APP
   ============================================================ */
function initApp(){
  document.getElementById('dropArea').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');
  document.getElementById('statusPill').textContent='Planilha carregada ✓';
  buildMonthSelector();
  renderAll();
  renderSnapshots();
}

function buildMonthSelector(){
  const data=currentData();
  const months=[...new Set(data.map(r=>ymKey(r.dataCadastro)).filter(Boolean))].sort().reverse();
  const sel=document.getElementById('selMonth');
  sel.innerHTML=months.map(m=>`<option value="${m}">${ymLabel(m)}</option>`).join('');
}

function renderAll(){
  const data=currentData();
  // chips
  document.getElementById('chTotal').textContent=fmt(data.length);
  document.getElementById('chTest').textContent=fmt(ROWS.filter(r=>r.isTest).length);
  const cadDates=data.map(r=>r.dataCadastro).filter(Boolean).sort((a,b)=>a-b);
  document.getElementById('chRange').textContent=cadDates.length
    ? `${cadDates[0].toLocaleDateString('pt-BR')} → ${cadDates[cadDates.length-1].toLocaleDateString('pt-BR')}` : '—';

  renderKPIs(data);
  renderGrowth(data);
  renderSubCanc(data);
  renderSource(data);
  renderState(data);
  renderReports(data);
  renderMonth();
}

/* ---------- KPIs gerais (estado atual) ---------- */
function renderKPIs(data){
  const total=data.length;
  const c={assinante:0,gratuito:0,exassinante:0,cortesia:0,outro:0};
  data.forEach(r=>c[r.tipo]++);
  const pct=n=> total? (100*n/total).toFixed(1)+'% da base':'—';
  document.getElementById('kAss').textContent=fmt(c.assinante);
  document.getElementById('kAssPct').textContent=pct(c.assinante);
  document.getElementById('kGra').textContent=fmt(c.gratuito);
  document.getElementById('kGraPct').textContent=pct(c.gratuito);
  document.getElementById('kEx').textContent=fmt(c.exassinante);
  document.getElementById('kExPct').textContent=pct(c.exassinante);
  document.getElementById('kCor').textContent=fmt(c.cortesia);
  document.getElementById('kCorPct').textContent=pct(c.cortesia);
  document.getElementById('kTot').textContent=fmt(total);

  // tempo médio cadastro -> assinatura (quem tem as duas datas)
  const difs=data.map(r=>r.diasAteAssinatura).filter(d=>d!=null && d>=0);
  const avg= difs.length? Math.round(difs.reduce((a,b)=>a+b,0)/difs.length) : null;
  document.getElementById('kConv').textContent= avg!=null? avg+' dias' : '—';
}

/* ---------- gráfico: crescimento (cadastros/mês) ---------- */
function renderGrowth(data){
  const byMonth={};
  data.forEach(r=>{const k=ymKey(r.dataCadastro); if(k) byMonth[k]=(byMonth[k]||0)+1;});
  const keys=Object.keys(byMonth).sort();
  drawChart('chartGrowth','line',{
    labels:keys.map(ymLabel),
    datasets:[{label:'Novos cadastros',data:keys.map(k=>byMonth[k]),
      borderColor:'#4f8cff',backgroundColor:'rgba(79,140,255,.15)',fill:true,tension:.3,pointRadius:2}]
  });
}

/* ---------- gráfico: assinaturas vs cancelamentos por mês ---------- */
function renderSubCanc(data){
  const sub={},canc={};
  data.forEach(r=>{
    const ks=ymKey(r.dataAssinatura); if(ks) sub[ks]=(sub[ks]||0)+1;
    const kc=ymKey(r.dataCancelamento); if(kc) canc[kc]=(canc[kc]||0)+1;
  });
  const keys=[...new Set([...Object.keys(sub),...Object.keys(canc)])].sort();
  drawChart('chartSubCanc','bar',{
    labels:keys.map(ymLabel),
    datasets:[
      {label:'Novas assinaturas',data:keys.map(k=>sub[k]||0),backgroundColor:'#2bd4a7'},
      {label:'Cancelamentos',data:keys.map(k=>canc[k]||0),backgroundColor:'#ff5d6c'}
    ]
  });
}

/* ---------- gráfico: origem ---------- */
function renderSource(data){
  const by={};
  data.forEach(r=>{const o=r.origem||'—'; by[o]=(by[o]||0)+1;});
  const entries=Object.entries(by).sort((a,b)=>b[1]-a[1]);
  drawChart('chartSource','doughnut',{
    labels:entries.map(e=>e[0]),
    datasets:[{data:entries.map(e=>e[1]),
      backgroundColor:['#4f8cff','#2bd4a7','#ffb020','#a98bff','#ff5d6c','#5ad1e6','#f78fb3','#8c9bb5','#c9d1e0']}]
  },{plugins:{legend:{position:'right',labels:{color:'#9aa7bd',font:{size:11}}}}});
}

/* ---------- gráfico: estados ---------- */
function renderState(data){
  const by={};
  data.forEach(r=>{const s=r.estado||'—'; by[s]=(by[s]||0)+1;});
  const entries=Object.entries(by).sort((a,b)=>b[1]-a[1]).slice(0,12);
  drawChart('chartState','bar',{
    labels:entries.map(e=>e[0]),
    datasets:[{label:'Usuários',data:entries.map(e=>e[1]),backgroundColor:'#4f8cff'}]
  },{indexAxis:'y'});
}

/* ---------- gráfico: faixas de laudos ---------- */
function reportBucket(n){
  if(n<=0) return '0 laudos';
  if(n<=10) return '1–10';
  if(n<=50) return '11–50';
  if(n<=200) return '51–200';
  return '200+';
}
function renderReports(data){
  const order=['0 laudos','1–10','11–50','51–200','200+'];
  const by={}; order.forEach(o=>by[o]=0);
  data.forEach(r=>by[reportBucket(r.qtdLaudos)]++);
  drawChart('chartReports','bar',{
    labels:order,
    datasets:[{label:'Usuários',data:order.map(o=>by[o]),
      backgroundColor:['#8c9bb5','#4f8cff','#2bd4a7','#ffb020','#a98bff']}]
  },{plugins:{legend:{display:false}}});
}

/* ---------- fechamento mensal ---------- */
function renderMonth(){
  const data=currentData();
  const mk=document.getElementById('selMonth').value;
  if(!mk) return;

  const novos=data.filter(r=>ymKey(r.dataCadastro)===mk);
  const assinaturas=data.filter(r=>ymKey(r.dataAssinatura)===mk);
  const cancels=data.filter(r=>ymKey(r.dataCancelamento)===mk);

  document.getElementById('mNew').textContent=fmt(novos.length);
  document.getElementById('mSub').textContent=fmt(assinaturas.length);
  document.getElementById('mCanc').textContent=fmt(cancels.length);

  // taxa de ativação: dos novos, quantos fizeram >=1 laudo
  const ativados=novos.filter(r=>r.qtdLaudos>0).length;
  document.getElementById('mAct').textContent= novos.length
    ? (100*ativados/novos.length).toFixed(0)+'%' : '—';

  document.getElementById('monthHint').textContent=
    `${novos.length} novos cadastros em ${ymLabel(mk)}`;

  // comportamento dos novos gratuitos do mês (foco do pedido original)
  const gratuitosDoMes=novos.filter(r=>r.tipo==='gratuito');
  const semLaudo=gratuitosDoMes.filter(r=>r.qtdLaudos===0).length;
  const ate10=gratuitosDoMes.filter(r=>r.qtdLaudos>=1&&r.qtdLaudos<=10).length;
  const mais10=gratuitosDoMes.filter(r=>r.qtdLaudos>10).length;

  document.getElementById('freeMonthSub').textContent=
    `${gratuitosDoMes.length} gratuitos cadastrados em ${ymLabel(mk)}`;

  drawChart('chartFreeBehavior','doughnut',{
    labels:['Nenhum laudo','1–10 laudos','Mais de 10 laudos'],
    datasets:[{data:[semLaudo,ate10,mais10],
      backgroundColor:['#ff5d6c','#4f8cff','#2bd4a7']}]
  },{plugins:{legend:{position:'bottom',labels:{color:'#9aa7bd',font:{size:11}}}}});
}

/* ---------- helper Chart.js ---------- */
function drawChart(id,type,data,extra={}){
  const ctx=document.getElementById(id).getContext('2d');
  if(CHARTS[id]) CHARTS[id].destroy();
  const base={
    responsive:true,maintainAspectRatio:false,
    plugins:{legend:{labels:{color:'#9aa7bd'}}},
    scales:(type==='doughnut'||type==='pie')?{}:{
      x:{ticks:{color:'#9aa7bd',font:{size:11}},grid:{color:'rgba(255,255,255,.04)'}},
      y:{ticks:{color:'#9aa7bd'},grid:{color:'rgba(255,255,255,.04)'}}
    }
  };
  // merge raso
  const opts={...base,...extra,plugins:{...base.plugins,...(extra.plugins||{})},
    scales:extra.scales||base.scales};
  CHARTS[id]=new Chart(ctx,{type,data,options:opts});
}

/* ============================================================
   SNAPSHOTS (IndexedDB) — só agregados, sem dados pessoais
   ============================================================ */
const DB_NAME='nreport_snapshots', STORE='snaps';
function openDB(){
  return new Promise((res,rej)=>{
    const req=indexedDB.open(DB_NAME,1);
    req.onupgradeneeded=()=>req.result.createObjectStore(STORE,{keyPath:'id'});
    req.onsuccess=()=>res(req.result);
    req.onerror=()=>rej(req.error);
  });
}
async function saveSnapshot(){
  const data=currentData();
  const c={assinante:0,gratuito:0,exassinante:0,cortesia:0};
  data.forEach(r=>{if(c[r.tipo]!==undefined)c[r.tipo]++;});
  const snap={
    id:new Date().toISOString().slice(0,10), // 1 por dia (último do dia vence)
    data:new Date().toISOString(),
    assinante:c.assinante,gratuito:c.gratuito,
    exassinante:c.exassinante,cortesia:c.cortesia,total:data.length
  };
  const db=await openDB();
  await new Promise((res,rej)=>{
    const tx=db.transaction(STORE,'readwrite');
    tx.objectStore(STORE).put(snap);
    tx.oncomplete=res; tx.onerror=()=>rej(tx.error);
  });
  renderSnapshots();
  document.getElementById('statusPill').textContent='Snapshot salvo ✓';
}
async function getSnapshots(){
  const db=await openDB();
  return new Promise((res,rej)=>{
    const tx=db.transaction(STORE,'readonly');
    const req=tx.objectStore(STORE).getAll();
    req.onsuccess=()=>res(req.result.sort((a,b)=>b.id.localeCompare(a.id)));
    req.onerror=()=>rej(req.error);
  });
}
async function deleteSnapshot(id){
  const db=await openDB();
  await new Promise(res=>{
    const tx=db.transaction(STORE,'readwrite');
    tx.objectStore(STORE).delete(id); tx.oncomplete=res;
  });
  renderSnapshots();
}
async function renderSnapshots(){
  const snaps=await getSnapshots();
  const tb=document.querySelector('#snapTable tbody');
  const no=document.getElementById('noSnap');
  if(!snaps.length){ tb.innerHTML=''; no.classList.remove('hidden'); return; }
  no.classList.add('hidden');
  tb.innerHTML=snaps.map(s=>`<tr>
    <td>${new Date(s.id).toLocaleDateString('pt-BR')}</td>
    <td class="num">${fmt(s.assinante)}</td>
    <td class="num">${fmt(s.gratuito)}</td>
    <td class="num">${fmt(s.exassinante)}</td>
    <td class="num">${fmt(s.cortesia)}</td>
    <td class="num"><b>${fmt(s.total)}</b></td>
    <td class="num"><button class="ghost" data-del="${s.id}" style="padding:4px 8px">🗑</button></td>
  </tr>`).join('');
  tb.querySelectorAll('[data-del]').forEach(b=>
    b.onclick=()=>deleteSnapshot(b.getAttribute('data-del')));
}

/* ============================================================
   EVENTOS DE UI
   ============================================================ */
const drop=document.getElementById('dropArea');
const input=document.getElementById('fileInput');
drop.onclick=()=>input.click();
input.onchange=e=>{ if(e.target.files[0]) handleFile(e.target.files[0]); };
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{
  e.preventDefault();drop.classList.add('hot');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{
  e.preventDefault();drop.classList.remove('hot');}));
drop.addEventListener('drop',e=>{ if(e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });

document.getElementById('chkFilterTest').onchange=()=>{buildMonthSelector();renderAll();};
document.getElementById('selMonth').onchange=renderMonth;
document.getElementById('btnSnap').onclick=saveSnapshot;
document.getElementById('btnReset').onclick=()=>{
  RAW=[];ROWS=[];
  document.getElementById('app').classList.add('hidden');
  document.getElementById('dropArea').classList.remove('hidden');
  document.getElementById('mapWarn').classList.add('hidden');
  document.getElementById('statusPill').textContent='Aguardando planilha…';
  input.value='';
};
