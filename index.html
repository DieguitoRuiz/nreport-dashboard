<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>nReport • Painel de Métricas de Usuários</title>
<!-- SheetJS: lê .xlsx e .csv localmente, sem enviar nada para servidor -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<!-- Chart.js: gráficos -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#0f1420; --panel:#171e2e; --panel2:#1e2638; --line:#2a344a;
    --txt:#e8edf6; --muted:#9aa7bd; --accent:#4f8cff; --accent2:#2bd4a7;
    --warn:#ffb020; --danger:#ff5d6c; --purple:#a98bff;
    --radius:14px; --shadow:0 6px 24px rgba(0,0,0,.28);
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--txt);
    font-family:'Segoe UI',system-ui,-apple-system,Roboto,sans-serif;font-size:14px}
  header{padding:20px 26px;border-bottom:1px solid var(--line);
    display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}
  header h1{margin:0;font-size:19px;font-weight:700;letter-spacing:.3px}
  header h1 span{color:var(--accent)}
  .sub{color:var(--muted);font-size:12.5px;margin-top:3px}
  .wrap{padding:22px 26px;max-width:1500px;margin:0 auto}

  /* Dropzone */
  .drop{border:2px dashed var(--line);border-radius:var(--radius);padding:34px;
    text-align:center;background:var(--panel);transition:.2s;cursor:pointer}
  .drop:hover,.drop.hot{border-color:var(--accent);background:var(--panel2)}
  .drop h2{margin:0 0 6px;font-size:16px}
  .drop p{margin:4px 0;color:var(--muted)}
  .drop .big{font-size:40px;margin-bottom:6px}

  .toolbar{display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin:18px 0}
  .toolbar .chip{background:var(--panel2);border:1px solid var(--line);
    padding:8px 12px;border-radius:10px;color:var(--muted);font-size:12.5px}
  .toolbar .chip b{color:var(--txt)}
  select,button{font-family:inherit;font-size:13px}
  select{background:var(--panel2);color:var(--txt);border:1px solid var(--line);
    padding:9px 12px;border-radius:10px;cursor:pointer}
  label.switch{display:inline-flex;align-items:center;gap:8px;cursor:pointer;color:var(--muted)}
  label.switch input{width:16px;height:16px;accent-color:var(--accent)}
  button.ghost{background:var(--panel2);color:var(--txt);border:1px solid var(--line);
    padding:9px 14px;border-radius:10px;cursor:pointer}
  button.ghost:hover{border-color:var(--accent)}

  /* KPI cards */
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:8px}
  .kpi{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);
    padding:16px 18px;box-shadow:var(--shadow)}
  .kpi .lab{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.6px}
  .kpi .val{font-size:30px;font-weight:800;margin-top:6px}
  .kpi .hint{font-size:12px;color:var(--muted);margin-top:4px}
  .kpi.accent .val{color:var(--accent)}
  .kpi.green .val{color:var(--accent2)}
  .kpi.warn .val{color:var(--warn)}
  .kpi.danger .val{color:var(--danger)}
  .kpi.purple .val{color:var(--purple)}

  .grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px;margin-top:16px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:var(--radius);
    padding:16px 18px;box-shadow:var(--shadow)}
  .card h3{margin:0 0 2px;font-size:14.5px}
  .card .csub{color:var(--muted);font-size:12px;margin-bottom:12px}
  .col-12{grid-column:span 12}.col-8{grid-column:span 8}.col-6{grid-column:span 6}
  .col-4{grid-column:span 4}.col-5{grid-column:span 5}.col-7{grid-column:span 7}
  @media(max-width:1000px){[class^=col-]{grid-column:span 12}}
  canvas{max-height:300px}

  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{text-align:left;padding:9px 10px;border-bottom:1px solid var(--line)}
  th{color:var(--muted);font-weight:600;font-size:11.5px;text-transform:uppercase;letter-spacing:.5px}
  td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
  tbody tr:hover{background:var(--panel2)}
  .badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11.5px;font-weight:600}
  .b-ass{background:rgba(43,212,167,.15);color:var(--accent2)}
  .b-gra{background:rgba(79,140,255,.15);color:var(--accent)}
  .b-ex{background:rgba(255,93,108,.15);color:var(--danger)}
  .b-cor{background:rgba(169,139,255,.15);color:var(--purple)}

  .hidden{display:none!important}
  .foot{color:var(--muted);font-size:12px;text-align:center;padding:24px;border-top:1px solid var(--line);margin-top:30px}
  .pill{font-size:11px;padding:3px 8px;border-radius:20px;background:var(--panel2);border:1px solid var(--line);color:var(--muted)}
  .warnbox{background:rgba(255,176,32,.08);border:1px solid rgba(255,176,32,.3);
    color:#ffd27a;padding:10px 14px;border-radius:10px;font-size:12.5px;margin-top:12px}
</style>
</head>
<body>
<header>
  <div>
    <h1><span>nReport</span> • Painel de Métricas de Usuários</h1>
    <div class="sub">100% local no seu navegador — nenhum dado é enviado para servidores. 🔒 (LGPD)</div>
  </div>
  <div class="pill" id="statusPill">Aguardando planilha…</div>
</header>

<div class="wrap">

  <!-- UPLOAD -->
  <div id="dropArea" class="drop">
    <div class="big">📊</div>
    <h2>Arraste a planilha aqui, ou clique para selecionar</h2>
    <p>Aceita <b>.xlsx</b> e <b>.csv</b> (separador <code>;</code> ou <code>,</code>)</p>
    <p style="font-size:12px">A leitura acontece inteiramente no seu computador.</p>
    <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" class="hidden">
  </div>

  <div id="mapWarn" class="warnbox hidden"></div>

  <!-- APP -->
  <div id="app" class="hidden">

    <div class="toolbar">
      <div class="chip">Registros válidos: <b id="chTotal">0</b></div>
      <div class="chip">Período: <b id="chRange">—</b></div>
      <label class="switch"><input type="checkbox" id="chkFilterTest" checked>
        Ocultar cadastros de teste/lixo (<span id="chTest">0</span>)</label>
      <div style="flex:1"></div>
      <button class="ghost" id="btnSnap">💾 Salvar snapshot do dia</button>
      <button class="ghost" id="btnReset">↻ Trocar planilha</button>
    </div>

    <!-- KPIs -->
    <div class="kpis">
      <div class="kpi green"><div class="lab">Assinantes ativos</div>
        <div class="val" id="kAss">0</div><div class="hint" id="kAssPct">—</div></div>
      <div class="kpi accent"><div class="lab">Gratuitos</div>
        <div class="val" id="kGra">0</div><div class="hint" id="kGraPct">—</div></div>
      <div class="kpi danger"><div class="lab">Ex-assinantes</div>
        <div class="val" id="kEx">0</div><div class="hint" id="kExPct">—</div></div>
      <div class="kpi purple"><div class="lab">Cortesia</div>
        <div class="val" id="kCor">0</div><div class="hint" id="kCorPct">—</div></div>
      <div class="kpi"><div class="lab">Total de usuários</div>
        <div class="val" id="kTot">0</div><div class="hint">base inteira</div></div>
      <div class="kpi warn"><div class="lab">Tempo médio p/ assinar</div>
        <div class="val" id="kConv">—</div><div class="hint">cadastro → assinatura</div></div>
    </div>

    <!-- Seletor de mês -->
    <div class="toolbar">
      <strong style="font-size:14px">Fechamento mensal</strong>
      <select id="selMonth"></select>
      <span class="chip" id="monthHint"></span>
    </div>

    <div class="kpis">
      <div class="kpi accent"><div class="lab">Novos cadastros no mês</div>
        <div class="val" id="mNew">0</div><div class="hint">por DataCadastro</div></div>
      <div class="kpi green"><div class="lab">Novas assinaturas no mês</div>
        <div class="val" id="mSub">0</div><div class="hint">por DataAssinatura</div></div>
      <div class="kpi danger"><div class="lab">Cancelamentos no mês</div>
        <div class="val" id="mCanc">0</div><div class="hint">por DataCancelamento</div></div>
      <div class="kpi warn"><div class="lab">Taxa de ativação</div>
        <div class="val" id="mAct">—</div><div class="hint">cadastrados que fizeram ≥1 laudo</div></div>
    </div>

    <!-- Comportamento dos gratuitos no mês -->
    <div class="grid">
      <div class="card col-5">
        <h3>Comportamento dos novos gratuitos no mês</h3>
        <div class="csub" id="freeMonthSub">—</div>
        <canvas id="chartFreeBehavior"></canvas>
      </div>
      <div class="card col-7">
        <h3>Curva de crescimento — cadastros por mês</h3>
        <div class="csub">Novos usuários a cada mês (por DataCadastro)</div>
        <canvas id="chartGrowth"></canvas>
      </div>
    </div>

    <div class="grid">
      <div class="card col-7">
        <h3>Composição da base por mês</h3>
        <div class="csub">Assinaturas vs. cancelamentos (eventos por mês)</div>
        <canvas id="chartSubCanc"></canvas>
      </div>
      <div class="card col-5">
        <h3>Origem dos usuários (ComoConheceu)</h3>
        <div class="csub">Distribuição de toda a base filtrada</div>
        <canvas id="chartSource"></canvas>
      </div>
    </div>

    <div class="grid">
      <div class="card col-6">
        <h3>Top estados</h3>
        <div class="csub">Usuários por UF</div>
        <canvas id="chartState"></canvas>
      </div>
      <div class="card col-6">
        <h3>Faixas de laudos emitidos</h3>
        <div class="csub">Distribuição de QtdLaudos (base filtrada)</div>
        <canvas id="chartReports"></canvas>
      </div>
    </div>

    <!-- Snapshots -->
    <div class="grid">
      <div class="card col-12">
        <h3>Histórico de snapshots mensais</h3>
        <div class="csub">Fotos agregadas salvas no seu navegador (IndexedDB). Sem dados pessoais — só números.</div>
        <table id="snapTable">
          <thead><tr>
            <th>Data do snapshot</th><th class="num">Assinantes</th><th class="num">Gratuitos</th>
            <th class="num">Ex-assinantes</th><th class="num">Cortesia</th><th class="num">Total</th><th></th>
          </tr></thead>
          <tbody></tbody>
        </table>
        <p id="noSnap" class="csub" style="margin-top:10px">Nenhum snapshot salvo ainda.</p>
      </div>
    </div>

  </div><!-- /app -->

</div>

<div class="foot">
  nReport Dashboard • roda offline no navegador • dados sensíveis nunca saem da sua máquina.
</div>

<script src="app.js"></script>
</body>
</html>
