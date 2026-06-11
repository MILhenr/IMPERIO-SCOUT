<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Scout Stats</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#0a0e1a;--surface:#111827;--card:#1a2235;--border:#2a3650;
  --blue:#3b82f6;--blue2:#2563eb;--green:#22c55e;--red:#ef4444;
  --yellow:#f59e0b;--text:#f1f5f9;--muted:#64748b;--light:#94a3b8;
}
body{background:var(--bg);color:var(--text);font-family:'Inter',system-ui,sans-serif;min-height:100vh}
nav{background:var(--surface);border-bottom:1px solid var(--border);padding:0 20px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.logo{font-size:18px;font-weight:900;letter-spacing:2px;color:var(--blue)}
.nav-right{display:flex;gap:8px;flex-wrap:wrap}
.page{display:none;padding:20px;max-width:960px;margin:0 auto}
.page.active{display:block}
.teams-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-top:20px}
.team-card{background:var(--card);border:1.5px solid var(--border);border-radius:14px;padding:20px;cursor:pointer;transition:all .2s;text-align:center}
.team-card:hover{border-color:var(--blue);transform:translateY(-2px)}
.team-emoji{font-size:36px;margin-bottom:8px}
.team-name{font-size:16px;font-weight:800;text-transform:uppercase;letter-spacing:.5px}
.team-stats{font-size:12px;color:var(--muted);margin-top:6px}
.btn{border:none;border-radius:9px;font-weight:700;cursor:pointer;transition:all .2s;display:inline-flex;align-items:center;gap:7px;white-space:nowrap}
.btn-blue{background:var(--blue);color:#fff;padding:10px 18px;font-size:13px}
.btn-blue:hover{background:var(--blue2)}
.btn-green{background:rgba(34,197,94,.15);border:1.5px solid var(--green);color:var(--green);padding:9px 16px;font-size:13px}
.btn-green:hover{background:var(--green);color:#fff}
.btn-yellow{background:rgba(245,158,11,.15);border:1.5px solid var(--yellow);color:var(--yellow);padding:9px 16px;font-size:13px}
.btn-yellow:hover{background:var(--yellow);color:#fff}
.btn-outline{background:transparent;border:1.5px solid var(--border);color:var(--light);padding:8px 16px;font-size:13px}
.btn-outline:hover{border-color:var(--blue);color:var(--blue)}
.team-header{background:linear-gradient(135deg,var(--blue) 0%,#1d4ed8 100%);border-radius:16px;padding:24px;margin-bottom:20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.team-header-left{display:flex;align-items:center;gap:16px}
.team-header-emoji{font-size:48px}
.team-header-name{font-size:28px;font-weight:900;text-transform:uppercase;letter-spacing:1px}
.team-header-sub{font-size:13px;color:rgba(255,255,255,.6);margin-top:2px}
.btn-jogo{background:rgba(255,255,255,.15);border:2px solid rgba(255,255,255,.3);color:#fff;padding:12px 20px;border-radius:10px;font-size:15px;font-weight:800;cursor:pointer;transition:all .2s;display:flex;align-items:center;gap:8px}
.btn-jogo:hover{background:rgba(255,255,255,.25);transform:scale(1.04)}
.jogos-badge{background:rgba(255,255,255,.2);border-radius:8px;padding:12px 20px;text-align:center}
.jogos-num{font-size:32px;font-weight:900;line-height:1}
.jogos-label{font-size:11px;color:rgba(255,255,255,.6);text-transform:uppercase;letter-spacing:.1em;margin-top:2px}
.section-title{font-size:11px;font-weight:700;letter-spacing:.15em;text-transform:uppercase;color:var(--muted);margin-bottom:12px;display:flex;align-items:center;gap:8px}
.section-title::before{content:'';width:20px;height:2px;background:var(--blue);display:inline-block}
.players-table{width:100%;border-collapse:collapse;background:var(--card);border-radius:12px;overflow:hidden;border:1.5px solid var(--border)}
.players-table th{background:var(--surface);padding:10px 14px;text-align:left;font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);border-bottom:1.5px solid var(--border)}
.players-table td{padding:10px 14px;border-bottom:1px solid var(--border);font-size:14px;vertical-align:middle}
.players-table tr:last-child td{border-bottom:none}
.players-table tr:hover td{background:rgba(59,130,246,.05)}
.stat-val{font-size:18px;font-weight:900;color:var(--blue);font-variant-numeric:tabular-nums;min-width:28px;text-align:center;display:inline-block}
.btn-stat{border:none;border-radius:7px;width:34px;height:34px;font-size:16px;font-weight:900;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;justify-content:center}
.btn-gol{background:rgba(34,197,94,.15);color:var(--green)}
.btn-gol:hover{background:var(--green);color:#fff;transform:scale(1.1)}
.btn-ass{background:rgba(59,130,246,.15);color:var(--blue)}
.btn-ass:hover{background:var(--blue);color:#fff;transform:scale(1.1)}
.btn-minus{background:rgba(239,68,68,.1);color:var(--red);font-size:13px}
.btn-minus:hover{background:var(--red);color:#fff}
.stat-cell{display:flex;align-items:center;gap:8px;justify-content:center}
.add-player-row{display:flex;gap:8px;align-items:center;margin-top:12px;flex-wrap:wrap}
.input-s{background:var(--card);border:1.5px solid var(--border);color:var(--text);padding:10px 12px;border-radius:8px;font-size:14px;outline:none;transition:border-color .2s}
.input-s:focus{border-color:var(--blue)}
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.75);backdrop-filter:blur(8px);display:flex;align-items:center;justify-content:center;z-index:500;opacity:0;pointer-events:none;transition:opacity .2s;padding:20px}
.modal-overlay.open{opacity:1;pointer-events:all}
.modal{background:var(--surface);border:1.5px solid var(--border);border-radius:18px;padding:32px;width:100%;max-width:460px;transform:scale(.95);transition:transform .2s;max-height:90vh;overflow-y:auto}
.modal-overlay.open .modal{transform:scale(1)}
.modal-title{font-size:20px;font-weight:900;text-transform:uppercase;margin-bottom:20px}
.input-full{background:var(--card);border:1.5px solid var(--border);color:var(--text);padding:12px 14px;border-radius:9px;font-size:14px;outline:none;width:100%;margin-bottom:10px}
.input-full:focus{border-color:var(--blue)}
.btn-primary{background:var(--blue);border:none;color:#fff;padding:14px;border-radius:9px;font-size:15px;font-weight:800;cursor:pointer;text-transform:uppercase;letter-spacing:.5px;width:100%;margin-bottom:8px}
.btn-primary:hover{background:var(--blue2)}
.btn-cancel{background:transparent;border:1.5px solid var(--border);color:var(--muted);padding:12px;border-radius:9px;font-size:14px;font-weight:600;cursor:pointer;width:100%}
.toast{position:fixed;bottom:24px;right:24px;background:var(--surface);border:1.5px solid var(--border);color:var(--text);padding:12px 20px;border-radius:10px;font-size:14px;font-weight:600;z-index:9000;transform:translateY(8px);opacity:0;transition:all .3s;pointer-events:none}
.toast.show{transform:translateY(0);opacity:1}
.empty{text-align:center;padding:60px 20px;color:var(--muted)}
.empty-icon{font-size:48px;margin-bottom:12px}
.drop-zone{border:2px dashed var(--border);border-radius:12px;padding:32px;text-align:center;cursor:pointer;transition:border-color .2s;background:var(--card);margin-bottom:12px}
.drop-zone:hover,.drop-zone.drag-over{border-color:var(--blue);background:rgba(59,130,246,.05)}
.drop-zone-icon{font-size:32px;margin-bottom:8px}
.drop-zone-text{font-size:14px;font-weight:600;color:var(--light)}
.drop-zone-sub{font-size:12px;color:var(--muted);margin-top:4px}
.preview-list{max-height:200px;overflow-y:auto;border:1.5px solid var(--border);border-radius:8px;margin-bottom:12px}
.preview-item{padding:8px 14px;border-bottom:1px solid var(--border);font-size:13px;display:flex;gap:12px}
.preview-item:last-child{border-bottom:none}
.tab-btns{display:flex;gap:4px;background:var(--card);border-radius:10px;padding:4px;margin-bottom:20px}
.tab-btn{flex:1;padding:8px;border:none;border-radius:7px;font-size:13px;font-weight:700;cursor:pointer;background:transparent;color:var(--muted);transition:all .2s}
.tab-btn.active{background:var(--blue);color:#fff}
@media(max-width:600px){
  .nav-right{gap:4px}
  .btn{padding:8px 10px;font-size:12px}
  .team-header{flex-direction:column}
}
</style>
</head>
<body>
<div class="toast" id="toast"></div>

<!-- MODAL: NOVO TIME -->
<div class="modal-overlay" id="modalTime">
  <div class="modal">
    <div class="modal-title">⚽ Novo Time</div>
    <input class="input-full" id="new-team-name" placeholder="Nome do time (ex: Marreco)" maxlength="40">
    <input class="input-full" id="new-team-emoji" placeholder="Emoji (ex: 🦅)" maxlength="4">
    <button class="btn-primary" onclick="criarTime()">CRIAR TIME →</button>
    <button class="btn-cancel" onclick="fecharModal('modalTime')">Cancelar</button>
  </div>
</div>

<!-- MODAL: IMPORTAR -->
<div class="modal-overlay" id="modalImport">
  <div class="modal">
    <div class="modal-title">📥 Importar CSV</div>
    <div class="tab-btns">
      <button class="tab-btn active" id="tab-jogadores" onclick="switchTab('jogadores')">Jogadores</button>
      <button class="tab-btn" id="tab-times" onclick="switchTab('times')">Times Completos</button>
    </div>

    <!-- TAB JOGADORES -->
    <div id="import-jogadores">
      <p style="font-size:13px;color:var(--muted);margin-bottom:12px">Modelo: <code style="background:var(--card);padding:2px 6px;border-radius:4px">Nº,Nome</code></p>
      <div class="drop-zone" id="dz-jogadores" onclick="document.getElementById('file-jogadores').click()" ondragover="dzOver(event,'dz-jogadores')" ondragleave="dzLeave('dz-jogadores')" ondrop="dzDrop(event,'jogadores')">
        <div class="drop-zone-icon">📄</div>
        <div class="drop-zone-text">Arraste o CSV aqui ou clique</div>
        <div class="drop-zone-sub">Arquivo .csv</div>
      </div>
      <input type="file" id="file-jogadores" accept=".csv" style="display:none" onchange="handleCSV(this,'jogadores')">
      <div id="preview-jogadores" style="display:none">
        <div style="font-size:13px;font-weight:700;margin-bottom:8px" id="preview-jogadores-title"></div>
        <div class="preview-list" id="preview-jogadores-list"></div>
        <button class="btn-primary" onclick="importarJogadores()">IMPORTAR TODOS →</button>
      </div>
    </div>

    <!-- TAB TIMES -->
    <div id="import-times" style="display:none">
      <p style="font-size:13px;color:var(--muted);margin-bottom:12px">Modelo: <code style="background:var(--card);padding:2px 6px;border-radius:4px">Time,Emoji,Jogos,Nº,Nome,Gols,Passes</code></p>
      <div class="drop-zone" id="dz-times" onclick="document.getElementById('file-times').click()" ondragover="dzOver(event,'dz-times')" ondragleave="dzLeave('dz-times')" ondrop="dzDrop(event,'times')">
        <div class="drop-zone-icon">📊</div>
        <div class="drop-zone-text">Arraste o CSV aqui ou clique</div>
        <div class="drop-zone-sub">Arquivo .csv</div>
      </div>
      <input type="file" id="file-times" accept=".csv" style="display:none" onchange="handleCSV(this,'times')">
      <div id="preview-times" style="display:none">
        <div style="font-size:13px;font-weight:700;margin-bottom:8px" id="preview-times-title"></div>
        <div class="preview-list" id="preview-times-list"></div>
        <button class="btn-primary" onclick="importarTimes()">IMPORTAR TUDO →</button>
      </div>
    </div>

    <button class="btn-cancel" style="margin-top:8px" onclick="fecharModal('modalImport')">Fechar</button>
  </div>
</div>

<!-- NAV -->
<nav>
  <div class="logo">⚽ SCOUT STATS</div>
  <div class="nav-right">
    <button class="btn btn-yellow" onclick="exportarTudo()">📊 Exportar</button>
    <button class="btn btn-green" onclick="abrirImport()">📥 Importar</button>
    <button class="btn btn-blue" onclick="abrirModal('modalTime')">+ Novo Time</button>
  </div>
</nav>

<!-- PAGE: TIMES -->
<div id="page-times" class="page active">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap;gap:8px">
    <div>
      <div style="font-size:22px;font-weight:900;text-transform:uppercase;letter-spacing:1px">Meus Times</div>
      <div style="font-size:13px;color:var(--muted);margin-top:2px">Clique num time para registrar estatísticas</div>
    </div>
  </div>
  <div class="teams-grid" id="teams-grid"></div>
</div>

<!-- PAGE: TIME -->
<div id="page-time" class="page">
  <button class="btn btn-outline" style="margin-bottom:16px" onclick="voltarTimes()">← Todos os Times</button>
  <div class="team-header">
    <div class="team-header-left">
      <div class="team-header-emoji" id="th-emoji">⚽</div>
      <div>
        <div class="team-header-name" id="th-name">Time</div>
        <div class="team-header-sub" id="th-sub">0 jogadores</div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <div class="jogos-badge">
        <div class="jogos-num" id="th-jogos">0</div>
        <div class="jogos-label">Jogos</div>
      </div>
      <button class="btn-jogo" onclick="addJogo()">🏁 +1 Jogo</button>
      <button class="btn-jogo" style="background:rgba(239,68,68,.2);border-color:rgba(239,68,68,.4)" onclick="removeJogo()">−1 Jogo</button>
    </div>
  </div>

  <div class="section-title">Jogadores</div>
  <table class="players-table">
    <thead>
      <tr>
        <th style="width:44px">#</th>
        <th>Jogador</th>
        <th style="width:110px;text-align:center">⚽ Gols</th>
        <th style="width:110px;text-align:center">🅰️ Passes</th>
        <th style="width:40px"></th>
      </tr>
    </thead>
    <tbody id="players-tbody"></tbody>
  </table>

  <div class="add-player-row">
    <input class="input-s" id="new-num" placeholder="Nº" type="number" min="1" max="99" style="width:64px">
    <input class="input-s" id="new-nome" placeholder="Nome do jogador" style="flex:1;min-width:160px">
    <button class="btn btn-blue" onclick="addJogador()">+ Adicionar</button>
  </div>
</div>

<script>
var times = JSON.parse(localStorage.getItem('scout_times')||'[]');
var currentIdx = null;
var csvJogadoresParsed = [], csvTimesParsed = [];

function salvar(){ localStorage.setItem('scout_times', JSON.stringify(times)); }

function toast(msg){
  var t=document.getElementById('toast');
  t.textContent=msg; t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); },2500);
}

function showPage(p){
  document.querySelectorAll('.page').forEach(function(el){ el.classList.remove('active'); });
  document.getElementById('page-'+p).classList.add('active');
}

function abrirModal(id){ document.getElementById(id).classList.add('open'); }
function fecharModal(id){ document.getElementById(id).classList.remove('open'); }
document.querySelectorAll('.modal-overlay').forEach(function(el){
  el.addEventListener('click',function(e){ if(e.target===this) this.classList.remove('open'); });
});

// ── TIMES ────────────────────────────────────────────────────────────────────
function criarTime(){
  var nome=document.getElementById('new-team-name').value.trim();
  var emoji=document.getElementById('new-team-emoji').value.trim()||'⚽';
  if(!nome){ toast('⚠️ Digite o nome'); return; }
  times.push({nome:nome,emoji:emoji,jogos:0,jogadores:[]});
  salvar(); fecharModal('modalTime');
  document.getElementById('new-team-name').value='';
  document.getElementById('new-team-emoji').value='';
  renderTimes(); toast('✅ '+nome+' criado!');
}

function renderTimes(){
  var grid=document.getElementById('teams-grid');
  if(!times.length){
    grid.innerHTML='<div class="empty"><div class="empty-icon">🏟</div><div>Nenhum time ainda.<br>Clique em "+ Novo Time" para começar!</div></div>';
    return;
  }
  grid.innerHTML=times.map(function(t,i){
    var gols=t.jogadores.reduce(function(s,j){return s+(j.gols||0);},0);
    var passes=t.jogadores.reduce(function(s,j){return s+(j.passes||0);},0);
    return '<div class="team-card" onclick="abrirTime('+i+')">'
      +'<div class="team-emoji">'+t.emoji+'</div>'
      +'<div class="team-name">'+t.nome+'</div>'
      +'<div class="team-stats">'+t.jogos+' jogos · '+gols+' gols · '+passes+' passes</div>'
      +'</div>';
  }).join('');
}

function abrirTime(idx){
  currentIdx=idx;
  var t=times[idx];
  document.getElementById('th-emoji').textContent=t.emoji;
  document.getElementById('th-name').textContent=t.nome;
  document.getElementById('th-jogos').textContent=t.jogos;
  renderJogadores();
  showPage('time');
}

function voltarTimes(){ currentIdx=null; showPage('times'); renderTimes(); }

function addJogo(){
  var t=times[currentIdx]; t.jogos++;
  document.getElementById('th-jogos').textContent=t.jogos;
  salvar(); toast('🏁 +1 jogo — '+t.nome+' ('+t.jogos+' total)');
}
function removeJogo(){
  var t=times[currentIdx];
  if(t.jogos<=0){ toast('⚠️ Já está em 0'); return; }
  t.jogos--;
  document.getElementById('th-jogos').textContent=t.jogos;
  salvar(); toast('↩️ -1 jogo ('+t.jogos+' total)');
}

// ── JOGADORES ────────────────────────────────────────────────────────────────
function addJogador(){
  var num=document.getElementById('new-num').value.trim();
  var nome=document.getElementById('new-nome').value.trim();
  if(!nome){ toast('⚠️ Digite o nome'); return; }
  var t=times[currentIdx];
  if(num && t.jogadores.find(function(j){ return String(j.num)===String(num); })){
    toast('⚠️ Número '+num+' já existe'); return;
  }
  t.jogadores.push({num:num||'—',nome:nome,gols:0,passes:0});
  t.jogadores.sort(function(a,b){ return parseInt(a.num)-parseInt(b.num); });
  salvar();
  document.getElementById('new-num').value='';
  document.getElementById('new-nome').value='';
  renderJogadores(); toast('✅ '+nome+' adicionado!');
}

function addGol(i){ var t=times[currentIdx]; t.jogadores[i].gols++; salvar(); renderJogadores(); toast('⚽ +1 gol — '+t.jogadores[i].nome); }
function removeGol(i){ var t=times[currentIdx]; if(t.jogadores[i].gols<=0) return; t.jogadores[i].gols--; salvar(); renderJogadores(); toast('↩️ -1 gol'); }
function addPasse(i){ var t=times[currentIdx]; t.jogadores[i].passes++; salvar(); renderJogadores(); toast('🅰️ +1 passe — '+t.jogadores[i].nome); }
function removePasse(i){ var t=times[currentIdx]; if(t.jogadores[i].passes<=0) return; t.jogadores[i].passes--; salvar(); renderJogadores(); toast('↩️ -1 passe'); }
function removerJogador(i){
  var t=times[currentIdx];
  if(!confirm('Remover '+t.jogadores[i].nome+'?')) return;
  var nome=t.jogadores[i].nome;
  t.jogadores.splice(i,1); salvar(); renderJogadores(); toast('🗑 '+nome+' removido');
}

function renderJogadores(){
  var t=times[currentIdx];
  var tbody=document.getElementById('players-tbody');
  document.getElementById('th-sub').textContent=t.jogadores.length+' jogadores';
  document.getElementById('th-jogos').textContent=t.jogos;
  if(!t.jogadores.length){
    tbody.innerHTML='<tr><td colspan="5" style="text-align:center;padding:40px;color:var(--muted)">Nenhum jogador. Adicione abaixo!</td></tr>';
    return;
  }
  tbody.innerHTML=t.jogadores.map(function(j,i){
    return '<tr>'
      +'<td><span style="font-size:13px;font-weight:700;color:var(--muted)">'+j.num+'</span></td>'
      +'<td style="font-weight:700">'+j.nome+'</td>'
      +'<td><div class="stat-cell">'
      +'<button class="btn-stat btn-minus" onclick="removeGol('+i+')">−</button>'
      +'<span class="stat-val">'+(j.gols||0)+'</span>'
      +'<button class="btn-stat btn-gol" onclick="addGol('+i+')">⚽</button>'
      +'</div></td>'
      +'<td><div class="stat-cell">'
      +'<button class="btn-stat btn-minus" onclick="removePasse('+i+')">−</button>'
      +'<span class="stat-val">'+(j.passes||0)+'</span>'
      +'<button class="btn-stat btn-ass" onclick="addPasse('+i+')">🅰️</button>'
      +'</div></td>'
      +'<td><button onclick="removerJogador('+i+')" style="background:rgba(239,68,68,.1);border:none;color:var(--red);border-radius:6px;width:28px;height:28px;cursor:pointer;font-size:13px">🗑</button></td>'
      +'</tr>';
  }).join('');
}

// ── EXPORTAR XLSX ─────────────────────────────────────────────────────────────
function exportarTudo(){
  if(!times.length){ toast('⚠️ Nenhum time para exportar'); return; }
  var wb = XLSX.utils.book_new();

  var hFill = {fgColor:{rgb:'1F4E79'}};
  var hFont = {bold:true,color:{rgb:'FFFFFF'},name:'Arial',sz:11};
  var hAlign = {horizontal:'center',vertical:'center'};
  var border = {top:{style:'thin',color:{rgb:'CCCCCC'}},bottom:{style:'thin',color:{rgb:'CCCCCC'}},left:{style:'thin',color:{rgb:'CCCCCC'}},right:{style:'thin',color:{rgb:'CCCCCC'}}};
  var evenFill = {fgColor:{rgb:'EBF5FB'}};

  function makeCell(v, bold, fillRgb, color){
    var c = {v:v, t:typeof v==='number'?'n':'s'};
    var s = {border:border, alignment:{horizontal:'center',vertical:'center'}, font:{name:'Arial',sz:10}};
    if(bold) s.font.bold = true;
    if(fillRgb) s.fill = {fgColor:{rgb:fillRgb}};
    if(color) s.font.color = {rgb:color};
    c.s = s;
    return c;
  }

  // Aba por time
  times.forEach(function(t){
    var data = [];
    // Header
    var hr = ['Nº','Jogador','Gols','Passes','Jogos do Time'];
    var hRow = hr.map(function(h){ return {v:h,t:'s',s:{fill:hFill,font:hFont,alignment:hAlign,border:border}}; });
    data.push(hRow);
    // Jogadores
    t.jogadores.forEach(function(j,i){
      var fill = i%2===0 ? null : 'EBF5FB';
      data.push([
        makeCell(j.num||'—',false,fill),
        makeCell(j.nome,true,fill),
        makeCell(j.gols||0,false,fill,'1F78D1'),
        makeCell(j.passes||0,false,fill,'1F78D1'),
        i===0 ? makeCell(t.jogos,true,fill,'15803D') : makeCell('',false,fill),
      ]);
    });
    if(!t.jogadores.length){
      data.push([makeCell('—'),makeCell('Sem jogadores'),makeCell(0),makeCell(0),makeCell(t.jogos)]);
    }
    var ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{wch:6},{wch:22},{wch:8},{wch:8},{wch:14}];
    ws['!rows'] = [{hpt:22}];
    var nome = t.nome.substring(0,31).replace(/[:\\\/?*\[\]]/g,'');
    XLSX.utils.book_append_sheet(wb, ws, nome);
  });

  // Aba Artilharia
  var artData = [];
  artData.push(['Pos','Jogador','Time','Gols','Passes'].map(function(h){
    return {v:h,t:'s',s:{fill:{fgColor:{rgb:'7F1D1D'}},font:{bold:true,color:{rgb:'FFFFFF'},name:'Arial',sz:11},alignment:hAlign,border:border}};
  }));
  var todos = [];
  times.forEach(function(t){
    t.jogadores.forEach(function(j){
      if((j.gols||0)>0) todos.push({nome:j.nome,time:t.nome,gols:j.gols||0,passes:j.passes||0});
    });
  });
  todos.sort(function(a,b){ return b.gols-a.gols; });
  todos.forEach(function(j,i){
    var fill = i%2===0 ? null : 'FEF2F2';
    var pos = i===0?'🥇':i===1?'🥈':i===2?'🥉':(i+1)+'º';
    artData.push([makeCell(pos,i<3,fill),makeCell(j.nome,true,fill),makeCell(j.time,false,fill),makeCell(j.gols,true,fill,'991B1B'),makeCell(j.passes,false,fill)]);
  });
  if(!todos.length) artData.push([makeCell('—'),makeCell('Nenhum gol registrado'),makeCell(''),makeCell(0),makeCell(0)]);
  var wsArt = XLSX.utils.aoa_to_sheet(artData);
  wsArt['!cols'] = [{wch:6},{wch:22},{wch:18},{wch:8},{wch:8}];
  wsArt['!rows'] = [{hpt:22}];
  XLSX.utils.book_append_sheet(wb, wsArt, '🏆 Artilharia');

  // Aba Assistências
  var assData = [];
  assData.push(['Pos','Jogador','Time','Passes','Gols'].map(function(h){
    return {v:h,t:'s',s:{fill:{fgColor:{rgb:'1A4F2D'}},font:{bold:true,color:{rgb:'FFFFFF'},name:'Arial',sz:11},alignment:hAlign,border:border}};
  }));
  var todos2 = [];
  times.forEach(function(t){
    t.jogadores.forEach(function(j){
      if((j.passes||0)>0) todos2.push({nome:j.nome,time:t.nome,gols:j.gols||0,passes:j.passes||0});
    });
  });
  todos2.sort(function(a,b){ return b.passes-a.passes; });
  todos2.forEach(function(j,i){
    var fill = i%2===0 ? null : 'F0FDF4';
    var pos = i===0?'🥇':i===1?'🥈':i===2?'🥉':(i+1)+'º';
    assData.push([makeCell(pos,i<3,fill),makeCell(j.nome,true,fill),makeCell(j.time,false,fill),makeCell(j.passes,true,fill,'15803D'),makeCell(j.gols,false,fill)]);
  });
  if(!todos2.length) assData.push([makeCell('—'),makeCell('Nenhum passe registrado'),makeCell(''),makeCell(0),makeCell(0)]);
  var wsAss = XLSX.utils.aoa_to_sheet(assData);
  wsAss['!cols'] = [{wch:6},{wch:22},{wch:18},{wch:8},{wch:8}];
  wsAss['!rows'] = [{hpt:22}];
  XLSX.utils.book_append_sheet(wb, wsAss, '🅰️ Assistências');

  // Aba Resumo Times
  var timesData = [];
  timesData.push(['Time','Jogos','Gols Marcados','Total Passes'].map(function(h){
    return {v:h,t:'s',s:{fill:{fgColor:{rgb:'1F3D5C'}},font:{bold:true,color:{rgb:'FFFFFF'},name:'Arial',sz:11},alignment:hAlign,border:border}};
  }));
  times.forEach(function(t,i){
    var fill = i%2===0 ? null : 'E8F4FD';
    var gols=t.jogadores.reduce(function(s,j){return s+(j.gols||0);},0);
    var passes=t.jogadores.reduce(function(s,j){return s+(j.passes||0);},0);
    timesData.push([makeCell(t.emoji+' '+t.nome,true,fill),makeCell(t.jogos,false,fill),makeCell(gols,true,fill,'1F78D1'),makeCell(passes,false,fill)]);
  });
  var wsT = XLSX.utils.aoa_to_sheet(timesData);
  wsT['!cols'] = [{wch:22},{wch:8},{wch:16},{wch:14}];
  wsT['!rows'] = [{hpt:22}];
  XLSX.utils.book_append_sheet(wb, wsT, '📊 Resumo Times');

  XLSX.writeFile(wb, 'scout_stats.xlsx');
  toast('📊 Excel exportado com sucesso!');
}

// ── IMPORTAR CSV ─────────────────────────────────────────────────────────────
function abrirImport(){
  csvJogadoresParsed=[]; csvTimesParsed=[];
  document.getElementById('preview-jogadores').style.display='none';
  document.getElementById('preview-times').style.display='none';
  document.getElementById('file-jogadores').value='';
  document.getElementById('file-times').value='';
  switchTab('jogadores');
  abrirModal('modalImport');
}

function switchTab(tab){
  document.getElementById('import-jogadores').style.display=tab==='jogadores'?'block':'none';
  document.getElementById('import-times').style.display=tab==='times'?'block':'none';
  document.getElementById('tab-jogadores').classList.toggle('active',tab==='jogadores');
  document.getElementById('tab-times').classList.toggle('active',tab==='times');
}

function dzOver(e,id){ e.preventDefault(); document.getElementById(id).classList.add('drag-over'); }
function dzLeave(id){ document.getElementById(id).classList.remove('drag-over'); }
function dzDrop(e,tipo){ e.preventDefault(); dzLeave('dz-'+tipo); var f=e.dataTransfer.files[0]; if(f) processCSVFile(f,tipo); }
function handleCSV(input,tipo){ var f=input.files[0]; if(f) processCSVFile(f,tipo); }

function parseCSV(text){
  var lines=text.trim().split('\n').filter(function(l){ return l.trim(); });
  var headers=lines[0].split(',').map(function(h){ return h.trim().toLowerCase().replace(/[^\w]/g,''); });
  return lines.slice(1).map(function(line){
    var vals=line.split(',').map(function(v){ return v.trim().replace(/^"|"$/g,''); });
    var obj={};
    headers.forEach(function(h,i){ obj[h]=vals[i]||''; });
    return obj;
  });
}

function processCSVFile(file,tipo){
  var reader=new FileReader();
  reader.onload=function(e){
    var rows=parseCSV(e.target.result);
    if(tipo==='jogadores'){
      csvJogadoresParsed=rows.filter(function(r){ return r.nome||r.jogador; });
      var list=document.getElementById('preview-jogadores-list');
      document.getElementById('preview-jogadores-title').textContent=csvJogadoresParsed.length+' jogador(es) encontrado(s)';
      list.innerHTML=csvJogadoresParsed.map(function(r){
        return '<div class="preview-item"><strong>'+(r.n||r.numero||r['']||'—')+'</strong><span>'+(r.nome||r.jogador)+'</span></div>';
      }).join('');
      document.getElementById('preview-jogadores').style.display='block';
    } else {
      csvTimesParsed=rows;
      var timesMap={};
      rows.forEach(function(r){
        var tn=r.time||r.nome;
        if(!tn) return;
        if(!timesMap[tn]) timesMap[tn]={nome:tn,emoji:r.emoji||'⚽',jogos:parseInt(r.jogos)||0,jogadores:[]};
        if(r.jogador||r.nome_jogador||r.jogadores){
          var nj=r.jogador||r.nome_jogador||r.jogadores;
          if(nj) timesMap[tn].jogadores.push({num:r.n||r.numero||'—',nome:nj,gols:parseInt(r.gols)||0,passes:parseInt(r.passes)||0});
        }
      });
      var tArr=Object.values(timesMap);
      document.getElementById('preview-times-title').textContent=tArr.length+' time(s) encontrado(s)';
      document.getElementById('preview-times-list').innerHTML=tArr.map(function(t){
        return '<div class="preview-item"><strong>'+t.emoji+' '+t.nome+'</strong><span>'+t.jogadores.length+' jogadores · '+t.jogos+' jogos</span></div>';
      }).join('');
      document.getElementById('preview-times').style.display='block';
    }
  };
  reader.readAsText(file);
}

function importarJogadores(){
  if(currentIdx===null){ toast('⚠️ Abra um time primeiro'); fecharModal('modalImport'); return; }
  if(!csvJogadoresParsed.length){ toast('⚠️ Nenhum jogador no CSV'); return; }
  var t=times[currentIdx]; var count=0;
  csvJogadoresParsed.forEach(function(r){
    var nome=r.nome||r.jogador; if(!nome) return;
    var num=r.n||r.numero||r['']||'—';
    if(t.jogadores.find(function(j){ return j.nome===nome; })) return;
    t.jogadores.push({num:num,nome:nome,gols:parseInt(r.gols)||0,passes:parseInt(r.passes)||0});
    count++;
  });
  t.jogadores.sort(function(a,b){ return parseInt(a.num)-parseInt(b.num); });
  salvar(); renderJogadores();
  fecharModal('modalImport');
  toast('✅ '+count+' jogador(es) importado(s)!');
}

function importarTimes(){
  if(!csvTimesParsed.length){ toast('⚠️ CSV vazio'); return; }
  var timesMap={};
  csvTimesParsed.forEach(function(r){
    var tn=r.time||r.nome; if(!tn) return;
    if(!timesMap[tn]) timesMap[tn]={nome:tn,emoji:r.emoji||'⚽',jogos:parseInt(r.jogos)||0,jogadores:[]};
    var nj=r.jogador||r.nome_jogador||r.jogadores;
    if(nj) timesMap[tn].jogadores.push({num:r.n||r.numero||'—',nome:nj,gols:parseInt(r.gols)||0,passes:parseInt(r.passes)||0});
  });
  var count=0;
  Object.values(timesMap).forEach(function(t){
    if(!times.find(function(x){ return x.nome===t.nome; })){ times.push(t); count++; }
  });
  salvar(); renderTimes();
  fecharModal('modalImport');
  toast('✅ '+count+' time(s) importado(s)!');
}

// ── ENTER KEYS ────────────────────────────────────────────────────────────────
document.getElementById('new-nome').addEventListener('keyup',function(e){ if(e.key==='Enter') addJogador(); });
document.getElementById('new-team-name').addEventListener('keyup',function(e){ if(e.key==='Enter') criarTime(); });

// ── INIT ─────────────────────────────────────────────────────────────────────
renderTimes();
</script>
</body>
</html>
