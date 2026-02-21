// ============================================================================
// DASHBOARD V21.0 - app.js (VERSÃO FUNCIONAL)
// ============================================================================

const API_URL = 'https://script.google.com/macros/s/AKfycby5ffZWf5lrHveg3SZwqkX6U5e0d87NkNufTWR9vZzzTAq0r7kIheKF5CT1QgiNzXUQHA/exec';

// ============================================================================
// CORES E CONSTANTES
// ============================================================================

const C = {
    green:'#00a651', greenFade:'rgba(0,166,81,0.12)',
    lime:'#7ed321',  teal:'#0d9488',
    warn:'#f59e0b',  warnFade:'rgba(245,158,11,0.12)',
    danger:'#ef4444', gray:'#6b8072', border:'#d1e8d9',
};

const RANKING_SETORES = ['VENDAS','RECEPCAO','REFILIACAO'];
const SEM_FILTRO      = ['recorrencia','recorrencia_vendedor'];

// ============================================================================
// CACHE
// ============================================================================

const CACHE_MEM  = new Map();
const CACHE_TTL  = 30 * 60 * 1000;
const LS_PREFIX  = 'cdt_dash_';

function cacheKey(ep,m,y){ return SEM_FILTRO.includes(ep)?ep:`${ep}:${m}:${y}`; }

function getCached(k){
    const mem = CACHE_MEM.get(k);
    if(mem && Date.now()-mem.ts < CACHE_TTL) return mem.data;
    try{
        const raw = localStorage.getItem(LS_PREFIX+k);
        if(raw){
            const entry = JSON.parse(raw);
            if(Date.now()-entry.ts < CACHE_TTL){
                CACHE_MEM.set(k,{data:entry.data,ts:entry.ts});
                return entry.data;
            }
            localStorage.removeItem(LS_PREFIX+k);
        }
    }catch(_){}
    return null;
}

function setCache(k,d){
    const entry = {data:d, ts:Date.now()};
    CACHE_MEM.set(k, entry);
    try{ localStorage.setItem(LS_PREFIX+k, JSON.stringify(entry)); }catch(_){}
}

function invalidateCache(k){
    CACHE_MEM.delete(k);
    try{ localStorage.removeItem(LS_PREFIX+k); }catch(_){}
}

// ============================================================================
// FUNÇÕES DE DADOS
// ============================================================================

function buildUrl(ep,m,y){ 
    let u=`${API_URL}?endpoint=${ep}`; 
    if(!SEM_FILTRO.includes(ep)) u+=`&mes=${m}&ano=${y}`; 
    return u; 
}

async function fetchData(endpoint, mes, ano) {
    const url = buildUrl(endpoint, mes, ano);
    console.log('Buscando dados de:', url);
    
    try {
        const resp = await fetch(url);
        const json = await resp.json();
        console.log('Resposta:', json);
        return json;
    } catch (error) {
        console.error('Erro na requisição:', error);
        return { status: 'error', error: error.message };
    }
}

// ============================================================================
// VARIÁVEIS GLOBAIS
// ============================================================================

let currentDashboard = 'resumo';
let currentMonth     = new Date().getMonth() + 1;
let currentYear      = new Date().getFullYear();
const chartInstances = {};

let dashboardBtns, monthSelect, yearSelect, dashboardContent,
    refreshBtn, downloadBtn, loadingEl, lastUpdateEl, periodSelector;

// ============================================================================
// INICIALIZAÇÃO
// ============================================================================

document.addEventListener('DOMContentLoaded', function(){
    Chart.defaults.font.family="'Plus Jakarta Sans', sans-serif";
    Chart.defaults.font.size=12; 
    Chart.defaults.color='#5a7a65';

    dashboardBtns    = document.querySelectorAll('.dashboard-btn');
    monthSelect      = document.getElementById('monthSelect');
    yearSelect       = document.getElementById('yearSelect');
    dashboardContent = document.getElementById('dashboardContent');
    refreshBtn       = document.getElementById('refreshBtn');
    downloadBtn      = document.getElementById('downloadBtn');
    loadingEl        = document.getElementById('loading');
    lastUpdateEl     = document.getElementById('lastUpdate');
    periodSelector   = document.getElementById('periodSelector');

    // Popular anos
    const anoAtual = new Date().getFullYear();
    for(let y = anoAtual; y >= anoAtual - 3; y--){
        const opt = document.createElement('option');
        opt.value = y; 
        opt.textContent = y;
        yearSelect.appendChild(opt);
    }

    monthSelect.value = currentMonth;
    yearSelect.value  = currentYear;

    // Event listeners
    dashboardBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            dashboardBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentDashboard = btn.dataset.dashboard;
            periodSelector.style.display = SEM_FILTRO.includes(currentDashboard) ? 'none' : 'flex';
            loadDashboard();
        });
    });

    monthSelect.addEventListener('change', () => { 
        currentMonth = parseInt(monthSelect.value); 
        loadDashboard(); 
    });
    
    yearSelect.addEventListener('change',  () => { 
        currentYear  = parseInt(yearSelect.value);  
        loadDashboard(); 
    });
    
    refreshBtn.addEventListener('click',   () => { 
        invalidateCache(cacheKey(currentDashboard, currentMonth, currentYear)); 
        loadDashboard(); 
    });
    
    downloadBtn.addEventListener('click',  exportPage);

    // Carregar dashboard inicial
    loadDashboard();
});

// ============================================================================
// LOAD DASHBOARD
// ============================================================================

async function loadDashboard(){
    if(currentDashboard === 'resumo'){ 
        await loadResumoDashboard(); 
        return; 
    }

    const k = cacheKey(currentDashboard, currentMonth, currentYear);
    const cached = getCached(k);

    if(cached){
        window._dashboardData = cached;
        renderDashboard();
        updateLastUpdateTime();
        return;
    }

    showProgressLoading('Buscando dados...', 10);
    
    try{
        updateProgress(40, 'Conectando...');
        const result = await fetchData(currentDashboard, currentMonth, currentYear);
        updateProgress(90, 'Renderizando...');
        
        if(result.status === 'success'){
            setCache(k, result.data);
            window._dashboardData = result.data;
            renderDashboard();
        } else {
            showError(result.error || 'Erro ao carregar dados');
        }
    } catch(err){
        showError('Erro de conexão: ' + err.message);
    }
    
    updateLastUpdateTime();
}

// ============================================================================
// RENDERIZAÇÃO DOS DASHBOARDS
// ============================================================================

function renderDashboard(){
    hideLoading();
    const data = window._dashboardData;
    
    if (!data) {
        showError('Dados não encontrados');
        return;
    }
    
    switch(currentDashboard){
        case 'documentacao':
            renderDocumentacaoDashboard(data);
            break;
        case 'app':
            renderAppDashboard(data);
            break;
        case 'adimplencia':
            renderAdimplenciaDashboard(data);
            break;
        case 'recorrencia':
            renderRecorrenciaDashboard(data);
            break;
        case 'recorrencia_vendedor':
            renderRecorrenciaVendedorDashboard(data);
            break;
        case 'refuturiza':
            renderRefuturizaDashboard(data);
            break;
        default:
            dashboardContent.innerHTML = `
                <h2>Dashboard ${currentDashboard}</h2>
                <pre>${JSON.stringify(data, null, 2)}</pre>
            `;
    }
}

// ============================================================================
// DASHBOARD: VENDAS (DOCUMENTAÇÃO)
// ============================================================================

function renderDocumentacaoDashboard(d){
    if (!d || !d.geral) {
        showError('Dados de Vendas não disponíveis');
        return;
    }
    
    const {geral, consultores, mes, ano} = d;
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-folder"></i> Dashboard de Vendas — ${mes} ${ano}</h2>
        <div class="main-cards">
            <div class="card card-doc">
                <div class="card-header">
                    <div class="card-title">Total de Vendas</div>
                    <div class="card-icon"><i class="fas fa-chart-bar"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItem('Total', geral.total || 0)}
                    ${metricItem('Aprovados', geral.aprovados || 0, C.green)}
                    ${metricItem('Pendências', geral.pendencias || 0, C.warn)}
                    ${metricPercent('% Aprovados', calcPercent(geral.aprovados, geral.total))}
                </div>
            </div>
        </div>
    `;
    
    if (consultores && consultores.length > 0) {
        html += `<h3 class="section-title">Consultores</h3><div class="consultant-grid">`;
        consultores.forEach(c => {
            const pct = calcPercent(c.aprovados, c.total);
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome || 'Sem nome'}</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Vendas', c.total || 0)}
                        ${metricItem('Aprovados', c.aprovados || 0, C.green)}
                        ${metricPercent('%', pct)}
                    </div>
                </div>
            `;
        });
        html += `</div>`;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: APP
// ============================================================================

function renderAppDashboard(d){
    if (!d || !d.geral) {
        showError('Dados do App não disponíveis');
        return;
    }
    
    const {geral, consultores, mes, ano} = d;
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-mobile-alt"></i> Dashboard App — ${mes} ${ano}</h2>
        <div class="main-cards">
            <div class="card card-app">
                <div class="card-header">
                    <div class="card-title">App - Total Geral</div>
                    <div class="card-icon"><i class="fas fa-chart-pie"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItem('Total', geral.total || 0)}
                    ${metricItem('Com App', geral.sim || 0, C.green)}
                    ${metricItem('Sem App', geral.nao || 0, C.danger)}
                    ${metricPercent('% Com App', calcPercent(geral.sim, geral.total))}
                </div>
            </div>
        </div>
    `;
    
    if (consultores && consultores.length > 0) {
        html += `<h3 class="section-title">Consultores</h3><div class="consultant-grid">`;
        consultores.forEach(c => {
            const pct = calcPercent(c.sim, c.total);
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome || 'Sem nome'}</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Total', c.total || 0)}
                        ${metricItem('Com App', c.sim || 0, C.green)}
                        ${metricPercent('%', pct)}
                    </div>
                </div>
            `;
        });
        html += `</div>`;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: ADIMPLÊNCIA
// ============================================================================

function renderAdimplenciaDashboard(d){
    if (!d || !d.geral) {
        showError('Dados de Adimplência não disponíveis');
        return;
    }
    
    const {geral, consultores, mes, ano} = d;
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-credit-card"></i> Dashboard Adimplência — ${mes} ${ano}</h2>
        <div class="main-cards">
            <div class="card card-adim">
                <div class="card-header">
                    <div class="card-title">Adimplência - Total</div>
                    <div class="card-icon"><i class="fas fa-chart-line"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItem('Total Trocas', geral.totalTrocas || 0)}
                    ${metricItem('Mens. OK', geral.mensOk || 0, C.green)}
                    ${metricItem('Mens. Aberto', geral.mensAberto || 0, C.warn)}
                    ${metricItem('Mens. Atraso', geral.mensAtraso || 0, C.danger)}
                    ${metricPercent('% Aprovados', geral.percentualAprovado || 0)}
                </div>
            </div>
        </div>
    `;
    
    if (consultores && consultores.length > 0) {
        html += `<h3 class="section-title">Consultores</h3><div class="table-wrapper"><table class="data-table"><thead><tr><th>Consultor</th><th>Trocas</th><th>Mens. OK</th><th>%</th></tr></thead><tbody>`;
        consultores.forEach(c => {
            html += `
                <tr>
                    <td><strong>${c.nome || 'Sem nome'}</strong></td>
                    <td>${c.totalTrocas || 0}</td>
                    <td style="color:${C.green}">${c.mensOk || 0}</td>
                    <td><span class="metric-percent ${getPercentClass(c.percentualAprovado)}">${c.percentualAprovado || 0}%</span></td>
                </tr>
            `;
        });
        html += `</tbody></table></div>`;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: RECORRÊNCIA
// ============================================================================

function renderRecorrenciaDashboard(d){
    dashboardContent.innerHTML = `
        <h2 class="dash-title"><i class="fas fa-redo"></i> Dashboard Recorrência</h2>
        <pre>${JSON.stringify(d, null, 2)}</pre>
    `;
}

// ============================================================================
// DASHBOARD: REC. VENDEDOR
// ============================================================================

function renderRecorrenciaVendedorDashboard(d){
    if (!d || !d.geral) {
        showError('Dados de Recorrência Vendedor não disponíveis');
        return;
    }
    
    const {geral, consultores, totalConsultores, totalRegistros} = d;
    const pG = calcPercent(geral.totalOk + geral.totalEmAberto, geral.totalVendasPromocao);
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-handshake"></i> Recorrência Vendedor — Dados Completos</h2>
        <div style="background:#e0f2f1;padding:14px 18px;border-radius:12px;margin-bottom:24px;">
            <p><i class="fas fa-database"></i> <strong>${totalRegistros || 0}</strong> registros | <strong>${totalConsultores}</strong> consultores</p>
        </div>
        <div class="main-cards">
            <div class="card">
                <div class="card-header">
                    <div class="card-title">Total Geral</div>
                    <div class="card-icon" style="background:#e0f2f1;color:${C.teal}"><i class="fas fa-chart-bar"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItem('Total Vendas', geral.totalVendasPromocao || 0)}
                    ${metricItem('OK', geral.totalOk || 0, C.green)}
                    ${metricItem('Em Aberto', geral.totalEmAberto || 0, C.warn)}
                    ${metricPercent('% OK+Aberto', pG)}
                </div>
            </div>
        </div>
    `;
    
    if (consultores && consultores.length > 0) {
        html += `<h3 class="section-title">Consultores</h3><div class="consultant-grid">`;
        consultores.slice(0, 10).forEach(c => {
            const p = calcPercent((c.totalOk||0) + (c.totalEmAberto||0), c.totalVendasPromocao);
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome || 'Sem nome'}</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Vendas', c.totalVendasPromocao || 0)}
                        ${metricItem('OK', c.totalOk || 0, C.green)}
                        ${metricItem('Aberto', c.totalEmAberto || 0, C.warn)}
                        ${metricPercent('%', p)}
                    </div>
                </div>
            `;
        });
        html += `</div>`;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: REFUTURIZA
// ============================================================================

function renderRefuturizaDashboard(d){
    if (!d || !d.geral) {
        showError('Dados do Refuturiza não disponíveis');
        return;
    }
    
    const {geral, consultores, mes, ano} = d;
    const pG = calcPercent(geral.comLigacao, geral.total);
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-book"></i> Dashboard Refuturiza — ${mes} ${ano}</h2>
        <div class="main-cards">
            <div class="card card-refut">
                <div class="card-header">
                    <div class="card-title">Refuturiza - Total</div>
                    <div class="card-icon"><i class="fas fa-book-open"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItem('Total', geral.total || 0)}
                    ${metricItem('Com Ligação', geral.comLigacao || 0, C.green)}
                    ${metricItem('Sem Ligação', geral.semLigacao || 0, C.danger)}
                    ${metricPercent('% Com Ligação', pG)}
                </div>
            </div>
        </div>
    `;
    
    if (consultores && consultores.length > 0) {
        html += `<h3 class="section-title">Consultores</h3><div class="consultant-grid">`;
        consultores.forEach(c => {
            const p = calcPercent(c.comLigacao, c.total);
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome || 'Sem nome'}</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Total', c.total || 0)}
                        ${metricItem('Com Ligação', c.comLigacao || 0, C.green)}
                        ${metricPercent('%', p)}
                    </div>
                </div>
            `;
        });
        html += `</div>`;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: RESUMO
// ============================================================================

async function loadResumoDashboard(){
    dashboardContent.innerHTML = '<h2>Carregando Resumo...</h2>';
    
    // Carregar dados dos 3 principais dashboards
    const endpoints = ['documentacao', 'app', 'adimplencia'];
    const results = {};
    
    for (const ep of endpoints) {
        try {
            const result = await fetchData(ep, currentMonth, currentYear);
            if (result.status === 'success') {
                results[ep] = result.data;
            }
        } catch (e) {
            console.warn(`Erro ao carregar ${ep}:`, e);
        }
    }
    
    if (results.documentacao) {
        const v = results.documentacao;
        const pAprov = calcPercent(v.geral?.aprovados, v.geral?.total);
        
        dashboardContent.innerHTML = `
            <h2 class="dash-title">Resumo Geral — ${v.mes} ${v.ano}</h2>
            <div class="kpi-strip">
                ${kpiCard('Total de Vendas', v.geral?.total || 0, 'fas fa-shopping-bag', C.green)}
                ${kpiCard('% Aprovação', pAprov + '%', 'fas fa-check-circle', pAprov>=90?C.green:pAprov>=80?C.warn:C.danger)}
            </div>
            <div class="main-cards">
                ${results.app ? `
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">App</div>
                            <div class="card-icon" style="background:${C.teal}20;color:${C.teal}"><i class="fas fa-mobile-alt"></i></div>
                        </div>
                        <div class="metric-grid">
                            ${metricItem('Total', results.app.geral?.total || 0)}
                            ${metricItem('Com App', results.app.geral?.sim || 0, C.green)}
                            ${metricPercent('%', calcPercent(results.app.geral?.sim, results.app.geral?.total))}
                        </div>
                    </div>
                ` : ''}
                
                ${results.adimplencia ? `
                    <div class="card">
                        <div class="card-header">
                            <div class="card-title">Adimplência</div>
                            <div class="card-icon" style="background:${C.green}20;color:${C.green}"><i class="fas fa-credit-card"></i></div>
                        </div>
                        <div class="metric-grid">
                            ${metricItem('Aprovados', results.adimplencia.geral?.aprovados || 0, C.green)}
                            ${metricPercent('%', results.adimplencia.geral?.percentualAprovado || 0)}
                        </div>
                    </div>
                ` : ''}
            </div>
        `;
    } else {
        dashboardContent.innerHTML = '<h2>Não foi possível carregar o resumo</h2>';
    }
    
    updateLastUpdateTime();
}

// ============================================================================
// HELPERS
// ============================================================================

function metricItem(l, v, c){ 
    const s = c ? `style="color:${c};"` : ''; 
    return `<div class="metric-item"><div class="metric-label">${l}</div><div class="metric-value" ${s}>${v ?? 0}</div></div>`; 
}

function metricPercent(l, v){ 
    const cls = getPercentClass(typeof v === 'number' ? v : parseInt(v) || 0); 
    return `<div class="metric-item"><div class="metric-label">${l}</div><div class="metric-percent ${cls}">${v}%</div></div>`; 
}

function kpiCard(label, value, icon, cor){
    return `<div class="kpi-card"><div class="kpi-icon" style="background:${cor}20;color:${cor}"><i class="${icon}"></i></div><div class="kpi-body"><div class="kpi-label">${label}</div><div class="kpi-value" style="color:${cor}">${value}</div></div></div>`;
}

function calcPercent(n, d){ 
    if(!d || d === 0) return 0; 
    return Math.round((n/d)*100); 
}

function groupBySector(list){ 
    const m = {}; 
    list?.forEach(c => {
        const s = c.setor || 'OUTROS';
        if(!m[s]) m[s] = [];
        m[s].push(c);
    }); 
    return m; 
}

function sortSectors(keys, order){ 
    return keys.sort((a,b)=>{
        const ia = order.indexOf(a), ib = order.indexOf(b);
        if(ia === -1 && ib === -1) return a.localeCompare(b);
        if(ia === -1) return 1;
        if(ib === -1) return -1;
        return ia - ib;
    }); 
}

function getPercentClass(p){ 
    p = typeof p === 'number' ? p : parseInt(p) || 0; 
    if(p >= 90) return 'percent-high'; 
    if(p >= 80) return 'percent-medium'; 
    return 'percent-low'; 
}

function getSectorClass(s){ 
    switch((s||'').toUpperCase()){
        case 'VENDAS': return 'sector-vendas';
        case 'RECEPCAO': return 'sector-recepcao';
        case 'REFILIACAO': return 'sector-refiliacao';
        case 'RETENÇÃO': case 'RETENCAO': return 'sector-retencao';
        default: return 'sector-outros';
    } 
}

function getSectorIcon(s){ 
    switch((s||'').toUpperCase()){
        case 'VENDAS': return 'fas fa-shopping-cart';
        case 'RECEPCAO': return 'fas fa-headset';
        case 'REFILIACAO': return 'fas fa-user-plus';
        case 'RETENÇÃO': case 'RETENCAO': return 'fas fa-crown';
        default: return 'fas fa-users';
    } 
}

function scalesXY(){ 
    return {
        x:{grid:{display:false},border:{display:false}},
        y:{grid:{color:C.border},border:{display:false}}
    }; 
}

function legendTop(){ 
    return {position:'top',labels:{boxWidth:12,padding:14}}; 
}

function destroyChart(id){ 
    if(chartInstances[id]){
        chartInstances[id].destroy();
        delete chartInstances[id];
    } 
}

function createChart(id,cfg){ 
    destroyChart(id); 
    const ctx = document.getElementById(id); 
    if(!ctx) return; 
    chartInstances[id] = new Chart(ctx, cfg); 
}

function updateLastUpdateTime(){
    if(lastUpdateEl) lastUpdateEl.textContent = new Date().toLocaleString('pt-BR');
}

// ============================================================================
// LOADING / PROGRESS
// ============================================================================

function showProgressLoading(label='Carregando...', pct=0){
    hideLoading();
    const div = document.createElement('div');
    div.id = 'progressLoader';
    div.className = 'progress-loader';
    div.innerHTML = `
        <div class="progress-icon"><i class="fas fa-chart-line"></i></div>
        <div class="progress-label">${label}</div>
        <div class="progress-track"><div class="progress-bar" style="width:${pct}%"></div></div>
        <div class="progress-pct">${pct}%</div>
    `;
    dashboardContent.appendChild(div);
}

function updateProgress(pct, label){
    const bar = document.querySelector('.progress-bar');
    const lbl = document.querySelector('.progress-label');
    const num = document.querySelector('.progress-pct');
    if(bar) bar.style.width = pct + '%';
    if(lbl) lbl.textContent = label;
    if(num) num.textContent = pct + '%';
}

function hideLoading(){
    if(loadingEl) loadingEl.style.display = 'none';
    const pl = document.getElementById('progressLoader');
    if(pl) pl.remove();
}

function showToast(msg){
    let t = document.getElementById('toast');
    if(!t){ 
        t = document.createElement('div'); 
        t.id = 'toast'; 
        document.body.appendChild(t); 
    }
    t.textContent = msg; 
    t.className = 'toast toast-show';
    setTimeout(() => { t.className = 'toast'; }, 2500);
}

function showError(msg){ 
    hideLoading(); 
    dashboardContent.innerHTML = `
        <div class="error-message">
            <i class="fas fa-exclamation-triangle"></i>
            <h3>Erro ao carregar dados</h3>
            <p>${msg}</p>
            <button class="btn btn-success" onclick="loadDashboard()">Tentar Novamente</button>
        </div>
    `; 
}

// ============================================================================
// EXPORT
// ============================================================================

async function exportPage(){
    alert('Função de exportação será implementada em breve');
}

function getMonthName(m){ 
    return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][m-1]; 
}

// ============================================================================
// PAINEL ADMIN
// ============================================================================

function toggleAdminPanel() {
    alert('Painel Admin será ativado após testes');
}
