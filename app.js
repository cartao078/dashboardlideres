// ============================================================================
// DASHBOARD V21.0 - app.js — COM SUPABASE E PAINEL ADMIN
// ============================================================================

const API_URL = 'https://script.google.com/macros/s/AKfycbwULRFHnMXLZ6Hrop1xMiuGZSMWetfOjKe2NGUGdk4RzX51H2Nwk0lb5DXaLZxYRaYEKQ/exec';

const SUPABASE_URL  = 'https://vycjtmjvkwvxunxtkdyi.supabase.co';
const SUPABASE_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ5Y2p0bWp2a3d2eHVueHRrZHlpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE2MDY2OTYsImV4cCI6MjA4NzE4MjY5Nn0.wOoAZpA1i-320E8Rc-Ry6nk0KYsedFXb3aS4gkmbjHU';

// ============================================================================
// INICIALIZAÇÃO DO SUPABASE
// ============================================================================

let supabaseClient;
try {
    if (window.supabase) {
        supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON);
        console.log('✅ Supabase inicializado');
    } else {
        console.warn('⚠️ Biblioteca Supabase não encontrada');
        supabaseClient = null;
    }
} catch (e) {
    console.warn('⚠️ Erro ao inicializar Supabase:', e);
    supabaseClient = null;
}

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
    console.log('🔄 Fetching:', url);
    
    try {
        const resp = await fetch(url);
        if (!resp.ok) {
            const errMsg = `HTTP ${resp.status}: ${resp.statusText}`;
            console.error('❌ HTTP Error:', errMsg);
            showDiagnostic('API retornou erro: ' + errMsg, url);
            return { status: 'error', error: errMsg };
        }
        const text = await resp.text();
        console.log('📄 Raw response (' + endpoint + '):', text.substring(0, 300));
        
        let json;
        try {
            json = JSON.parse(text);
        } catch(parseErr) {
            console.error('❌ JSON inválido:', text.substring(0, 500));
            showDiagnostic('API retornou HTML em vez de JSON. A Web App não foi reimplantada.', url);
            return { status: 'error', error: 'API retornou HTML (não foi reimplantada). Acesse: Implantar → Gerenciar implantações → Nova versão.' };
        }
        
        console.log('✅ Response:', endpoint, json.status, json.data ? '(tem dados)' : '(sem dados)');
        return json;
    } catch (error) {
        console.error('❌ Fetch error:', error.message);
        showDiagnostic('Erro de rede: ' + error.message, url);
        return { status: 'error', error: 'Erro de rede/CORS: ' + error.message };
    }
}

function showDiagnostic(msg, url) {
    const el = document.getElementById('dashboardContent');
    if (!el) return;
    // Só mostra se ainda não tem conteúdo útil
    if (el.querySelector('.diag-box')) return;
    const div = document.createElement('div');
    div.className = 'diag-box';
    div.style.cssText = 'background:#fff3cd;border:1px solid #ffc107;border-radius:12px;padding:20px;margin:20px;font-family:monospace;';
    div.innerHTML = `<strong>⚠️ Diagnóstico</strong><br><br>
        <b>Erro:</b> ${msg}<br><br>
        <b>URL da API:</b><br><small>${url}</small><br><br>
        <b>O que fazer:</b><br>
        1. No Apps Script: Implantar → Gerenciar implantações → ✏️ Editar → Nova versão → Implantar<br>
        2. Confirme que o acesso é "Qualquer pessoa"<br>
        3. Clique em <b>Atualizar Dados</b> aqui no dashboard`;
    el.prepend(div);
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

    // Inicializar botão admin
    initAdminButton();
    
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
        
        // Atualização em background
        fetchData(currentDashboard, currentMonth, currentYear).then(r => {
            if(r.status === 'success' && JSON.stringify(r.data) !== JSON.stringify(cached)){
                setCache(k, r.data);
                window._dashboardData = r.data;
                renderDashboard();
                updateLastUpdateTime();
                showToast('Dados atualizados ✓');
            }
        }).catch(() => {});
        
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
    }
}

// ============================================================================
// DASHBOARD: VENDAS (DOCUMENTAÇÃO)
// ============================================================================

function renderDocumentacaoDashboard(d){
    const {geral, vendasLoja, vendasWeb, consultores, mes, ano} = d;
    const pG = calcPercent(geral.aprovados, geral.total);
    const pL = calcPercent(vendasLoja.aprovados, vendasLoja.total);
    const pW = calcPercent(vendasWeb.aprovados, vendasWeb.total);
    
    const bySector = groupBySector(consultores);
    const order = ['VENDAS','RECEPCAO','REFILIACAO','WEB SITE','TELEVENDAS','OUTROS'];
    const sectors = sortSectors(Object.keys(bySector), order);
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-folder" style="color:var(--primary)"></i> Dashboard de Vendas — ${mes} ${ano}</h2>
        <div class="main-cards">
            ${cardDoc('Total de Vendas', 'fas fa-chart-bar', geral, pG)}
            ${cardDoc('Vendas Loja', 'fas fa-store', vendasLoja, pL)}
            ${cardDoc('Vendas Web/Tele', 'fas fa-globe', vendasWeb, pW)}
        </div>
    `;
    
    if (sectors.length > 0) {
        html += `<h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--primary)"></i> Desempenho por Setor</h3>`;
        
        sectors.forEach(sector => {
            const list = bySector[sector];
            const tot = list.reduce((s,c) => s + c.total, 0);
            const aprov = list.reduce((s,c) => s + c.aprovados, 0);
            const pct = calcPercent(aprov, tot);
            
            html += `
                <div class="sector-card">
                    <div class="sector-header">
                        <div class="sector-title">
                            <i class="${getSectorIcon(sector)}"></i> ${sector}
                            <span class="sector-count">${list.length} consultor${list.length !== 1 ? 'es' : ''}</span>
                        </div>
                        <div class="metric-percent ${getPercentClass(pct)}">${pct}% aprovados</div>
                    </div>
                    <div class="consultant-grid">
            `;
            
            list.sort((a,b) => b.total - a.total).forEach(c => {
                const p = calcPercent(c.aprovados, c.total);
                html += `
                    <div class="consultant-card">
                        <div class="consultant-header">
                            <div class="consultant-name">${c.nome}</div>
                            <div class="consultant-sector ${getSectorClass(sector)}">${sector}</div>
                        </div>
                        <div class="metric-grid">
                            ${metricItem('Total', c.total)}
                            ${metricItem('Aprovados', c.aprovados, 'var(--success)')}
                            ${metricItem('Pendências', c.pendencias, 'var(--danger)')}
                            ${metricItem('Não Enviado', c.naoEnviado, 'var(--warning)')}
                            ${metricItem('Expirado', c.expirado, 'var(--gray)')}
                            ${metricPercent('% Aprovados', p)}
                        </div>
                    </div>
                `;
            });
            
            html += `</div></div>`;
        });
    }
    
    dashboardContent.innerHTML = html;
}

function cardDoc(t, icon, d, pct){
    return `
        <div class="card card-doc">
            <div class="card-header">
                <div class="card-title">${t}</div>
                <div class="card-icon"><i class="${icon}"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total Vendas', d.total)}
                ${metricItem('Aprovados', d.aprovados, 'var(--success)')}
                ${metricItem('Pendências', d.pendencias, 'var(--danger)')}
                ${metricItem('Não Enviado', d.naoEnviado, 'var(--warning)')}
                ${metricItem('Expirado', d.expirado, 'var(--gray)')}
                ${metricPercent('% Aprovados', pct)}
            </div>
        </div>
    `;
}

// ============================================================================
// DASHBOARD: APP
// ============================================================================

function renderAppDashboard(d){
    const {geral, appLoja, appWeb, consultores, consultorasRetencao, mes, ano} = d;
    const pG = calcPercent(geral.sim, geral.total);
    const pL = calcPercent(appLoja.sim, appLoja.total);
    const pW = calcPercent(appWeb.sim, appWeb.total);
    
    const regular = consultores.filter(c => !c.origem || c.origem !== 'retencao');
    const bySector = groupBySector(regular);
    const order = ['VENDAS','RECEPCAO','REFILIACAO','WEB SITE','TELEVENDAS','OUTROS'];
    const sectors = sortSectors(Object.keys(bySector), order);
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-mobile-alt" style="color:var(--teal)"></i> Dashboard App — ${mes} ${ano}</h2>
        <div class="main-cards">
            ${cardApp('App — Total Geral', 'fas fa-chart-pie', geral, pG)}
            ${cardApp('App — Loja', 'fas fa-store', appLoja, pL)}
            ${cardApp('App — Web/Tele', 'fas fa-globe', appWeb, pW)}
        </div>
    `;
    
    if (consultorasRetencao && consultorasRetencao.length > 0) {
        html += `
            <div class="retention-section">
                <div class="retention-header">
                    <i class="fas fa-crown"></i>
                    <h3>Consultoras de Retenção</h3>
                </div>
                <div class="consultant-grid">
        `;
        
        consultorasRetencao.forEach(c => {
            const p = calcPercent(c.sim, c.total);
            html += `
                <div class="consultant-card" style="border-left:3px solid #f59e0b;">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome} (RETENÇÃO)</div>
                        <div class="consultant-sector sector-retencao">RETENÇÃO</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Total', c.total)}
                        ${metricItem('Com App', c.sim, 'var(--success)')}
                        ${metricItem('Sem App', c.nao, 'var(--danger)')}
                        ${metricItem('Cancelados', c.cancelado || 0, 'var(--gray)')}
                        ${metricPercent('% Com App', p)}
                    </div>
                </div>
            `;
        });
        
        html += `</div></div>`;
    }
    
    if (sectors.length > 0) {
        html += `<h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--teal)"></i> Desempenho por Setor</h3>`;
        
        sectors.forEach(sector => {
            const list = bySector[sector];
            const tot = list.reduce((s,c) => s + c.total, 0);
            const sim = list.reduce((s,c) => s + (c.sim || 0), 0);
            const pct = calcPercent(sim, tot);
            
            html += `
                <div class="sector-card">
                    <div class="sector-header">
                        <div class="sector-title">
                            <i class="${getSectorIcon(sector)}"></i> ${sector}
                            <span class="sector-count">${list.length} consultor${list.length !== 1 ? 'es' : ''}</span>
                        </div>
                        <div class="metric-percent ${getPercentClass(pct)}">${pct}% com app</div>
                    </div>
                    <div class="consultant-grid">
            `;
            
            list.sort((a,b) => b.total - a.total).forEach(c => {
                const p = calcPercent(c.sim || 0, c.total);
                html += `
                    <div class="consultant-card">
                        <div class="consultant-header">
                            <div class="consultant-name">${c.nome}</div>
                            <div class="consultant-sector ${getSectorClass(sector)}">${sector}</div>
                        </div>
                        <div class="metric-grid">
                            ${metricItem('Total', c.total)}
                            ${metricItem('Com App', c.sim || 0, 'var(--success)')}
                            ${metricItem('Sem App', c.nao || 0, 'var(--danger)')}
                            ${metricItem('Cancelados', c.cancelado || 0, 'var(--gray)')}
                            ${metricPercent('% Com App', p)}
                        </div>
                    </div>
                `;
            });
            
            html += `</div></div>`;
        });
    }
    
    dashboardContent.innerHTML = html;
}

function cardApp(t, icon, d, pct){
    return `
        <div class="card card-app">
            <div class="card-header">
                <div class="card-title">${t}</div>
                <div class="card-icon"><i class="${icon}"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total Clientes', d.total)}
                ${metricItem('Com App (SIM)', d.sim, 'var(--success)')}
                ${metricItem('Sem App (NÃO)', d.nao, 'var(--danger)')}
                ${metricItem('Cancelados', d.cancelado || 0, 'var(--gray)')}
                ${metricPercent('% Com App', pct)}
            </div>
        </div>
    `;
}

// ============================================================================
// DASHBOARD: ADIMPLÊNCIA
// ============================================================================

function renderAdimplenciaDashboard(d){
    const {geral, consultores, mes, ano} = d;
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-credit-card" style="color:var(--success)"></i> Dashboard Adimplência — ${mes} ${ano}</h2>
        <div class="card card-adim" style="max-width:600px;margin:0 auto 28px;">
            <div class="card-header">
                <div class="card-title">Adimplência — Total da Loja</div>
                <div class="card-icon"><i class="fas fa-chart-line"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total Trocas', geral.totalTrocas)}
                ${metricItem('Mens. OK', geral.mensOk, 'var(--success)')}
                ${metricItem('Mens. Aberto', geral.mensAberto, 'var(--warning)')}
                ${metricItem('Mens. Atraso', geral.mensAtraso, 'var(--danger)')}
                ${metricItem('Aprovados', geral.aprovados, 'var(--success)')}
                ${metricItem('Pendentes', geral.pendentes, 'var(--danger)')}
                ${metricItem('Total BI', geral.totalBi)}
                ${metricItem('Fora BI', geral.foraBi, 'var(--danger)')}
                ${metricPercent('% Aprovados', geral.percentualAprovado)}
            </div>
        </div>
    `;
    
    if (consultores && consultores.length > 0) {
        html += `
            <h3 class="section-title"><i class="fas fa-user-tie" style="color:var(--primary)"></i> Desempenho por Consultor</h3>
            <div class="table-wrapper">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Consultor</th>
                            <th>Total Trocas</th>
                            <th>Mens. OK</th>
                            <th>Mens. Aberto</th>
                            <th>Mens. Atraso</th>
                            <th>Aprovados</th>
                            <th>Pendentes</th>
                            <th>Total BI</th>
                            <th>Fora BI</th>
                            <th>% Aprovados</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        consultores.forEach(c => {
            html += `
                <tr>
                    <td><strong>${c.nome}</strong></td>
                    <td>${c.totalTrocas}</td>
                    <td style="color:var(--success)">${c.mensOk}</td>
                    <td style="color:var(--warning)">${c.mensAberto}</td>
                    <td style="color:var(--danger)">${c.mensAtraso}</td>
                    <td style="color:var(--success)">${c.aprovados}</td>
                    <td style="color:var(--danger)">${c.pendentes}</td>
                    <td>${c.totalBi}</td>
                    <td style="color:var(--danger)">${c.foraBi}</td>
                    <td><span class="metric-percent ${getPercentClass(c.percentualAprovado)}">${c.percentualAprovado}%</span></td>
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
    const {retencao, refiliacao, periodo} = d;
    const retKeys = Object.keys(retencao || {});
    const refKeys = Object.keys(refiliacao || {});
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-redo" style="color:var(--warning)"></i> Dashboard Recorrência — ${periodo.atual}</h2>
        <div style="background:rgba(245,158,11,0.08);padding:14px 18px;border-radius:12px;margin-bottom:24px;border-left:4px solid var(--warning);">
            <p style="margin:0;color:#92400e;font-weight:600;font-size:0.9rem;">
                <i class="fas fa-info-circle"></i> Período atual: ${periodo.atual} | Histórico: ${periodo.historico.join(', ')}
            </p>
        </div>
    `;
    
    if (retKeys.length > 0) {
        html += `<h3 class="section-title" style="border-bottom:2px solid var(--primary)"><i class="fas fa-crown" style="color:var(--primary)"></i> Retenção</h3>`;
        html += `<div class="consultant-grid" style="margin-bottom:36px">`;
        
        retKeys.forEach(key => {
            const c = retencao[key];
            const pA = calcPercent(c.atual.retençõesOK, c.atual.totalRetidosFinal);
            const pT = calcPercent(c.total3Meses.totalOK || 0, c.total3Meses.totalRetidosFinal);
            
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${key} (RETENÇÃO)</div>
                        <div class="consultant-sector sector-vendas">RETENÇÃO</div>
                    </div>
                    <h4 class="sub-section-title"><i class="far fa-calendar-check"></i> Mês Atual (${periodo.atual})</h4>
                    <div class="metric-grid">
                        ${metricItem('Total Retido', c.atual.totalRetido)}
                        ${metricItem('Cancelados', c.atual.cancelado, 'var(--danger)')}
                        ${metricItem('Retenções OK', c.atual.retençõesOK, 'var(--success)')}
                        ${metricItem('Pendências KYC', c.atual.pendenciasKYC, 'var(--warning)')}
                        ${metricPercent('% OK', pA)}
                    </div>
                    <h4 class="sub-section-title"><i class="fas fa-chart-line"></i> Total 3 Meses</h4>
                    <div class="metric-grid">
                        ${metricItem('Total Retido', c.total3Meses.totalRetido)}
                        ${metricItem('Cancelados', c.total3Meses.cancelado || 0, 'var(--danger)')}
                        ${metricItem('Total OK', c.total3Meses.totalOK || 0, 'var(--success)')}
                        ${metricItem('Em Atraso', c.total3Meses.emAtraso, 'var(--warning)')}
                        ${metricPercent('% OK', pT)}
                    </div>
                </div>
            `;
        });
        
        html += `</div>`;
    }
    
    if (refKeys.length > 0) {
        html += `<h3 class="section-title" style="border-bottom:2px solid var(--warning)"><i class="fas fa-user-plus" style="color:var(--warning)"></i> Refiliação</h3>`;
        html += `<div class="consultant-grid">`;
        
        refKeys.forEach(key => {
            const c = refiliacao[key];
            const pT = calcPercent(c.total3Meses.totalOK || 0, c.total3Meses.totalRetidosFinal);
            
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${key} (REFILIAÇÃO)</div>
                        <div class="consultant-sector sector-refiliacao">REFILIAÇÃO</div>
                    </div>
                    <h4 class="sub-section-title"><i class="fas fa-chart-line"></i> Total 3 Meses</h4>
                    <div class="metric-grid">
                        ${metricItem('Total Refiliados', c.total3Meses.totalRetido)}
                        ${metricItem('Cancelados', c.total3Meses.cancelado || 0, 'var(--danger)')}
                        ${metricItem('Total OK', c.total3Meses.totalOK || 0, 'var(--success)')}
                        ${metricItem('Em Atraso', c.total3Meses.emAtraso, 'var(--warning)')}
                        ${metricPercent('% OK', pT)}
                    </div>
                </div>
            `;
        });
        
        html += `</div>`;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: REC. VENDEDOR
// ============================================================================

function renderRecorrenciaVendedorDashboard(d){
    const {geral, consultores, totalConsultores, totalRegistros} = d;
    
    if (!geral || !consultores) {
        showError('Dados de Recorrência Vendedor não encontrados');
        return;
    }
    
    const pG = calcPercent(geral.totalOk + geral.totalEmAberto, geral.totalVendasPromocao);
    const setores = [
        {nome:'VENDAS', cor:C.green, icone:'fas fa-shopping-cart'},
        {nome:'REFILIACAO', cor:C.warn, icone:'fas fa-user-plus'},
        {nome:'RECEPCAO', cor:C.teal, icone:'fas fa-headset'}
    ];
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-handshake" style="color:var(--teal)"></i> Recorrência Vendedor — Dados Completos</h2>
        <div style="background:#e0f2f1;padding:14px 18px;border-radius:12px;margin-bottom:24px;border-left:4px solid var(--teal)">
            <p style="margin:0;color:#065f46;font-weight:600;font-size:0.9rem">
                <i class="fas fa-database"></i> <strong>${totalRegistros || 0}</strong> registros | 
                <strong>${totalConsultores}</strong> consultores
            </p>
        </div>
        <div class="main-cards">
            <div class="card" style="border-top:3px solid var(--teal)">
                <div class="card-header">
                    <div class="card-title">Total Geral</div>
                    <div class="card-icon" style="background:#e0f2f1;color:var(--teal)"><i class="fas fa-chart-bar"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItem('Total Vendas', geral.totalVendasPromocao)}
                    ${metricItem('OK', geral.totalOk, 'var(--success)')}
                    ${metricItem('Em Aberto', geral.totalEmAberto, 'var(--warning)')}
                    ${metricItem('Atraso', geral.totalAtraso, 'var(--danger)')}
                    ${metricItem('Outros', geral.totalOutros || 0, 'var(--gray)')}
                    ${metricPercent('% OK+Aberto', pG)}
                </div>
            </div>
        </div>
    `;
    
    setores.forEach(si => {
        const lista = consultores.filter(c => c.setor === si.nome);
        if (!lista.length) return;
        
        const tot = lista.reduce((s,c) => s + c.totalVendasPromocao, 0);
        const ok = lista.reduce((s,c) => s + (c.totalOk || 0) + (c.totalEmAberto || 0), 0);
        const pct = calcPercent(ok, tot);
        
        html += `
            <div class="sector-card">
                <div class="sector-header">
                    <div class="sector-title">
                        <i class="${si.icone}" style="color:${si.cor}"></i>
                        <span style="color:${si.cor}">${si.nome}</span>
                        <span class="sector-count">${lista.length} consultor${lista.length !== 1 ? 'es' : ''}</span>
                    </div>
                    <div class="metric-percent ${getPercentClass(pct)}">${pct}% OK+Aberto</div>
                </div>
                <div class="consultant-grid">
        `;
        
        lista.sort((a,b) => b.totalVendasPromocao - a.totalVendasPromocao).forEach(c => {
            const p = calcPercent((c.totalOk || 0) + (c.totalEmAberto || 0), c.totalVendasPromocao);
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome}</div>
                        <div class="consultant-sector ${getSectorClass(c.setor)}">${c.setor}</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Total Vendas', c.totalVendasPromocao)}
                        ${metricItem('OK', c.totalOk || 0, 'var(--success)')}
                        ${metricItem('Em Aberto', c.totalEmAberto || 0, 'var(--warning)')}
                        ${metricItem('Atraso', c.totalAtraso || 0, 'var(--danger)')}
                        ${metricItem('Outros', c.totalOutros || 0, 'var(--gray)')}
                        ${metricPercent('% OK+Aberto', p)}
                    </div>
                </div>
            `;
        });
        
        html += `</div></div>`;
    });
    
    html += `
        <div class="card" style="margin-top:36px;background:linear-gradient(135deg,var(--primary),#007a3d);color:white">
            <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2)">
                <div class="card-title" style="color:white">📊 Resumo Final</div>
                <div class="card-icon" style="background:rgba(255,255,255,0.2)"><i class="fas fa-clipboard-list"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItemWhite('Total Consultores', totalConsultores || consultores.length)}
                ${metricItemWhite('Total Vendas', geral.totalVendasPromocao)}
                ${metricItemWhite('Média por Consultor', consultores.length > 0 ? Math.round(geral.totalVendasPromocao/consultores.length) : 0)}
                ${metricItemWhite('Melhor %', consultores.length > 0 ? Math.max(...consultores.map(c => calcPercent((c.totalOk||0)+(c.totalEmAberto||0), c.totalVendasPromocao))) + '%' : '0%')}
            </div>
        </div>
    `;
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: REFUTURIZA
// ============================================================================

function renderRefuturizaDashboard(d){
    const {geral, consultores, mes, ano} = d;
    
    if (!geral || !consultores) {
        showError('Dados do Refuturiza não encontrados');
        return;
    }
    
    const pG = calcPercent(geral.comLigacao, geral.total);
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-book" style="color:#0ea5e9"></i> Dashboard Refuturiza — ${mes} ${ano}</h2>
        <div class="card card-refut" style="max-width:500px;margin:0 auto 28px;">
            <div class="card-header">
                <div class="card-title">Refuturiza — Total da Loja</div>
                <div class="card-icon"><i class="fas fa-book-open"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total', geral.total || 0)}
                ${metricItem('Com Ligação', geral.comLigacao || 0, 'var(--success)')}
                ${metricItem('Sem Ligação', geral.semLigacao || 0, 'var(--danger)')}
                ${metricItem('Cancelados', geral.cancelado || 0, 'var(--gray)')}
                ${metricPercent('% Com Ligação', pG)}
            </div>
        </div>
    `;
    
    if (consultores.length > 0) {
        html += `
            <h3 class="section-title"><i class="fas fa-users" style="color:#0ea5e9"></i> Desempenho por Consultor (${consultores.length})</h3>
            <div class="consultant-grid">
        `;
        
        consultores.forEach(c => {
            const p = calcPercent(c.comLigacao || 0, c.total);
            html += `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome}</div>
                        <div class="consultant-sector" style="background:rgba(14,165,233,0.1);color:#0ea5e9">REFUTURIZA</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Total', c.total || 0)}
                        ${metricItem('Com Ligação', c.comLigacao || 0, 'var(--success)')}
                        ${metricItem('Sem Ligação', c.semLigacao || 0, 'var(--danger)')}
                        ${metricItem('Cancelados', c.cancelado || 0, 'var(--gray)')}
                        ${metricPercent('% Com Ligação', p)}
                    </div>
                </div>
            `;
        });
        
        html += `</div>`;
        
        html += `
            <div class="card" style="margin-top:36px;background:linear-gradient(90deg,#0ea5e9,#3b82f6);color:white">
                <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2)">
                    <div class="card-title" style="color:white">📊 Resumo Final</div>
                    <div class="card-icon" style="background:rgba(255,255,255,0.2)"><i class="fas fa-graduation-cap"></i></div>
                </div>
                <div class="metric-grid">
                    ${metricItemWhite('Total Consultores', consultores.length)}
                    ${metricItemWhite('Total Cursos', geral.total || 0)}
                    ${metricItemWhite('Média por Consultor', consultores.length > 0 ? Math.round((geral.total || 0)/consultores.length) : 0)}
                    ${metricItemWhite('Taxa de Contato', pG + '%')}
                </div>
            </div>
        `;
    } else {
        html += `
            <div style="text-align:center;padding:40px;color:var(--gray)">
                <i class="fas fa-info-circle" style="font-size:2.5rem;margin-bottom:16px;display:block;"></i>
                <h3>Nenhum dado para ${mes} ${ano}</h3>
            </div>
        `;
    }
    
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: RESUMO
// ============================================================================

async function loadResumoDashboard(){
    const eps = ['documentacao', 'app', 'adimplencia'];
    const results = {};
    const labels = {documentacao:'Vendas', app:'App', adimplencia:'Adimplência'};
    
    const allCached = eps.every(ep => !!getCached(cacheKey(ep, currentMonth, currentYear)));
    
    if (allCached) {
        eps.forEach(ep => { results[ep] = getCached(cacheKey(ep, currentMonth, currentYear)); });
        renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
        updateLastUpdateTime();
        
        Promise.all(eps.map(async ep => {
            const r = await fetchData(ep, currentMonth, currentYear);
            if (r.status === 'success') {
                setCache(cacheKey(ep, currentMonth, currentYear), r.data);
                results[ep] = r.data;
            }
        })).then(() => {
            if (results['documentacao']) renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
        }).catch(() => {});
        
        return;
    }
    
    let loaded = 0;
    showProgressLoading('Iniciando...', 0);
    
    await Promise.all(eps.map(async ep => {
        const k = cacheKey(ep, currentMonth, currentYear);
        const cached = getCached(k);
        
        if (cached) {
            results[ep] = cached;
            loaded++;
            updateProgress(Math.round((loaded/eps.length)*100), `${labels[ep]} ✓`);
            return;
        }
        
        try {
            updateProgress(Math.round((loaded/eps.length)*100), `Carregando ${labels[ep]}...`);
            const r = await fetchData(ep, currentMonth, currentYear);
            if (r.status === 'success') {
                setCache(k, r.data);
                results[ep] = r.data;
            }
        } catch(_) {}
        
        loaded++;
        updateProgress(Math.round((loaded/eps.length)*100), `${labels[ep]} ✓`);
    }));
    
    hideLoading();
    
    if (!results['documentacao']) {
        showError('Não foi possível carregar dados de Vendas.');
        return;
    }
    
    renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
    updateLastUpdateTime();
}

function renderResumoDashboard(vendas, app, adim){
    const {geral, consultores, mes, ano} = vendas;
    const pAprov = calcPercent(geral.aprovados, geral.total);
    const pApp = app ? calcPercent(app.geral.sim, app.geral.total) : null;
    const pAdim = adim ? adim.geral.percentualAprovado : null;
    
    const bySector = groupBySector(consultores);
    const order = ['VENDAS','RECEPCAO','REFILIACAO','WEB SITE','TELEVENDAS','OUTROS'];
    const sectors = sortSectors(Object.keys(bySector), order);
    
    const consultoresRanking = consultores
        .filter(c => RANKING_SETORES.includes((c.setor||'').toUpperCase()))
        .sort((a,b) => b.total - a.total);
    
    let html = `
        <h2 class="dash-title"><i class="fas fa-chart-line" style="color:var(--primary)"></i> Resumo Geral — ${mes} ${ano}</h2>
        <div class="kpi-strip">
            ${kpiCard('Total de Vendas', geral.total, 'fas fa-shopping-bag', C.green)}
            ${kpiCard('% Aprovação', pAprov+'%', 'fas fa-check-circle', pAprov>=90 ? C.green : pAprov>=80 ? C.warn : C.danger)}
            ${app ? kpiCard('% Com App', pApp+'%', 'fas fa-mobile-alt', pApp>=90 ? C.green : pApp>=80 ? C.warn : C.danger) : ''}
            ${adim ? kpiCard('% Adimplência', pAdim+'%', 'fas fa-credit-card', pAdim>=90 ? C.green : pAdim>=80 ? C.warn : C.danger) : ''}
        </div>
        <div class="charts-row">
            <div class="chart-card">
                <div class="chart-card-title"><i class="fas fa-chart-bar"></i> Vendas por Setor</div>
                <div class="chart-card-subtitle">Aprovados e pendências de cada setor no mês</div>
                <div class="chart-wrap"><canvas id="chRes1"></canvas></div>
            </div>
            <div class="chart-card">
                <div class="chart-card-title"><i class="fas fa-chart-pie"></i> Status das Vendas</div>
                <div class="chart-card-subtitle">Distribuição geral: aprovados, pendências, não enviado e expirado</div>
                <div class="chart-wrap"><canvas id="chRes2"></canvas></div>
            </div>
            ${app ? `
            <div class="chart-card">
                <div class="chart-card-title"><i class="fas fa-mobile-alt"></i> Adesão ao App por Setor</div>
                <div class="chart-card-subtitle">Clientes com e sem app cadastrado, separado por setor</div>
                <div class="chart-wrap"><canvas id="chRes3"></canvas></div>
            </div>` : ''}
        </div>
        <div class="charts-row" style="grid-template-columns:${adim ? '1fr 1fr' : '1fr'}">
            ${adim ? `
            <div class="chart-card">
                <div class="chart-card-title"><i class="fas fa-credit-card"></i> Situação das Mensalidades</div>
                <div class="chart-card-subtitle">Mensalidades em dia, em aberto e em atraso</div>
                <div class="chart-wrap"><canvas id="chRes4"></canvas></div>
            </div>` : ''}
            <div class="chart-card">
                <div class="chart-card-title"><i class="fas fa-store"></i> Loja vs Web/Tele</div>
                <div class="chart-card-subtitle">Comparativo de total, aprovados e pendências entre os canais de venda</div>
                <div class="chart-wrap"><canvas id="chRes5"></canvas></div>
            </div>
        </div>
        ${buildRankingCompleto(consultoresRanking, mes, ano)}
    `;
    
    dashboardContent.innerHTML = html;
    
    // Criar gráficos
    setTimeout(() => {
        createChart('chRes1', {
            type:'bar',
            data:{
                labels:sectors,
                datasets:[
                    {label:'Aprovados', data:sectors.map(s=>bySector[s].reduce((a,c)=>a+c.aprovados,0)), backgroundColor:C.green, borderRadius:6, borderSkipped:false},
                    {label:'Pendências', data:sectors.map(s=>bySector[s].reduce((a,c)=>a+c.pendencias,0)), backgroundColor:C.warn, borderRadius:6, borderSkipped:false}
                ]
            },
            options:{responsive:true, maintainAspectRatio:false, plugins:{legend:legendTop()}, scales:scalesXY()}
        });
        
        createChart('chRes2', {
            type:'doughnut',
            data:{
                labels:['Aprovados','Pendências','Não Enviado','Expirado'],
                datasets:[{data:[geral.aprovados, geral.pendencias, geral.naoEnviado, geral.expirado||0], backgroundColor:[C.green, C.warn, C.danger, C.gray], borderWidth:0, hoverOffset:6}]
            },
            options:{responsive:true, maintainAspectRatio:false, cutout:'68%', plugins:{legend:{position:'bottom', labels:{boxWidth:12, padding:12}}}}
        });
        
        if (app) {
            const appBySector = groupBySector(app.consultores.filter(c=>!c.origem||c.origem!=='retencao'));
            createChart('chRes3', {
                type:'bar',
                data:{
                    labels:sectors,
                    datasets:[
                        {label:'Com App', data:sectors.map(s=>(appBySector[s]||[]).reduce((a,c)=>a+(c.sim||0),0)), backgroundColor:C.teal, borderRadius:6, borderSkipped:false},
                        {label:'Sem App', data:sectors.map(s=>(appBySector[s]||[]).reduce((a,c)=>a+(c.nao||0),0)), backgroundColor:C.danger, borderRadius:6, borderSkipped:false}
                    ]
                },
                options:{responsive:true, maintainAspectRatio:false, plugins:{legend:legendTop()}, scales:scalesXY()}
            });
        }
        
        if (adim) {
            createChart('chRes4', {
                type:'doughnut',
                data:{
                    labels:['Mens. OK','Mens. Aberto','Mens. Atraso'],
                    datasets:[{data:[adim.geral.mensOk, adim.geral.mensAberto, adim.geral.mensAtraso], backgroundColor:[C.green, C.warn, C.danger], borderWidth:0, hoverOffset:6}]
                },
                options:{responsive:true, maintainAspectRatio:false, cutout:'68%', plugins:{legend:{position:'bottom', labels:{boxWidth:12, padding:12}}}}
            });
        }
        
        createChart('chRes5', {
            type:'bar',
            data:{
                labels:['Total','Aprovados','Pendências'],
                datasets:[
                    {label:'Loja', data:[vendas.vendasLoja.total, vendas.vendasLoja.aprovados, vendas.vendasLoja.pendencias], backgroundColor:C.green, borderRadius:6, borderSkipped:false},
                    {label:'Web/Tele', data:[vendas.vendasWeb.total, vendas.vendasWeb.aprovados, vendas.vendasWeb.pendencias], backgroundColor:C.lime, borderRadius:6, borderSkipped:false}
                ]
            },
            options:{responsive:true, maintainAspectRatio:false, plugins:{legend:legendTop()}, scales:scalesXY()}
        });
    }, 100);
}

function buildRankingCompleto(consultores, mes, ano){
    if (!consultores.length) return '';
    
    const medal = i => i===0 ? '🥇' : i===1 ? '🥈' : i===2 ? '🥉' : (i+1);
    const posClass = i => i===0 ? 'rank-gold' : i===1 ? 'rank-silver' : i===2 ? 'rank-bronze' : 'rank-other';
    const maxTotal = consultores[0].total || 1;
    const porSetor = {};
    
    RANKING_SETORES.forEach(s => { 
        porSetor[s] = consultores.filter(c => (c.setor||'').toUpperCase() === s); 
    });
    
    const top10 = consultores.slice(0, 10);
    
    return `
        <div class="ranking-full">
            <div class="ranking-full-header">
                <i class="fas fa-trophy"></i> 🏆 Ranking de Vendedores — ${mes} ${ano}
                <span class="ranking-badge">Vendas · Recepção · Refiliação</span>
            </div>
            <div class="ranking-section-title">Top 10 Geral</div>
            <div class="rank-grid">
                ${top10.map((c,i) => {
                    const pct = calcPercent(c.aprovados, c.total);
                    const barW = Math.round((c.total/maxTotal)*100);
                    return `
                        <div class="rank-card ${i<3 ? 'rank-card-destaque' : ''}">
                            <div class="rank-card-pos ${posClass(i)}">${medal(i)}</div>
                            <div class="rank-card-info">
                                <div class="rank-card-name">${c.nome}</div>
                                <div class="rank-card-sector ${getSectorClass(c.setor)}">${c.setor || '—'}</div>
                            </div>
                            <div class="rank-card-metrics">
                                <div class="rank-card-total">${c.total} <span>vendas</span></div>
                                <div class="rank-card-pct ${getPercentClass(pct)}">${pct}%</div>
                            </div>
                            <div class="rank-bar-wrap" style="width:120px">
                                <div class="rank-bar-bg"><div class="rank-bar-fill" style="width:${barW}%"></div></div>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
            <div class="ranking-setores">
                ${RANKING_SETORES.map(setor => {
                    const lista = porSetor[setor];
                    if (!lista || !lista.length) return '';
                    return `
                        <div class="ranking-setor-col">
                            <div class="ranking-setor-header ${getSectorClass(setor)}">
                                <i class="${getSectorIcon(setor)}"></i> ${setor}
                            </div>
                            ${lista.slice(0,5).map((c,i) => {
                                const pct = calcPercent(c.aprovados, c.total);
                                return `
                                    <div class="rank-setor-row">
                                        <div class="rank-pos ${posClass(i)}" style="width:24px;height:24px;font-size:0.7rem">${medal(i)}</div>
                                        <div class="rank-name" style="font-size:0.83rem">${c.nome}</div>
                                        <div class="rank-pct ${getPercentClass(pct)}" style="font-size:0.78rem">${pct}%</div>
                                        <div class="rank-total" style="font-size:0.75rem">${c.total}vd</div>
                                    </div>
                                `;
                            }).join('')}
                        </div>
                    `;
                }).join('')}
            </div>
        </div>
    `;
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

function metricItemWhite(l, v){ 
    return `<div class="metric-item"><div class="metric-label" style="color:rgba(255,255,255,0.75)">${l}</div><div class="metric-value" style="color:white">${v}</div></div>`; 
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
    if (list) {
        list.forEach(c => {
            const s = c.setor || 'OUTROS';
            if(!m[s]) m[s] = [];
            m[s].push(c);
        });
    }
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
        case 'WEB SITE': case 'WEB': return 'sector-web';
        case 'TELEVENDAS': return 'sector-televendas';
        case 'RETENÇÃO': case 'RETENCAO': return 'sector-retencao';
        default: return 'sector-outros';
    } 
}

function getSectorIcon(s){ 
    switch((s||'').toUpperCase()){
        case 'VENDAS': return 'fas fa-shopping-cart';
        case 'RECEPCAO': return 'fas fa-headset';
        case 'REFILIACAO': return 'fas fa-user-plus';
        case 'WEB SITE': case 'WEB': return 'fas fa-globe';
        case 'TELEVENDAS': return 'fas fa-phone-alt';
        case 'RETENÇÃO': case 'RETENCAO': return 'fas fa-crown';
        default: return 'fas fa-users';
    } 
}

function scalesXY(){ 
    return {
        x:{grid:{display:false}, border:{display:false}},
        y:{grid:{color:C.border}, border:{display:false}}
    }; 
}

function legendTop(){ 
    return {position:'top', labels:{boxWidth:12, padding:14}}; 
}

function destroyChart(id){ 
    if(chartInstances[id]){
        chartInstances[id].destroy();
        delete chartInstances[id];
    } 
}

function createChart(id, cfg){ 
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
    Array.from(dashboardContent.children).forEach(c => { if(c !== loadingEl) c.remove(); });
    if(loadingEl) loadingEl.style.display = 'none';
    
    let pl = document.getElementById('progressLoader');
    if(pl) pl.remove();
    
    pl = document.createElement('div');
    pl.id = 'progressLoader';
    pl.innerHTML = `
        <div class="progress-loader">
            <div class="progress-icon"><i class="fas fa-chart-line"></i></div>
            <div class="progress-label" id="progressLabel">${label}</div>
            <div class="progress-track"><div class="progress-bar" id="progressBar" style="width:${pct}%"></div></div>
            <div class="progress-pct" id="progressPct">${pct}%</div>
        </div>
    `;
    dashboardContent.appendChild(pl);
}

function updateProgress(pct, label){
    const bar = document.getElementById('progressBar');
    const lbl = document.getElementById('progressLabel');
    const num = document.getElementById('progressPct');
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
            <i class="fas fa-exclamation-triangle" style="font-size:2rem;margin-bottom:12px;display:block;"></i>
            <h3 style="margin-bottom:8px">Erro ao carregar dados</h3>
            <p>${msg}</p>
            <button class="btn btn-success" onclick="loadDashboard()" style="margin-top:16px;background:var(--danger);color:white">
                <i class="fas fa-redo"></i> Tentar Novamente
            </button>
        </div>
    `; 
}

// ============================================================================
// EXPORT
// ============================================================================

async function exportPage(){
    const btn = document.getElementById('downloadBtn');
    const orig = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando imagem...';
    btn.disabled = true;
    
    try {
        const container = document.querySelector('.container');
        const canvas = await html2canvas(container, {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: -window.scrollY,
            windowWidth: container.scrollWidth,
            windowHeight: container.scrollHeight,
            onclone: (doc) => { doc.querySelector('.dashboard-content').style.overflow = 'visible'; }
        });
        
        const nome = currentDashboard.toUpperCase().replace('_','-');
        const periodo = ['recorrencia','recorrencia_vendedor'].includes(currentDashboard) ? '' : `_${getMonthName(currentMonth)}-${currentYear}`;
        
        const link = document.createElement('a');
        link.download = `Dashboard_${nome}${periodo}.png`;
        link.href = canvas.toDataURL('image/png');
        link.click();
    } catch(err) { 
        alert('Erro ao gerar imagem. Tente novamente.'); 
    } finally { 
        btn.innerHTML = orig; 
        btn.disabled = false; 
    }
}

function getMonthName(m){ 
    return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][m-1]; 
}

// ============================================================================
// PAINEL ADMIN - GESTÃO DE VENDEDORES
// ============================================================================

// Lista de emails administradores (AJUSTE PARA SEUS EMAILS)
const ADMIN_EMAILS = [
    'admin@cdt.com.br',
    'gestor@cdt.com.br',
    'documentostc01@gmail.com'
];

// Verificar se usuário é admin
async function checkIsAdmin() {
    try {
        if (!supabaseClient) {
            console.log('Modo admin: Supabase não disponível, liberando para testes');
            return true;
        }
        
        const { data: { user } } = await supabaseClient.auth.getUser();
        if (user && ADMIN_EMAILS.includes(user.email)) {
            console.log('Admin detectado:', user.email);
            return true;
        }
        
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get('admin') === 'true') {
            console.log('Modo admin ativado por URL');
            return true;
        }
        
        return true; // Libera para testes - em produção, mude para false
        
    } catch (e) {
        console.warn('Erro ao verificar admin:', e);
        return true; // Libera para testes
    }
}

// Inicializar botão admin
async function initAdminButton() {
    const isAdmin = await checkIsAdmin();
    const adminBtn = document.getElementById('adminBtn');
    if (adminBtn) {
        adminBtn.style.display = isAdmin ? 'inline-flex' : 'none';
        console.log('Botão admin:', isAdmin ? 'visível' : 'oculto');
    }
}

// Alternar painel admin
function toggleAdminPanel() {
    const panel = document.getElementById('adminPanel');
    const iframe = document.getElementById('adminFrame');
    
    if (panel.style.display === 'none') {
        // ✅ CORREÇÃO: API_URL já termina em /exec — apenas adiciona ?page=admin
        iframe.src = `${API_URL}?page=admin&t=${Date.now()}`;
        panel.style.display = 'block';
    } else {
        panel.style.display = 'none';
        iframe.src = '';
    }
}
