// ============================================================================
// DASHBOARD V20.10 - app.js OTIMIZADO
// MELHORIAS DE PERFORMANCE:
// 1. Cache por endpoint/mês/ano — evita re-buscar dados já carregados
// 2. Prefetch da próxima aba ao passar o mouse (hoverfetch)
// 3. Debounce nos seletores de mês/ano
// 4. Indicador de origem dos dados (cache vs. rede)
// ============================================================================

const API_URL = 'https://script.google.com/macros/s/AKfycby7E_l1q-sgkJV9oPYIdsOwjF3rJUnNjPwzSrf-jOhCwTRbk5NNLPtdF9S2320ngiI_Hw/exec';

// ── State ──────────────────────────────────────────────────────────────────
let currentDashboard = 'documentacao';
let currentMonth     = new Date().getMonth() + 1;
let currentYear      = new Date().getFullYear();

// Cache: chave = "endpoint:mes:ano"  →  { data, timestamp }
const CACHE       = new Map();
const CACHE_TTL   = 5 * 60 * 1000; // 5 minutos

// ── DOM refs ───────────────────────────────────────────────────────────────
let dashboardBtns, monthSelect, yearSelect, dashboardContent,
    refreshBtn, downloadBtn, loadingEl, lastUpdateEl, periodSelector;

// ── Debounce helper ────────────────────────────────────────────────────────
function debounce(fn, delay) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), delay); };
}

// ── Cache helpers ──────────────────────────────────────────────────────────
function cacheKey(endpoint, mes, ano) {
    const semFiltro = ['recorrencia', 'recorrencia_vendedor'];
    return semFiltro.includes(endpoint) ? endpoint : `${endpoint}:${mes}:${ano}`;
}

function getCached(key) {
    const entry = CACHE.get(key);
    if (!entry) return null;
    if (Date.now() - entry.timestamp > CACHE_TTL) { CACHE.delete(key); return null; }
    return entry.data;
}

function setCache(key, data) {
    CACHE.set(key, { data, timestamp: Date.now() });
}

// ── Build URL ──────────────────────────────────────────────────────────────
function buildUrl(endpoint, mes, ano) {
    const semFiltro = ['recorrencia', 'recorrencia_vendedor'];
    let url = `${API_URL}?endpoint=${endpoint}`;
    if (!semFiltro.includes(endpoint)) url += `&mes=${mes}&ano=${ano}`;
    return url;
}

// ── Init ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function () {
    dashboardBtns    = document.querySelectorAll('.dashboard-btn');
    monthSelect      = document.getElementById('monthSelect');
    yearSelect       = document.getElementById('yearSelect');
    dashboardContent = document.getElementById('dashboardContent');
    refreshBtn       = document.getElementById('refreshBtn');
    downloadBtn      = document.getElementById('downloadBtn');
    loadingEl        = document.getElementById('loading');
    lastUpdateEl     = document.getElementById('lastUpdate');
    periodSelector   = document.getElementById('periodSelector');

    initializeYearSelect();
    setCurrentPeriod();
    setupEventListeners();
    loadDashboard();
    updateLastUpdateTime();
});

function initializeYearSelect() {
    const cy = new Date().getFullYear();
    for (let year = cy - 2; year <= cy + 2; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearSelect.appendChild(option);
    }
    yearSelect.value = cy;
}

function setCurrentPeriod() {
    monthSelect.value = currentMonth;
    yearSelect.value  = currentYear;
}

function setupEventListeners() {
    const debouncedLoad = debounce(loadDashboard, 300);

    dashboardBtns.forEach(btn => {
        btn.addEventListener('click', function () {
            if (this.dataset.dashboard === currentDashboard) return; // já na aba
            dashboardBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            currentDashboard = this.dataset.dashboard;
            togglePeriodSelector();
            loadDashboard();
        });

        // Prefetch ao hover
        btn.addEventListener('mouseenter', function () {
            const ep  = this.dataset.dashboard;
            const key = cacheKey(ep, currentMonth, currentYear);
            if (!getCached(key)) {
                prefetchDashboard(ep);
            }
        });
    });

    monthSelect.addEventListener('change', function () {
        currentMonth = parseInt(this.value);
        debouncedLoad();
    });

    yearSelect.addEventListener('change', function () {
        currentYear = parseInt(this.value);
        debouncedLoad();
    });

    refreshBtn.addEventListener('click', () => {
        // Forçar refresh: apagar cache da aba atual
        const key = cacheKey(currentDashboard, currentMonth, currentYear);
        CACHE.delete(key);
        loadDashboard();
    });

    downloadBtn.addEventListener('click', exportPage);
}

// Prefetch silencioso
async function prefetchDashboard(endpoint) {
    const key = cacheKey(endpoint, currentMonth, currentYear);
    if (getCached(key)) return;
    try {
        const url  = buildUrl(endpoint, currentMonth, currentYear);
        const res  = await fetch(url);
        const data = await res.json();
        if (data.status === 'success') setCache(key, data.data);
    } catch (_) { /* silencioso */ }
}

function togglePeriodSelector() {
    const semFiltro = ['recorrencia_vendedor'];
    if (periodSelector) {
        periodSelector.style.display = semFiltro.includes(currentDashboard) ? 'none' : 'flex';
    }
}

function updateLastUpdateTime() {
    if (lastUpdateEl) lastUpdateEl.textContent = new Date().toLocaleString('pt-BR');
}

// ── Main load ──────────────────────────────────────────────────────────────
async function loadDashboard() {
    const key    = cacheKey(currentDashboard, currentMonth, currentYear);
    const cached = getCached(key);

    if (cached) {
        // Dados em cache — renderiza imediatamente sem spinner
        renderDashboardData(cached);
        updateLastUpdateTime();
        return;
    }

    showLoading();
    updateLastUpdateTime();

    try {
        const url      = buildUrl(currentDashboard, currentMonth, currentYear);
        const response = await fetch(url);
        const data     = await response.json();

        if (data.status === 'success') {
            setCache(key, data.data);
            renderDashboardData(data.data);
        } else {
            showError(data.error || 'Erro ao carregar dados');
        }
    } catch (error) {
        console.error('Fetch error:', error);
        showError('Erro de conexão: ' + error.message);
    }
}

function renderDashboardData(data) {
    // Injetar no estado global para funções de render
    window._dashboardData = data;
    renderDashboard();
}

// ── Render router ──────────────────────────────────────────────────────────
function renderDashboard() {
    hideLoading();
    const data = window._dashboardData;

    switch (currentDashboard) {
        case 'documentacao':         renderDocumentacaoDashboard(data);        break;
        case 'app':                  renderAppDashboard(data);                 break;
        case 'adimplencia':          renderAdimplenciaDashboard(data);         break;
        case 'recorrencia':          renderRecorrenciaDashboard(data);         break;
        case 'recorrencia_vendedor': renderRecorrenciaVendedorDashboard(data); break;
        case 'refuturiza':           renderRefuturizaDashboard(data);          break;
        default: showError('Dashboard não encontrado: ' + currentDashboard);
    }
}

// ============================================================================
// DASHBOARD: DOCUMENTAÇÃO / VENDAS
// ============================================================================
function renderDocumentacaoDashboard(dashboardData) {
    const { geral, vendasLoja, vendasWeb, consultores, mes, ano } = dashboardData;

    const percentGeral = calcPercent(geral.aprovados, geral.total);
    const percentLoja  = calcPercent(vendasLoja.aprovados, vendasLoja.total);
    const percentWeb   = calcPercent(vendasWeb.aprovados, vendasWeb.total);

    const consultantsBySector = groupBySector(consultores);
    const sectorOrder = ['VENDAS', 'RECEPCAO', 'REFILIACAO', 'WEB SITE', 'TELEVENDAS', 'OUTROS'];
    const sortedSectors = sortSectors(Object.keys(consultantsBySector), sectorOrder);

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-folder" style="color: var(--primary);"></i>
            Dashboard de Vendas — ${mes} ${ano}
        </h2>

        <div class="main-cards">
            ${cardPrincipalDoc('Total de Vendas', 'fas fa-chart-bar', geral, percentGeral)}
            ${cardPrincipalDoc('Vendas Loja', 'fas fa-store', vendasLoja, percentLoja)}
            ${cardPrincipalDoc('Vendas Web/Tele', 'fas fa-globe', vendasWeb, percentWeb)}
        </div>

        <h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--primary);"></i> Desempenho por Setor</h3>

        ${sortedSectors.map(sector => {
            const list = consultantsBySector[sector];
            const tot  = list.reduce((s, c) => s + c.total, 0);
            const aprov= list.reduce((s, c) => s + c.aprovados, 0);
            const pct  = calcPercent(aprov, tot);
            return `
            <div class="sector-card">
                <div class="sector-header">
                    <div class="sector-title">
                        <i class="${getSectorIcon(sector)}"></i> ${sector}
                        <span class="sector-count">${list.length} consultor${list.length !== 1 ? 'es' : ''}</span>
                    </div>
                    <div class="metric-percent ${getPercentClass(pct)}">${pct}% aprovados</div>
                </div>
                <div class="consultant-grid">
                    ${list.sort((a,b)=>b.total-a.total).map(c => {
                        const p = calcPercent(c.aprovados, c.total);
                        return `
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
                        </div>`;
                    }).join('')}
                </div>
            </div>`;
        }).join('')}
    `;
    dashboardContent.innerHTML = html;
}

function cardPrincipalDoc(titulo, icon, d, pct) {
    return `
    <div class="card card-doc">
        <div class="card-header">
            <div class="card-title">${titulo}</div>
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
    </div>`;
}

// ============================================================================
// DASHBOARD: APP
// ============================================================================
function renderAppDashboard(dashboardData) {
    const { geral, appLoja, appWeb, consultores, consultorasRetencao, mes, ano } = dashboardData;

    const percentGeral = calcPercent(geral.sim, geral.total);
    const percentLoja  = calcPercent(appLoja.sim, appLoja.total);
    const percentWeb   = calcPercent(appWeb.sim, appWeb.total);

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-mobile-alt" style="color: var(--teal);"></i>
            Dashboard App — ${mes} ${ano}
        </h2>

        <div class="main-cards">
            ${cardPrincipalApp('App — Total Geral', 'fas fa-chart-pie', geral, percentGeral)}
            ${cardPrincipalApp('App — Loja', 'fas fa-store', appLoja, percentLoja)}
            ${cardPrincipalApp('App — Web/Tele', 'fas fa-globe', appWeb, percentWeb)}
        </div>

        ${consultorasRetencao && consultorasRetencao.length > 0 ? `
        <div class="retention-section">
            <div class="retention-header">
                <i class="fas fa-crown"></i>
                <h3>Consultoras de Retenção (Dados Separados)</h3>
            </div>
            <div class="consultant-grid">
                ${consultorasRetencao.map(c => {
                    const p = calcPercent(c.sim, c.total);
                    return `
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
                    </div>`;
                }).join('')}
            </div>
        </div>` : ''}

        <h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--teal);"></i> Desempenho por Setor</h3>
        ${renderAppBySector(consultores)}
    `;
    dashboardContent.innerHTML = html;
}

function cardPrincipalApp(titulo, icon, d, pct) {
    return `
    <div class="card card-app">
        <div class="card-header">
            <div class="card-title">${titulo}</div>
            <div class="card-icon"><i class="${icon}"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItem('Total Clientes', d.total)}
            ${metricItem('Com App (SIM)', d.sim, 'var(--success)')}
            ${metricItem('Sem App (NÃO)', d.nao, 'var(--danger)')}
            ${metricItem('Cancelados', d.cancelado || 0, 'var(--gray)')}
            ${metricPercent('% Com App', pct)}
        </div>
    </div>`;
}

function renderAppBySector(consultores) {
    const regular = consultores.filter(c => !c.origem || c.origem !== 'retencao');
    const bySector = groupBySector(regular);
    const sectorOrder = ['VENDAS', 'RECEPCAO', 'REFILIACAO', 'WEB SITE', 'TELEVENDAS', 'OUTROS'];
    const sorted = sortSectors(Object.keys(bySector), sectorOrder);

    if (sorted.length === 0) return '<p style="text-align:center;color:var(--gray);padding:40px;">Nenhum consultor regular encontrado</p>';

    return sorted.map(sector => {
        const list = bySector[sector];
        const tot  = list.reduce((s,c)=>s+c.total,0);
        const sim  = list.reduce((s,c)=>s+(c.sim||0),0);
        const pct  = calcPercent(sim, tot);
        return `
        <div class="sector-card">
            <div class="sector-header">
                <div class="sector-title">
                    <i class="${getSectorIcon(sector)}"></i> ${sector}
                    <span class="sector-count">${list.length} consultor${list.length!==1?'es':''}</span>
                </div>
                <div class="metric-percent ${getPercentClass(pct)}">${pct}% com app</div>
            </div>
            <div class="consultant-grid">
                ${list.sort((a,b)=>b.total-a.total).map(c => {
                    const p = calcPercent(c.sim||0, c.total);
                    return `
                    <div class="consultant-card">
                        <div class="consultant-header">
                            <div class="consultant-name">${c.nome}</div>
                            <div class="consultant-sector ${getSectorClass(sector)}">${sector}</div>
                        </div>
                        <div class="metric-grid">
                            ${metricItem('Total', c.total)}
                            ${metricItem('Com App', c.sim||0, 'var(--success)')}
                            ${metricItem('Sem App', c.nao||0, 'var(--danger)')}
                            ${metricItem('Cancelados', c.cancelado||0, 'var(--gray)')}
                            ${metricPercent('% Com App', p)}
                        </div>
                    </div>`;
                }).join('')}
            </div>
        </div>`;
    }).join('');
}

// ============================================================================
// DASHBOARD: ADIMPLÊNCIA
// ============================================================================
function renderAdimplenciaDashboard(dashboardData) {
    const { geral, consultores, mes, ano } = dashboardData;

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-credit-card" style="color:var(--success);"></i>
            Dashboard Adimplência — ${mes} ${ano}
        </h2>

        <div class="card card-adim" style="max-width:600px;margin:0 auto 36px;">
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

        <h3 class="section-title"><i class="fas fa-user-tie" style="color:var(--primary);"></i> Desempenho por Consultor</h3>

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
                ${consultores.map(c => `
                <tr>
                    <td><strong>${c.nome}</strong></td>
                    <td>${c.totalTrocas}</td>
                    <td style="color:var(--success);">${c.mensOk}</td>
                    <td style="color:var(--warning);">${c.mensAberto}</td>
                    <td style="color:var(--danger);">${c.mensAtraso}</td>
                    <td style="color:var(--success);">${c.aprovados}</td>
                    <td style="color:var(--danger);">${c.pendentes}</td>
                    <td>${c.totalBi}</td>
                    <td style="color:var(--danger);">${c.foraBi}</td>
                    <td><span class="metric-percent ${getPercentClass(c.percentualAprovado)}">${c.percentualAprovado}%</span></td>
                </tr>`).join('')}
            </tbody>
        </table>
        </div>
    `;
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: RECORRÊNCIA
// ============================================================================
function renderRecorrenciaDashboard(dashboardData) {
    const { retencao, refiliacao, periodo } = dashboardData;

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-redo" style="color:var(--warning);"></i>
            Dashboard Recorrência — ${periodo.atual}
        </h2>

        <div style="background:rgba(245,158,11,0.08);padding:14px 18px;border-radius:12px;margin-bottom:28px;border-left:4px solid var(--warning);">
            <p style="margin:0;color:#92400e;font-weight:600;font-size:0.9rem;">
                <i class="fas fa-info-circle"></i>
                Período atual: ${periodo.atual} | Histórico: ${periodo.historico.join(', ')}
            </p>
        </div>

        <h3 class="section-title" style="border-bottom:2px solid var(--primary);">
            <i class="fas fa-crown" style="color:var(--primary);"></i> Retenção
        </h3>
        <div class="consultant-grid" style="margin-bottom:40px;">
            ${Object.keys(retencao).map(key => {
                const c = retencao[key];
                const pAtual = calcPercent(c.atual.retençõesOK, c.atual.totalRetidosFinal);
                const pTotal = calcPercent(c.total3Meses.totalOK||0, c.total3Meses.totalRetidosFinal);
                return `
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
                        ${metricPercent('% OK', pAtual)}
                    </div>
                    <h4 class="sub-section-title"><i class="fas fa-chart-line"></i> Total 3 Meses Anteriores</h4>
                    <div class="metric-grid">
                        ${metricItem('Total Retido', c.total3Meses.totalRetido)}
                        ${metricItem('Cancelados', c.total3Meses.cancelado||0, 'var(--danger)')}
                        ${metricItem('Total OK (OK+Aberto)', c.total3Meses.totalOK||0, 'var(--success)')}
                        ${metricItem('Em Atraso', c.total3Meses.emAtraso, 'var(--warning)')}
                        ${metricPercent('% OK', pTotal)}
                    </div>
                </div>`;
            }).join('')}
        </div>

        <h3 class="section-title" style="border-bottom:2px solid var(--warning);">
            <i class="fas fa-user-plus" style="color:var(--warning);"></i> Refiliação
        </h3>
        <div class="consultant-grid">
            ${Object.keys(refiliacao).map(key => {
                const c = refiliacao[key];
                const pTotal = calcPercent(c.total3Meses.totalOK||0, c.total3Meses.totalRetidosFinal);
                return `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${key} (REFILIAÇÃO)</div>
                        <div class="consultant-sector sector-refiliacao">REFILIAÇÃO</div>
                    </div>
                    <h4 class="sub-section-title"><i class="fas fa-chart-line"></i> Total 3 Meses Anteriores</h4>
                    <div class="metric-grid">
                        ${metricItem('Total Refiliados', c.total3Meses.totalRetido)}
                        ${metricItem('Cancelados', c.total3Meses.cancelado||0, 'var(--danger)')}
                        ${metricItem('Total OK (OK+Aberto)', c.total3Meses.totalOK||0, 'var(--success)')}
                        ${metricItem('Em Atraso', c.total3Meses.emAtraso, 'var(--warning)')}
                        ${metricPercent('% OK', pTotal)}
                    </div>
                </div>`;
            }).join('')}
        </div>
    `;
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: RECORRÊNCIA VENDEDOR
// ============================================================================
function renderRecorrenciaVendedorDashboard(dashboardData) {
    const { geral, consultores, dadosPorSetor, totalConsultores, totalRegistros, mesAtual, anoAtual } = dashboardData;

    if (!geral || !consultores) { showError('Dados de Recorrência Vendedor não encontrados'); return; }

    const percentGeral = calcPercent(geral.totalOk + geral.totalEmAberto, geral.totalVendasPromocao);

    const setores = [
        { nome: 'VENDAS',     cor: 'var(--primary)',  icone: 'fas fa-shopping-cart' },
        { nome: 'REFILIACAO', cor: 'var(--warning)',   icone: 'fas fa-user-plus' },
        { nome: 'RECEPCAO',   cor: 'var(--teal)',      icone: 'fas fa-headset' },
    ];

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-handshake" style="color:var(--teal);"></i>
            Recorrência Vendedor — Dados Completos
        </h2>

        <div style="background:var(--teal-light);padding:14px 18px;border-radius:12px;margin-bottom:28px;border-left:4px solid var(--teal);">
            <p style="margin:0;color:#065f46;font-weight:600;font-size:0.9rem;">
                <i class="fas fa-database"></i>
                Análise de <strong>${totalRegistros || consultores.reduce((s,c)=>s+c.totalVendasPromocao,0)}</strong> registros |
                <strong>${totalConsultores}</strong> consultores |
                Dados históricos completos (sem filtro de mês)
            </p>
        </div>

        <div class="card" style="max-width:600px;margin:0 auto 36px;border-top:3px solid var(--teal);">
            <div class="card-header">
                <div class="card-title">📈 Total Geral — Todos os Dados</div>
                <div class="card-icon" style="background:var(--teal-light);color:var(--teal);">
                    <i class="fas fa-chart-bar"></i>
                </div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total Vendas', geral.totalVendasPromocao)}
                ${metricItem('OK', geral.totalOk, 'var(--success)')}
                ${metricItem('Em Aberto', geral.totalEmAberto, 'var(--warning)')}
                ${metricItem('Atraso', geral.totalAtraso, 'var(--danger)')}
                ${metricItem('Outros', geral.totalOutros||0, 'var(--gray)')}
                ${metricPercent('% OK + Aberto', percentGeral)}
            </div>
        </div>

        ${setores.map(setorInfo => {
            const lista = consultores.filter(c => c.setor === setorInfo.nome);
            const totSetor = lista.reduce((s,c)=>s+c.totalVendasPromocao,0);
            const okSetor  = lista.reduce((s,c)=>s+(c.totalOk||0)+(c.totalEmAberto||0),0);
            const pctSetor = calcPercent(okSetor, totSetor);
            if (lista.length === 0) return '';
            return `
            <div class="sector-card">
                <div class="sector-header">
                    <div class="sector-title">
                        <i class="${setorInfo.icone}" style="color:${setorInfo.cor};"></i>
                        <span style="color:${setorInfo.cor};">${setorInfo.nome}</span>
                        <span class="sector-count">${lista.length} consultor${lista.length!==1?'es':''}</span>
                    </div>
                    <div class="metric-percent ${getPercentClass(pctSetor)}">${pctSetor}% OK+Aberto</div>
                </div>
                <div class="consultant-grid">
                    ${lista.sort((a,b)=>b.totalVendasPromocao-a.totalVendasPromocao).map(c => {
                        const p = calcPercent((c.totalOk||0)+(c.totalEmAberto||0), c.totalVendasPromocao);
                        return `
                        <div class="consultant-card">
                            <div class="consultant-header">
                                <div class="consultant-name">${c.nome}</div>
                                <div class="consultant-sector ${getSectorClass(c.setor)}">${c.setor}</div>
                            </div>
                            <div class="metric-grid">
                                ${metricItem('Total Vendas', c.totalVendasPromocao)}
                                ${metricItem('OK', c.totalOk||0, 'var(--success)')}
                                ${metricItem('Em Aberto', c.totalEmAberto||0, 'var(--warning)')}
                                ${metricItem('Atraso', c.totalAtraso||0, 'var(--danger)')}
                                ${metricItem('Outros', c.totalOutros||0, 'var(--gray)')}
                                ${metricPercent('% OK+Aberto', p)}
                            </div>
                        </div>`;
                    }).join('')}
                </div>
            </div>`;
        }).join('')}

        <div class="card" style="margin-top:36px;background:linear-gradient(135deg,var(--primary),var(--primary-dark));color:white;">
            <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2);">
                <div class="card-title" style="color:white;">📊 Resumo Final</div>
                <div class="card-icon" style="background:rgba(255,255,255,0.2);">
                    <i class="fas fa-clipboard-list"></i>
                </div>
            </div>
            <div class="metric-grid">
                ${metricItemWhite('Total Consultores', totalConsultores||consultores.length)}
                ${metricItemWhite('Total Vendas', geral.totalVendasPromocao)}
                ${metricItemWhite('Média por Consultor', consultores.length > 0 ? Math.round(geral.totalVendasPromocao/consultores.length) : 0)}
                ${metricItemWhite('Melhor % Individual', consultores.length > 0 ?
                    Math.max(...consultores.map(c=>calcPercent((c.totalOk||0)+(c.totalEmAberto||0),c.totalVendasPromocao)))+'%' : '0%')}
            </div>
        </div>
    `;
    dashboardContent.innerHTML = html;
}

// ============================================================================
// DASHBOARD: REFUTURIZA
// ============================================================================
function renderRefuturizaDashboard(dashboardData) {
    const { geral, consultores, mes, ano } = dashboardData;

    if (!geral || !consultores) { showError('Dados do Refuturiza não encontrados'); return; }

    const percentGeral = calcPercent(geral.comLigacao, geral.total);

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-book" style="color:var(--info);"></i>
            Dashboard Refuturiza — ${mes} ${ano}
        </h2>

        <div class="card card-refut" style="max-width:500px;margin:0 auto 36px;">
            <div class="card-header">
                <div class="card-title">Refuturiza — Total da Loja</div>
                <div class="card-icon"><i class="fas fa-book-open"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total', geral.total||0)}
                ${metricItem('Com Ligação', geral.comLigacao||0, 'var(--success)')}
                ${metricItem('Sem Ligação', geral.semLigacao||0, 'var(--danger)')}
                ${metricItem('Cancelados', geral.cancelado||0, 'var(--gray)')}
                ${metricPercent('% Com Ligação', percentGeral)}
            </div>
        </div>

        ${consultores && consultores.length > 0 ? `
        <h3 class="section-title"><i class="fas fa-users" style="color:var(--info);"></i> Desempenho por Consultor (${consultores.length})</h3>
        <div class="consultant-grid">
            ${consultores.map(c => {
                const p = calcPercent(c.comLigacao||0, c.total);
                return `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome}</div>
                        <div class="consultant-sector" style="background:rgba(14,165,233,0.1);color:var(--info);">REFUTURIZA</div>
                    </div>
                    <div class="metric-grid">
                        ${metricItem('Total', c.total||0)}
                        ${metricItem('Com Ligação', c.comLigacao||0, 'var(--success)')}
                        ${metricItem('Sem Ligação', c.semLigacao||0, 'var(--danger)')}
                        ${metricItem('Cancelados', c.cancelado||0, 'var(--gray)')}
                        ${metricPercent('% Com Ligação', p)}
                    </div>
                </div>`;
            }).join('')}
        </div>
        <div class="card" style="margin-top:36px;background:linear-gradient(90deg,var(--info),#3b82f6);color:white;">
            <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2);">
                <div class="card-title" style="color:white;">📊 Resumo Final</div>
                <div class="card-icon" style="background:rgba(255,255,255,0.2);"><i class="fas fa-graduation-cap"></i></div>
            </div>
            <div class="metric-grid">
                ${metricItemWhite('Total Consultores', consultores.length)}
                ${metricItemWhite('Total Cursos', geral.total||0)}
                ${metricItemWhite('Média por Consultor', consultores.length>0?Math.round((geral.total||0)/consultores.length):0)}
                ${metricItemWhite('Taxa de Contato', percentGeral+'%')}
            </div>
        </div>` : `
        <div class="error-message" style="text-align:center;padding:40px;background:var(--gray-light);border:none;color:var(--text-muted);">
            <i class="fas fa-info-circle" style="font-size:2.5rem;margin-bottom:16px;display:block;"></i>
            <h3>Nenhum consultor com cursos neste período</h3>
            <p>Não foram encontrados dados para ${mes} ${ano}.</p>
        </div>`}
    `;
    dashboardContent.innerHTML = html;
}

// ============================================================================
// HELPERS
// ============================================================================
function metricItem(label, value, color) {
    const style = color ? `style="color:${color};"` : '';
    return `
    <div class="metric-item">
        <div class="metric-label">${label}</div>
        <div class="metric-value" ${style}>${value ?? 0}</div>
    </div>`;
}

function metricPercent(label, value) {
    const cls = getPercentClass(typeof value === 'number' ? value : parseInt(value));
    return `
    <div class="metric-item">
        <div class="metric-label">${label}</div>
        <div class="metric-percent ${cls}">${value}%</div>
    </div>`;
}

function metricItemWhite(label, value) {
    return `
    <div class="metric-item">
        <div class="metric-label" style="color:rgba(255,255,255,0.75);">${label}</div>
        <div class="metric-value" style="color:white;">${value}</div>
    </div>`;
}

function calcPercent(numerator, denominator) {
    if (!denominator || denominator === 0) return 0;
    return Math.round((numerator / denominator) * 100);
}

function groupBySector(consultores) {
    const map = {};
    consultores.forEach(c => {
        const s = c.setor || 'OUTROS';
        if (!map[s]) map[s] = [];
        map[s].push(c);
    });
    return map;
}

function sortSectors(keys, order) {
    return keys.sort((a, b) => {
        const ia = order.indexOf(a), ib = order.indexOf(b);
        if (ia === -1 && ib === -1) return a.localeCompare(b);
        if (ia === -1) return 1;
        if (ib === -1) return -1;
        return ia - ib;
    });
}

function getPercentClass(percent) {
    const p = typeof percent === 'number' ? percent : parseInt(percent) || 0;
    if (p >= 90) return 'percent-high';
    if (p >= 80) return 'percent-medium';
    return 'percent-low';
}

function getSectorClass(sector) {
    if (!sector) return 'sector-outros';
    switch (sector.toUpperCase()) {
        case 'VENDAS':     return 'sector-vendas';
        case 'RECEPCAO':   return 'sector-recepcao';
        case 'REFILIACAO': return 'sector-refiliacao';
        case 'WEB SITE':
        case 'WEB':        return 'sector-web';
        case 'TELEVENDAS': return 'sector-televendas';
        case 'RETENÇÃO':
        case 'RETENCAO':   return 'sector-retencao';
        default:           return 'sector-outros';
    }
}

function getSectorIcon(sector) {
    if (!sector) return 'fas fa-users';
    switch (sector.toUpperCase()) {
        case 'VENDAS':     return 'fas fa-shopping-cart';
        case 'RECEPCAO':   return 'fas fa-headset';
        case 'REFILIACAO': return 'fas fa-user-plus';
        case 'WEB SITE':
        case 'WEB':        return 'fas fa-globe';
        case 'TELEVENDAS': return 'fas fa-phone-alt';
        case 'RETENÇÃO':
        case 'RETENCAO':   return 'fas fa-crown';
        default:           return 'fas fa-users';
    }
}

// ── Loading / Error ────────────────────────────────────────────────────────
function showLoading() {
    if (loadingEl) loadingEl.style.display = 'flex';
    Array.from(dashboardContent.children).forEach(child => {
        if (child !== loadingEl) child.remove();
    });
}

function hideLoading() {
    if (loadingEl) loadingEl.style.display = 'none';
}

function showError(message) {
    hideLoading();
    dashboardContent.innerHTML = `
        <div class="error-message">
            <i class="fas fa-exclamation-triangle" style="font-size:2rem;margin-bottom:12px;display:block;"></i>
            <h3 style="margin-bottom:8px;">Erro ao carregar dados</h3>
            <p>${message}</p>
            <button class="btn btn-success" onclick="loadDashboard()" style="margin-top:16px;background:var(--danger);color:white;">
                <i class="fas fa-redo"></i> Tentar Novamente
            </button>
        </div>
    `;
}

function updateLastUpdateTime() {
    if (lastUpdateEl) lastUpdateEl.textContent = new Date().toLocaleString('pt-BR');
}

// ── Export ─────────────────────────────────────────────────────────────────
async function exportPage() {
    const btn = document.getElementById('downloadBtn');
    const originalHtml = btn.innerHTML;

    // Feedback visual no botão
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando imagem...';
    btn.disabled = true;

    try {
        const container = document.querySelector('.container');

        const canvas = await html2canvas(container, {
            scale: 2,              // resolução 2x — boa para WhatsApp/grupos
            useCORS: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: -window.scrollY,
            windowWidth: container.scrollWidth,
            windowHeight: container.scrollHeight,
            onclone: (doc) => {
                // Garante que o conteúdo completo aparece na captura
                doc.querySelector('.dashboard-content').style.overflow = 'visible';
            }
        });

        // Nome do arquivo com aba e período
        const nomeDashboard = currentDashboard.toUpperCase().replace('_', '-');
        const periodo = ['recorrencia', 'recorrencia_vendedor'].includes(currentDashboard)
            ? ''
            : `_${getMonthName(currentMonth)}-${currentYear}`;
        const filename = `Dashboard_${nomeDashboard}${periodo}.png`;

        // Download automático
        const link = document.createElement('a');
        link.download = filename;
        link.href = canvas.toDataURL('image/png');
        link.click();

    } catch (err) {
        console.error('Erro ao gerar imagem:', err);
        alert('Erro ao gerar imagem. Tente novamente.');
    } finally {
        btn.innerHTML = originalHtml;
        btn.disabled = false;
    }
}

function getMonthName(month) {
    return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][month - 1];
}
