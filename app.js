// ============================================================================
// DASHBOARD V20.10 - app.js CORRIGIDO
// CORREÇÕES:
// 1. Adicionado suporte ao endpoint 'recorrencia_vendedor' (botão + render)
// 2. Corrigida condição de envio de mes/ano (excluir recorrencia E recorrencia_vendedor)
// 3. Corrigido showLoading() para não remover o elemento do DOM
// 4. Adicionada função renderRecorrenciaVendedorDashboard() completa
// ============================================================================

const API_URL = 'https://script.google.com/macros/s/AKfycby7E_l1q-sgkJV9oPYIdsOwjF3rJUnNjPwzSrf-jOhCwTRbk5NNLPtdF9S2320ngiI_Hw/exec';

// State
let currentDashboard = 'documentacao';
let currentMonth = new Date().getMonth() + 1;
let currentYear = new Date().getFullYear();
let dashboardData = {};

// DOM Elements (inicializados após DOMContentLoaded)
let dashboardBtns, monthSelect, yearSelect, dashboardContent,
    refreshBtn, downloadBtn, loadingEl, lastUpdateEl, periodSelector;

document.addEventListener('DOMContentLoaded', function () {
    dashboardBtns   = document.querySelectorAll('.dashboard-btn');
    monthSelect     = document.getElementById('monthSelect');
    yearSelect      = document.getElementById('yearSelect');
    dashboardContent= document.getElementById('dashboardContent');
    refreshBtn      = document.getElementById('refreshBtn');
    downloadBtn     = document.getElementById('downloadBtn');
    loadingEl       = document.getElementById('loading');
    lastUpdateEl    = document.getElementById('lastUpdate');
    periodSelector  = document.getElementById('periodSelector');

    initializeYearSelect();
    setCurrentPeriod();
    setupEventListeners();
    loadDashboard();
    updateLastUpdateTime();
});

function initializeYearSelect() {
    const currentYear = new Date().getFullYear();
    for (let year = currentYear - 2; year <= currentYear + 2; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearSelect.appendChild(option);
    }
    yearSelect.value = currentYear;
}

function setCurrentPeriod() {
    monthSelect.value = currentMonth;
    yearSelect.value  = currentYear;
}

function setupEventListeners() {
    dashboardBtns.forEach(btn => {
        btn.addEventListener('click', function () {
            dashboardBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            currentDashboard = this.dataset.dashboard;
            // Ocultar seletor de mês/ano para dashboards sem filtro temporal
            togglePeriodSelector();
            loadDashboard();
        });
    });

    monthSelect.addEventListener('change', function () {
        currentMonth = parseInt(this.value);
        loadDashboard();
    });

    yearSelect.addEventListener('change', function () {
        currentYear = parseInt(this.value);
        loadDashboard();
    });

    refreshBtn.addEventListener('click', loadDashboard);
    downloadBtn.addEventListener('click', exportPage);
}

// Ocultar o seletor de período para dashboards sem filtro de mês
function togglePeriodSelector() {
    const semFiltro = ['recorrencia_vendedor'];
    if (periodSelector) {
        periodSelector.style.display = semFiltro.includes(currentDashboard) ? 'none' : 'flex';
    }
}

function updateLastUpdateTime() {
    if (lastUpdateEl) lastUpdateEl.textContent = new Date().toLocaleString('pt-BR');
}

async function loadDashboard() {
    showLoading();
    updateLastUpdateTime();

    try {
        // Dashboards que NÃO usam filtro de mês/ano
        const semFiltroTemporal = ['recorrencia', 'recorrencia_vendedor'];

        let url = `${API_URL}?endpoint=${currentDashboard}`;
        if (!semFiltroTemporal.includes(currentDashboard)) {
            url += `&mes=${currentMonth}&ano=${currentYear}`;
        }

        console.log('Fetching:', url);
        const response = await fetch(url);
        const data = await response.json();
        console.log('API Response:', data);

        if (data.status === 'success') {
            dashboardData = data.data;
            renderDashboard();
        } else {
            showError(data.error || 'Erro ao carregar dados');
        }
    } catch (error) {
        console.error('Fetch error:', error);
        showError('Erro de conexão: ' + error.message);
    }
}

function renderDashboard() {
    hideLoading();

    switch (currentDashboard) {
        case 'documentacao':         renderDocumentacaoDashboard();        break;
        case 'app':                  renderAppDashboard();                 break;
        case 'adimplencia':          renderAdimplenciaDashboard();         break;
        case 'recorrencia':          renderRecorrenciaDashboard();         break;
        case 'recorrencia_vendedor': renderRecorrenciaVendedorDashboard(); break;
        case 'refuturiza':           renderRefuturizaDashboard();          break;
        default:
            showError('Dashboard não encontrado: ' + currentDashboard);
    }
}

// ============================================================================
// DASHBOARD: DOCUMENTAÇÃO
// ============================================================================
function renderDocumentacaoDashboard() {
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
function renderAppDashboard() {
    const { geral, appLoja, appWeb, consultores, consultorasRetencao, mes, ano } = dashboardData;

    const percentGeral = calcPercent(geral.sim, geral.total);
    const percentLoja  = calcPercent(appLoja.sim, appLoja.total);
    const percentWeb   = calcPercent(appWeb.sim, appWeb.total);

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-mobile-alt" style="color: var(--secondary);"></i>
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
                    <div class="consultant-card" style="border-left:4px solid #f59e0b;">
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

        <h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--secondary);"></i> Desempenho por Setor</h3>
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
function renderAdimplenciaDashboard() {
    const { geral, consultores, mes, ano } = dashboardData;

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-credit-card" style="color:var(--success);"></i>
            Dashboard Adimplência — ${mes} ${ano}
        </h2>

        <div class="card card-adim" style="max-width:600px;margin:0 auto 40px;">
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
function renderRecorrenciaDashboard() {
    const { retencao, refiliacao, periodo } = dashboardData;

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-redo" style="color:var(--warning);"></i>
            Dashboard Recorrência — ${periodo.atual}
        </h2>

        <div style="background:#fef3c7;padding:15px;border-radius:10px;margin-bottom:30px;">
            <p style="margin:0;color:#92400e;font-weight:500;">
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
// DASHBOARD: RECORRÊNCIA VENDEDOR ← NOVO / CORRIGIDO
// ============================================================================
function renderRecorrenciaVendedorDashboard() {
    const { geral, consultores, dadosPorSetor, totalConsultores, totalRegistros, mesAtual, anoAtual } = dashboardData;

    if (!geral || !consultores) {
        showError('Dados de Recorrência Vendedor não encontrados');
        return;
    }

    const percentGeral = calcPercent(geral.totalOk + geral.totalEmAberto, geral.totalVendasPromocao);

    const setores = [
        { nome: 'VENDAS',     cor: '#1e3a8a', icone: 'fas fa-shopping-cart' },
        { nome: 'REFILIACAO', cor: '#7c3aed', icone: 'fas fa-user-plus' },
        { nome: 'RECEPCAO',   cor: '#059669', icone: 'fas fa-headset' },
    ];

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-handshake" style="color:#0d9488;"></i>
            Recorrência Vendedor — Dados Completos
        </h2>

        <div style="background:#e0f2f1;padding:15px;border-radius:10px;margin-bottom:30px;border-left:5px solid #0d9488;">
            <p style="margin:0;color:#065f46;font-weight:500;">
                <i class="fas fa-database"></i>
                Análise de <strong>${totalRegistros || consultores.reduce((s,c)=>s+c.totalVendasPromocao,0)}</strong> registros |
                <strong>${totalConsultores}</strong> consultores |
                Dados históricos completos (sem filtro de mês)
            </p>
        </div>

        <!-- Card Geral -->
        <div class="card" style="max-width:600px;margin:0 auto 40px;border-top:4px solid #0d9488;">
            <div class="card-header">
                <div class="card-title">📈 Total Geral — Todos os Dados</div>
                <div class="card-icon" style="background:rgba(13,148,136,0.1);color:#0d9488;">
                    <i class="fas fa-chart-bar"></i>
                </div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total Vendas', geral.totalVendasPromocao)}
                ${metricItem('OK', geral.totalOk, 'var(--success)')}
                ${metricItem('Em Aberto', geral.totalEmAberto, 'var(--warning)')}
                ${metricItem('Atraso', geral.totalAtraso, 'var(--danger)')}
                ${metricItem('Outros', geral.totalOutros||0, 'var(--gray)')}
                ${metricPercent('% OK + Aberto', percentGeral, '#0d9488')}
            </div>
        </div>

        <!-- Por Setor -->
        ${setores.map(setorInfo => {
            // Usar filtro direto nos consultores como fallback (mais confiável)
            const lista = consultores.filter(c => c.setor === setorInfo.nome);
            const totSetor = lista.reduce((s,c)=>s+c.totalVendasPromocao,0);
            const okSetor  = lista.reduce((s,c)=>s+(c.totalOk||0)+(c.totalEmAberto||0),0);
            const pctSetor = calcPercent(okSetor, totSetor);

            if (lista.length === 0) return '';

            return `
            <div class="sector-card" style="border-top:4px solid ${setorInfo.cor};">
                <div class="sector-header" style="background:${setorInfo.cor}15;">
                    <div class="sector-title">
                        <i class="${setorInfo.icone}" style="color:${setorInfo.cor};"></i>
                        <span style="color:${setorInfo.cor};font-weight:700;">${setorInfo.nome}</span>
                        <span class="sector-count">${lista.length} consultor${lista.length!==1?'es':''}</span>
                    </div>
                    <div class="metric-percent ${getPercentClass(pctSetor)}">${pctSetor}% OK+Aberto</div>
                </div>
                <div class="consultant-grid">
                    ${lista.sort((a,b)=>b.totalVendasPromocao-a.totalVendasPromocao).map(c => {
                        const p = calcPercent((c.totalOk||0)+(c.totalEmAberto||0), c.totalVendasPromocao);
                        return `
                        <div class="consultant-card" style="border-left:3px solid ${setorInfo.cor};">
                            <div class="consultant-header">
                                <div class="consultant-name">${c.nome}</div>
                                <div class="consultant-sector" style="background:${setorInfo.cor}20;color:${setorInfo.cor};">${c.setor}</div>
                            </div>
                            <div class="metric-grid">
                                ${metricItem('Total Vendas', c.totalVendasPromocao)}
                                ${metricItem('OK', c.totalOk||0, 'var(--success)')}
                                ${metricItem('Em Aberto', c.totalEmAberto||0, 'var(--warning)')}
                                ${metricItem('Atraso', c.totalAtraso||0, 'var(--danger)')}
                                ${metricItem('Outros', c.totalOutros||0, 'var(--gray)')}
                                ${metricPercent('% OK+Aberto', p, setorInfo.cor)}
                            </div>
                        </div>`;
                    }).join('')}
                </div>
            </div>`;
        }).join('')}

        <!-- Resumo Final -->
        <div class="card" style="margin-top:40px;background:linear-gradient(135deg,#0d9488,#0f766e);color:white;">
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
function renderRefuturizaDashboard() {
    const { geral, consultores, mes, ano } = dashboardData;

    if (!geral || !consultores) {
        showError('Dados do Refuturiza não encontrados');
        return;
    }

    const percentGeral = calcPercent(geral.comLigacao, geral.total);

    const html = `
        <h2 class="dash-title">
            <i class="fas fa-book" style="color:#0ea5e9;"></i>
            Dashboard Refuturiza — ${mes} ${ano}
        </h2>

        <div class="card card-refut" style="max-width:500px;margin:0 auto 40px;">
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
        <h3 class="section-title"><i class="fas fa-users" style="color:#0ea5e9;"></i> Desempenho por Consultor (${consultores.length})</h3>
        <div class="consultant-grid">
            ${consultores.map(c => {
                const p = calcPercent(c.comLigacao||0, c.total);
                return `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${c.nome}</div>
                        <div class="consultant-sector" style="background:rgba(14,165,233,0.1);color:#0ea5e9;">REFUTURIZA</div>
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
        <div class="card" style="margin-top:40px;background:linear-gradient(90deg,#0ea5e9,#3b82f6);color:white;">
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
        <div class="error-message" style="text-align:center;padding:40px;">
            <i class="fas fa-info-circle" style="font-size:3rem;margin-bottom:20px;color:var(--gray);"></i>
            <h3>Nenhum consultor com cursos neste período</h3>
            <p>Não foram encontrados dados para ${mes} ${ano}.</p>
        </div>`}
    `;
    dashboardContent.innerHTML = html;
}

// ============================================================================
// HELPERS DE RENDERIZAÇÃO
// ============================================================================
function metricItem(label, value, color) {
    const style = color ? `style="color:${color};"` : '';
    return `
    <div class="metric-item">
        <div class="metric-label">${label}</div>
        <div class="metric-value" ${style}>${value ?? 0}</div>
    </div>`;
}

function metricPercent(label, value, bgColor) {
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
        <div class="metric-label" style="color:rgba(255,255,255,0.8);">${label}</div>
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
        const ia = order.indexOf(a);
        const ib = order.indexOf(b);
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

// ============================================================================
// LOADING / ERROR
// ============================================================================
function showLoading() {
    // CORRIGIDO: não usa innerHTML='' nem appendChild (evita perda do elemento do DOM)
    if (loadingEl) loadingEl.style.display = 'flex';
    // Limpar apenas o conteúdo que não seja o loader
    const children = Array.from(dashboardContent.children);
    children.forEach(child => {
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
            <i class="fas fa-exclamation-triangle" style="font-size:2rem;margin-bottom:15px;"></i>
            <h3 style="margin-bottom:10px;">Erro ao carregar dados</h3>
            <p>${message}</p>
            <button class="btn btn-primary" onclick="loadDashboard()" style="margin-top:15px;">
                <i class="fas fa-redo"></i> Tentar Novamente
            </button>
        </div>
    `;
}

// ============================================================================
// EXPORT
// ============================================================================
function exportPage() {
    const originalTitle = document.title;
    document.title = `Dashboard ${currentDashboard.toUpperCase()} - ${getMonthName(currentMonth)} ${currentYear}`;
    const toHide = document.querySelectorAll('.dashboard-selector, .period-selector, .btn');
    toHide.forEach(el => el.style.display = 'none');
    window.print();
    toHide.forEach(el => el.style.display = '');
    document.title = originalTitle;
}

function getMonthName(month) {
    return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][month - 1];
}
