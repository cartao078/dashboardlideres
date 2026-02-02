// Configuration
const API_URL = 'https://script.google.com/macros/s/AKfycbxsjfbI6Z33q4wr69Tb-Fw66WtfpfaOhspYZ12dqqyOGrPd1PV5ahe9FCzvS0k1wpx7Rg/exec';

// State
let currentDashboard = 'documentacao';
let currentMonth = new Date().getMonth() + 1;
let currentYear = new Date().getFullYear();
let dashboardData = {};

// DOM Elements
const dashboardBtns = document.querySelectorAll('.dashboard-btn');
const monthSelect = document.getElementById('monthSelect');
const yearSelect = document.getElementById('yearSelect');
const dashboardContent = document.getElementById('dashboardContent');
const refreshBtn = document.getElementById('refreshBtn');
const downloadBtn = document.getElementById('downloadBtn');
const loadingEl = document.getElementById('loading');
const lastUpdateEl = document.getElementById('lastUpdate');

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    initializeYearSelect();
    setCurrentPeriod();
    loadDashboard();
    setupEventListeners();
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
    yearSelect.value = currentYear;
}

function setupEventListeners() {
    // Dashboard buttons
    dashboardBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            dashboardBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            currentDashboard = this.dataset.dashboard;
            loadDashboard();
        });
    });

    // Period selectors
    monthSelect.addEventListener('change', function() {
        currentMonth = parseInt(this.value);
        loadDashboard();
    });

    yearSelect.addEventListener('change', function() {
        currentYear = parseInt(this.value);
        loadDashboard();
    });

    // Refresh button
    refreshBtn.addEventListener('click', loadDashboard);

    // Download button
    downloadBtn.addEventListener('click', exportPage);
}

function updateLastUpdateTime() {
    const now = new Date();
    lastUpdateEl.textContent = now.toLocaleString('pt-BR');
}

async function loadDashboard() {
    showLoading();
    updateLastUpdateTime();
    
    try {
        let url = `${API_URL}?endpoint=${currentDashboard}`;
        
        // Add month/year params for dashboards that need them
        if (currentDashboard !== 'recorrencia') {
            url += `&mes=${currentMonth}&ano=${currentYear}`;
        }
        
        console.log('Fetching from:', url);
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
    
    switch(currentDashboard) {
        case 'documentacao':
            renderDocumentacaoDashboard();
            break;
        case 'app':
            renderAppDashboard();
            break;
        case 'adimplencia':
            renderAdimplenciaDashboard();
            break;
        case 'recorrencia':
            renderRecorrenciaDashboard();
            break;
        case 'refuturiza':
            renderRefuturizaDashboard();
            break;
    }
}

function renderDocumentacaoDashboard() {
    const { geral, vendasLoja, vendasWeb, consultores, mes, ano } = dashboardData;
    
    // Calculate percentages
    const percentGeral = geral.total > 0 ? Math.round((geral.aprovados / geral.total) * 100) : 0;
    const percentLoja = vendasLoja.total > 0 ? Math.round((vendasLoja.aprovados / vendasLoja.total) * 100) : 0;
    const percentWeb = vendasWeb.total > 0 ? Math.round((vendasWeb.aprovados / vendasWeb.total) * 100) : 0;
    
    // Group consultants by sector
    const consultantsBySector = {};
    consultores.forEach(c => {
        const sector = c.setor || 'OUTROS';
        if (!consultantsBySector[sector]) {
            consultantsBySector[sector] = [];
        }
        consultantsBySector[sector].push(c);
    });
    
    // Order sectors
    const sectorOrder = ['VENDAS', 'RECEPCAO', 'REFILIACAO', 'WEB SITE', 'TELEVENDAS', 'OUTROS'];
    const sortedSectors = Object.keys(consultantsBySector).sort((a, b) => {
        return sectorOrder.indexOf(a) - sectorOrder.indexOf(b);
    });
    
    const html = `
        <h2 style="margin-bottom: 25px; color: var(--dark); font-size: 1.5rem;">
            <i class="fas fa-folder" style="color: var(--primary); margin-right: 10px;"></i>
            Dashboard de Vendas - ${mes} ${ano}
        </h2>
        
        <!-- Main Cards -->
        <div class="main-cards">
            <div class="card card-doc">
                <div class="card-header">
                    <div class="card-title">Total de Vendas</div>
                    <div class="card-icon">
                        <i class="fas fa-chart-bar"></i>
                    </div>
                </div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="metric-label">Total Vendas</div>
                        <div class="metric-value">${geral.total}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Aprovados</div>
                        <div class="metric-value">${geral.aprovados}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Pendências</div>
                        <div class="metric-value">${geral.pendencias}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">% Aprovados</div>
                        <div class="metric-percent ${getPercentClass(percentGeral)}">${percentGeral}%</div>
                    </div>
                </div>
            </div>
            
            <div class="card card-doc">
                <div class="card-header">
                    <div class="card-title">Vendas Loja</div>
                    <div class="card-icon">
                        <i class="fas fa-store"></i>
                    </div>
                </div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="metric-label">Total Vendas</div>
                        <div class="metric-value">${vendasLoja.total}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Aprovados</div>
                        <div class="metric-value">${vendasLoja.aprovados}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Pendências</div>
                        <div class="metric-value">${vendasLoja.pendencias}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">% Aprovados</div>
                        <div class="metric-percent ${getPercentClass(percentLoja)}">${percentLoja}%</div>
                    </div>
                </div>
            </div>
            
            <div class="card card-doc">
                <div class="card-header">
                    <div class="card-title">Vendas Web/Tele</div>
                    <div class="card-icon">
                        <i class="fas fa-globe"></i>
                    </div>
                </div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="metric-label">Total Vendas</div>
                        <div class="metric-value">${vendasWeb.total}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Aprovados</div>
                        <div class="metric-value">${vendasWeb.aprovados}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Pendências</div>
                        <div class="metric-value">${vendasWeb.pendencias}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">% Aprovados</div>
                        <div class="metric-percent ${getPercentClass(percentWeb)}">${percentWeb}%</div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Detailed by Sector -->
        <h3 style="margin: 40px 0 25px; color: var(--dark); font-size: 1.3rem; padding-bottom: 10px; border-bottom: 2px solid var(--gray-light);">
            <i class="fas fa-layer-group" style="color: var(--primary); margin-right: 10px;"></i>
            Desempenho por Setor
        </h3>
        
        ${sortedSectors.map(sector => {
            const sectorConsultants = consultantsBySector[sector];
            const sectorTotal = sectorConsultants.reduce((sum, c) => sum + c.total, 0);
            const sectorAprovados = sectorConsultants.reduce((sum, c) => sum + c.aprovados, 0);
            const sectorPercent = sectorTotal > 0 ? Math.round((sectorAprovados / sectorTotal) * 100) : 0;
            const sectorClass = getSectorClass(sector);
            
            return `
            <div class="sector-card">
                <div class="sector-header">
                    <div class="sector-title">
                        <i class="${getSectorIcon(sector)}"></i>
                        ${sector}
                        <span class="sector-count">${sectorConsultants.length} consultores</span>
                    </div>
                    <div class="metric-percent ${getPercentClass(sectorPercent)}">${sectorPercent}% aprovados</div>
                </div>
                
                <div class="consultant-grid">
                    ${sectorConsultants.sort((a, b) => b.total - a.total).map(consultor => {
                        const percent = consultor.total > 0 ? Math.round((consultor.aprovados / consultor.total) * 100) : 0;
                        
                        return `
                        <div class="consultant-card">
                            <div class="consultant-header">
                                <div class="consultant-name">${consultor.nome}</div>
                                <div class="consultant-sector ${sectorClass}">${consultor.setor}</div>
                            </div>
                            <div class="metric-grid">
                                <div class="metric-item">
                                    <div class="metric-label">Total</div>
                                    <div class="metric-value">${consultor.total}</div>
                                </div>
                                <div class="metric-item">
                                    <div class="metric-label">Aprovados</div>
                                    <div class="metric-value" style="color: var(--success);">${consultor.aprovados}</div>
                                </div>
                                <div class="metric-item">
                                    <div class="metric-label">Pendências</div>
                                    <div class="metric-value" style="color: var(--danger);">${consultor.pendencias}</div>
                                </div>
                                <div class="metric-item">
                                    <div class="metric-label">% Aprovados</div>
                                    <div class="metric-percent ${getPercentClass(percent)}">${percent}%</div>
                                </div>
                            </div>
                        </div>
                        `;
                    }).join('')}
                </div>
            </div>
            `;
        }).join('')}
    `;
    
    dashboardContent.innerHTML = html;
}

function renderAppDashboard() {
    const { geral, appLoja, appWeb, consultores, consultorasRetencao, mes, ano } = dashboardData;
    
    // Calculate percentages
    const percentGeral = geral.total > 0 ? Math.round((geral.sim / geral.total) * 100) : 0;
    const percentLoja = appLoja.total > 0 ? Math.round((appLoja.sim / appLoja.total) * 100) : 0;
    const percentWeb = appWeb.total > 0 ? Math.round((appWeb.sim / appWeb.total) * 100) : 0;
    
    const html = `
        <h2 style="margin-bottom: 25px; color: var(--dark); font-size: 1.5rem;">
            <i class="fas fa-mobile-alt" style="color: var(--secondary); margin-right: 10px;"></i>
            Dashboard App - ${mes} ${ano}
        </h2>
        
        <!-- Main Cards -->
        <div class="main-cards">
            <div class="card card-app">
                <div class="card-header">
                    <div class="card-title">App - Total Geral</div>
                    <div class="card-icon">
                        <i class="fas fa-chart-pie"></i>
                    </div>
                </div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="metric-label">Total Clientes</div>
                        <div class="metric-value">${geral.total}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Com App (SIM)</div>
                        <div class="metric-value" style="color: var(--success);">${geral.sim}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Sem App (NÃO)</div>
                        <div class="metric-value" style="color: var(--danger);">${geral.nao}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">% Com App</div>
                        <div class="metric-percent ${getPercentClass(percentGeral)}">${percentGeral}%</div>
                    </div>
                </div>
            </div>
            
            <div class="card card-app">
                <div class="card-header">
                    <div class="card-title">App - Loja</div>
                    <div class="card-icon">
                        <i class="fas fa-store"></i>
                    </div>
                </div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="metric-label">Total Clientes</div>
                        <div class="metric-value">${appLoja.total}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Com App (SIM)</div>
                        <div class="metric-value" style="color: var(--success);">${appLoja.sim}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Sem App (NÃO)</div>
                        <div class="metric-value" style="color: var(--danger);">${appLoja.nao}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">% Com App</div>
                        <div class="metric-percent ${getPercentClass(percentLoja)}">${percentLoja}%</div>
                    </div>
                </div>
            </div>
            
            <div class="card card-app">
                <div class="card-header">
                    <div class="card-title">App - Web/Tele</div>
                    <div class="card-icon">
                        <i class="fas fa-globe"></i>
                    </div>
                </div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="metric-label">Total Clientes</div>
                        <div class="metric-value">${appWeb.total}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Com App (SIM)</div>
                        <div class="metric-value" style="color: var(--success);">${appWeb.sim}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">Sem App (NÃO)</div>
                        <div class="metric-value" style="color: var(--danger);">${appWeb.nao}</div>
                    </div>
                    <div class="metric-item">
                        <div class="metric-label">% Com App</div>
                        <div class="metric-percent ${getPercentClass(percentWeb)}">${percentWeb}%</div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Retention Consultants -->
        ${consultorasRetencao && consultorasRetencao.length > 0 ? `
        <div class="retention-section">
            <div class="retention-header">
                <i class="fas fa-crown"></i>
                <h3>Consultoras de Retenção (Dados Separados)</h3>
            </div>
            
            <div class="consultant-grid">
                ${consultorasRetencao.map(consultora => {
                    const percent = consultora.total > 0 ? Math.round((consultora.sim / consultora.total) * 100) : 0;
                    
                    return `
                    <div class="consultant-card" style="border-left: 4px solid #f59e0b;">
                        <div class="consultant-header">
                            <div class="consultant-name">${consultora.nome} (RETENÇÃO)</div>
                            <div class="consultant-sector sector-retencao">RETENÇÃO</div>
                        </div>
                        <div class="metric-grid">
                            <div class="metric-item">
                                <div class="metric-label">Total Clientes</div>
                                <div class="metric-value">${consultora.total}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">Com App</div>
                                <div class="metric-value" style="color: var(--success);">${consultora.sim}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">Sem App</div>
                                <div class="metric-value" style="color: var(--danger);">${consultora.nao}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">Cancelados</div>
                                <div class="metric-value" style="color: var(--gray);">${consultora.cancelado || 0}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">% Com App</div>
                                <div class="metric-percent ${getPercentClass(percent)}">${percent}%</div>
                            </div>
                        </div>
                    </div>
                    `;
                }).join('')}
            </div>
        </div>
        ` : ''}
        
        <!-- Regular Consultants by Sector -->
        <h3 style="margin: 40px 0 25px; color: var(--dark); font-size: 1.3rem; padding-bottom: 10px; border-bottom: 2px solid var(--gray-light);">
            <i class="fas fa-layer-group" style="color: var(--secondary); margin-right: 10px;"></i>
            Desempenho por Setor
        </h3>
        
        ${renderAppBySector(consultores)}
    `;
    
    dashboardContent.innerHTML = html;
}

function renderAppBySector(consultores) {
    // Filter out retention consultants
    const regularConsultants = consultores.filter(c => !c.origem || c.origem !== 'retencao');
    
    // Group consultants by sector
    const consultantsBySector = {};
    regularConsultants.forEach(c => {
        const sector = c.setor || 'OUTROS';
        if (!consultantsBySector[sector]) {
            consultantsBySector[sector] = [];
        }
        consultantsBySector[sector].push(c);
    });
    
    // Order sectors
    const sectorOrder = ['VENDAS', 'RECEPCAO', 'REFILIACAO', 'WEB SITE', 'TELEVENDAS', 'OUTROS'];
    const sortedSectors = Object.keys(consultantsBySector).sort((a, b) => {
        return sectorOrder.indexOf(a) - sectorOrder.indexOf(b);
    });
    
    if (sortedSectors.length === 0) {
        return '<p style="text-align: center; color: var(--gray); padding: 40px;">Nenhum consultor regular encontrado</p>';
    }
    
    return sortedSectors.map(sector => {
        const sectorConsultants = consultantsBySector[sector];
        const sectorTotal = sectorConsultants.reduce((sum, c) => sum + c.total, 0);
        const sectorSim = sectorConsultants.reduce((sum, c) => sum + (c.sim || 0), 0);
        const sectorPercent = sectorTotal > 0 ? Math.round((sectorSim / sectorTotal) * 100) : 0;
        const sectorClass = getSectorClass(sector);
        
        return `
        <div class="sector-card">
            <div class="sector-header">
                <div class="sector-title">
                    <i class="${getSectorIcon(sector)}"></i>
                    ${sector}
                    <span class="sector-count">${sectorConsultants.length} consultores</span>
                </div>
                <div class="metric-percent ${getPercentClass(sectorPercent)}">${sectorPercent}% com app</div>
            </div>
            
            <div class="consultant-grid">
                ${sectorConsultants.sort((a, b) => b.total - a.total).map(consultor => {
                    const percent = consultor.total > 0 ? Math.round(((consultor.sim || 0) / consultor.total) * 100) : 0;
                    
                    return `
                    <div class="consultant-card">
                        <div class="consultant-header">
                            <div class="consultant-name">${consultor.nome}</div>
                            <div class="consultant-sector ${sectorClass}">${consultor.setor}</div>
                        </div>
                        <div class="metric-grid">
                            <div class="metric-item">
                                <div class="metric-label">Total</div>
                                <div class="metric-value">${consultor.total}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">Com App</div>
                                <div class="metric-value" style="color: var(--success);">${consultor.sim || 0}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">Sem App</div>
                                <div class="metric-value" style="color: var(--danger);">${consultor.nao || 0}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">Cancelados</div>
                                <div class="metric-value" style="color: var(--gray);">${consultor.cancelado || 0}</div>
                            </div>
                            <div class="metric-item">
                                <div class="metric-label">% Com App</div>
                                <div class="metric-percent ${getPercentClass(percent)}">${percent}%</div>
                            </div>
                        </div>
                    </div>
                    `;
                }).join('')}
            </div>
        </div>
        `;
    }).join('');
}

function renderAdimplenciaDashboard() {
    const { geral, consultores, mes, ano } = dashboardData;
    
    const html = `
        <h2 style="margin-bottom: 25px; color: var(--dark); font-size: 1.5rem;">
            <i class="fas fa-credit-card" style="color: var(--success); margin-right: 10px;"></i>
            Dashboard Adimplência - ${mes} ${ano}
        </h2>
        
        <!-- Main Card -->
        <div class="card card-adim" style="max-width: 500px; margin: 0 auto 40px;">
            <div class="card-header">
                <div class="card-title">Adimplência - Total da Loja</div>
                <div class="card-icon">
                    <i class="fas fa-chart-line"></i>
                </div>
            </div>
            <div class="metric-grid">
                <div class="metric-item">
                    <div class="metric-label">Total Trocas</div>
                    <div class="metric-value">${geral.totalTrocas}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Mens. OK</div>
                    <div class="metric-value" style="color: var(--success);">${geral.mensOk}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Mens. Aberto</div>
                    <div class="metric-value" style="color: var(--warning);">${geral.mensAberto}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Mens. Atraso</div>
                    <div class="metric-value" style="color: var(--danger);">${geral.mensAtraso}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Aprovados</div>
                    <div class="metric-value" style="color: var(--success);">${geral.aprovados}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Pendentes</div>
                    <div class="metric-value" style="color: var(--danger);">${geral.pendentes}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">% Aprovados</div>
                    <div class="metric-percent ${getPercentClass(geral.percentualAprovado)}">${geral.percentualAprovado}%</div>
                </div>
            </div>
        </div>
        
        <!-- Consultant Table -->
        <h3 style="margin: 30px 0 20px; color: var(--dark); font-size: 1.2rem;">
            <i class="fas fa-user-tie" style="color: var(--primary); margin-right: 10px;"></i>
            Desempenho por Consultor
        </h3>
        
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
                    <th>% Aprovados</th>
                </tr>
            </thead>
            <tbody>
                ${consultores.map(consultor => {
                    return `
                    <tr>
                        <td><strong>${consultor.nome}</strong></td>
                        <td>${consultor.totalTrocas}</td>
                        <td style="color: var(--success);">${consultor.mensOk}</td>
                        <td style="color: var(--warning);">${consultor.mensAberto}</td>
                        <td style="color: var(--danger);">${consultor.mensAtraso}</td>
                        <td style="color: var(--success);">${consultor.aprovados}</td>
                        <td style="color: var(--danger);">${consultor.pendentes}</td>
                        <td><span class="metric-percent ${getPercentClass(consultor.percentualAprovado)}">${consultor.percentualAprovado}%</span></td>
                    </tr>
                    `;
                }).join('')}
            </tbody>
        </table>
    `;
    
    dashboardContent.innerHTML = html;
}

function renderRecorrenciaDashboard() {
    const { retencao, refiliacao, periodo } = dashboardData;
    
    const html = `
        <h2 style="margin-bottom: 25px; color: var(--dark); font-size: 1.5rem;">
            <i class="fas fa-redo" style="color: var(--warning); margin-right: 10px;"></i>
            Dashboard Recorrência - ${periodo.atual}
        </h2>
        
        <div style="background: #fef3c7; padding: 15px; border-radius: 10px; margin-bottom: 30px;">
            <p style="margin: 0; color: #92400e; font-weight: 500;">
                <i class="fas fa-info-circle"></i> Período atual: ${periodo.atual} | Histórico: ${periodo.historico.join(', ')}
            </p>
        </div>
        
        <!-- Retention Section -->
        <h3 style="margin: 30px 0 20px; color: var(--dark); font-size: 1.2rem; border-bottom: 2px solid var(--primary); padding-bottom: 10px;">
            <i class="fas fa-crown" style="color: var(--primary); margin-right: 10px;"></i>
            Retenção
        </h3>
        
        <div class="consultant-grid" style="margin-bottom: 40px;">
            ${Object.keys(retencao).map(key => {
                const c = retencao[key];
                const percentAtual = c.atual.totalRetidosFinal > 0 ? Math.round((c.atual.retençõesOK / c.atual.totalRetidosFinal) * 100) : 0;
                const percentTotal = c.total3Meses.totalRetidosFinal > 0 ? Math.round((c.total3Meses.totalOK / c.total3Meses.totalRetidosFinal) * 100) : 0;
                
                return `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${key} (RETENÇÃO)</div>
                        <div class="consultant-sector sector-vendas">RETENÇÃO</div>
                    </div>
                    
                    <h4 style="margin: 15px 0 10px; color: var(--gray); font-size: 0.9rem;">
                        <i class="far fa-calendar-check"></i> Mês Atual (${periodo.atual})
                    </h4>
                    <div class="metric-grid">
                        <div class="metric-item">
                            <div class="metric-label">Total Retido</div>
                            <div class="metric-value">${c.atual.totalRetido}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">Retenções OK</div>
                            <div class="metric-value" style="color: var(--success);">${c.atual.retençõesOK}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">% OK</div>
                            <div class="metric-percent ${getPercentClass(percentAtual)}">${percentAtual}%</div>
                        </div>
                    </div>
                    
                    <h4 style="margin: 15px 0 10px; color: var(--gray); font-size: 0.9rem;">
                        <i class="fas fa-chart-line"></i> Total 3 Meses
                    </h4>
                    <div class="metric-grid">
                        <div class="metric-item">
                            <div class="metric-label">Total Retido</div>
                            <div class="metric-value">${c.total3Meses.totalRetido}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">Retenções OK</div>
                            <div class="metric-value" style="color: var(--success);">${c.total3Meses.totalOK || 0}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">% OK</div>
                            <div class="metric-percent ${getPercentClass(percentTotal)}">${percentTotal}%</div>
                        </div>
                    </div>
                </div>
                `;
            }).join('')}
        </div>
        
        <!-- Refiliation Section -->
        <h3 style="margin: 30px 0 20px; color: var(--dark); font-size: 1.2rem; border-bottom: 2px solid var(--warning); padding-bottom: 10px;">
            <i class="fas fa-user-plus" style="color: var(--warning); margin-right: 10px;"></i>
            Refiliação
        </h3>
        
        <div class="consultant-grid">
            ${Object.keys(refiliacao).map(key => {
                const c = refiliacao[key];
                const percentTotal = c.total3Meses.totalRetidosFinal > 0 ? Math.round((c.total3Meses.totalOK / c.total3Meses.totalRetidosFinal) * 100) : 0;
                
                return `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${key} (REFILIAÇÃO)</div>
                        <div class="consultant-sector sector-refiliacao">REFILIAÇÃO</div>
                    </div>
                    
                    <h4 style="margin: 15px 0 10px; color: var(--gray); font-size: 0.9rem;">
                        <i class="fas fa-chart-line"></i> Total 3 Meses
                    </h4>
                    <div class="metric-grid">
                        <div class="metric-item">
                            <div class="metric-label">Total Refiliados</div>
                            <div class="metric-value">${c.total3Meses.totalRetido}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">Refiliações OK</div>
                            <div class="metric-value" style="color: var(--success);">${c.total3Meses.totalOK || 0}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">% OK</div>
                            <div class="metric-percent ${getPercentClass(percentTotal)}">${percentTotal}%</div>
                        </div>
                    </div>
                </div>
                `;
            }).join('')}
        </div>
    `;
    
    dashboardContent.innerHTML = html;
}

function renderRefuturizaDashboard() {
    const { geral, consultores, mes, ano } = dashboardData;
    
    if (!geral || !consultores) {
        showError('Dados do Refuturiza não encontrados');
        return;
    }
    
    const percentGeral = geral.total > 0 ? Math.round((geral.comLigacao / geral.total) * 100) : 0;
    
    const html = `
        <h2 style="margin-bottom: 25px; color: var(--dark); font-size: 1.5rem;">
            <i class="fas fa-book" style="color: #0ea5e9; margin-right: 10px;"></i>
            Dashboard Refuturiza - ${mes} ${ano}
        </h2>
        
        <!-- Main Card -->
        <div class="card card-refut" style="max-width: 500px; margin: 0 auto 40px;">
            <div class="card-header">
                <div class="card-title">Refuturiza - Cursos Online</div>
                <div class="card-icon">
                    <i class="fas fa-book-open"></i>
                </div>
            </div>
            <div class="metric-grid">
                <div class="metric-item">
                    <div class="metric-label">Total Cursos</div>
                    <div class="metric-value">${geral.total || 0}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Com Ligação</div>
                    <div class="metric-value" style="color: var(--success);">${geral.comLigacao || 0}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Sem Ligação</div>
                    <div class="metric-value" style="color: var(--danger);">${geral.semLigacao || 0}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">Cancelados</div>
                    <div class="metric-value" style="color: var(--gray);">${geral.cancelado || 0}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label">% Com Ligação</div>
                    <div class="metric-percent ${getPercentClass(percentGeral)}">${percentGeral}%</div>
                </div>
            </div>
        </div>
        
        ${consultores && consultores.length > 0 ? `
        <!-- All Consultants -->
        <h3 style="margin: 40px 0 20px; color: var(--dark); font-size: 1.2rem;">
            <i class="fas fa-users" style="color: #0ea5e9; margin-right: 10px;"></i>
            Desempenho por Consultor (${consultores.length})
        </h3>
        
        <div class="consultant-grid">
            ${consultores.map(consultor => {
                const percent = consultor.total > 0 ? Math.round(((consultor.comLigacao || 0) / consultor.total) * 100) : 0;
                
                return `
                <div class="consultant-card">
                    <div class="consultant-header">
                        <div class="consultant-name">${consultor.nome}</div>
                        <div class="consultant-sector" style="background: rgba(14, 165, 233, 0.1); color: #0ea5e9;">CURSOS ONLINE</div>
                    </div>
                    <div class="metric-grid">
                        <div class="metric-item">
                            <div class="metric-label">Total Cursos</div>
                            <div class="metric-value">${consultor.total || 0}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">Com Ligação</div>
                            <div class="metric-value" style="color: var(--success);">${consultor.comLigacao || 0}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">Sem Ligação</div>
                            <div class="metric-value" style="color: var(--danger);">${consultor.semLigacao || 0}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">Cancelados</div>
                            <div class="metric-value" style="color: var(--gray);">${consultor.cancelado || 0}</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-label">% Com Ligação</div>
                            <div class="metric-percent ${getPercentClass(percent)}">${percent}%</div>
                        </div>
                    </div>
                </div>
                `;
            }).join('')}
        </div>
        
        <!-- Summary -->
        <div class="card" style="margin-top: 40px; background: linear-gradient(90deg, #0ea5e9 0%, #3b82f6 100%); color: white;">
            <div class="card-header" style="border-bottom-color: rgba(255,255,255,0.2);">
                <div class="card-title" style="color: white;">📊 Resumo Final - Cursos Online</div>
                <div class="card-icon" style="background: rgba(255,255,255,0.2);">
                    <i class="fas fa-graduation-cap"></i>
                </div>
            </div>
            <div class="metric-grid">
                <div class="metric-item">
                    <div class="metric-label" style="color: rgba(255,255,255,0.8);">Total Consultores</div>
                    <div class="metric-value" style="color: white;">${consultores.length}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label" style="color: rgba(255,255,255,0.8);">Total Cursos</div>
                    <div class="metric-value" style="color: white;">${geral.total || 0}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label" style="color: rgba(255,255,255,0.8);">Média por Consultor</div>
                    <div class="metric-value" style="color: white;">${consultores.length > 0 ? Math.round((geral.total || 0) / consultores.length) : 0}</div>
                </div>
                <div class="metric-item">
                    <div class="metric-label" style="color: rgba(255,255,255,0.8);">Taxa de Contato</div>
                    <div class="metric-percent" style="background: rgba(255,255,255,0.2); color: white;">${percentGeral}%</div>
                </div>
            </div>
        </div>
        ` : `
        <div class="error-message" style="text-align: center; padding: 40px;">
            <i class="fas fa-info-circle" style="font-size: 3rem; margin-bottom: 20px; color: var(--gray);"></i>
            <h3 style="margin-bottom: 10px; color: var(--gray);">Nenhum consultor com cursos neste período</h3>
            <p style="color: var(--gray);">Não foram encontrados dados de cursos online para ${mes} ${ano}.</p>
        </div>
        `}
    `;
    
    dashboardContent.innerHTML = html;
}

// Helper Functions
function getPercentClass(percent) {
    if (percent >= 90) return 'percent-high';
    if (percent >= 80) return 'percent-medium';
    return 'percent-low';
}

function getSectorClass(sector) {
    if (!sector) return 'sector-outros';
    
    switch(sector.toUpperCase()) {
        case 'VENDAS': return 'sector-vendas';
        case 'RECEPCAO': return 'sector-recepcao';
        case 'REFILIACAO': return 'sector-refiliacao';
        case 'WEB SITE': case 'WEB': return 'sector-web';
        case 'TELEVENDAS': return 'sector-televendas';
        case 'RETENÇÃO': case 'RETENCAO': return 'sector-retencao';
        default: return 'sector-outros';
    }
}

function getSectorIcon(sector) {
    if (!sector) return 'fas fa-users';
    
    switch(sector.toUpperCase()) {
        case 'VENDAS': return 'fas fa-shopping-cart';
        case 'RECEPCAO': return 'fas fa-headset';
        case 'REFILIACAO': return 'fas fa-user-plus';
        case 'WEB SITE': case 'WEB': return 'fas fa-globe';
        case 'TELEVENDAS': return 'fas fa-phone-alt';
        case 'RETENÇÃO': case 'RETENCAO': return 'fas fa-crown';
        default: return 'fas fa-users';
    }
}

function showLoading() {
    loadingEl.style.display = 'block';
    dashboardContent.innerHTML = '';
    dashboardContent.appendChild(loadingEl);
}

function hideLoading() {
    loadingEl.style.display = 'none';
}

function showError(message) {
    hideLoading();
    dashboardContent.innerHTML = `
        <div class="error-message">
            <i class="fas fa-exclamation-triangle" style="font-size: 2rem; margin-bottom: 15px;"></i>
            <h3 style="margin-bottom: 10px;">Erro ao carregar dados</h3>
            <p>${message}</p>
            <button class="btn btn-primary" onclick="loadDashboard()" style="margin-top: 15px;">
                <i class="fas fa-redo"></i> Tentar Novamente
            </button>
        </div>
    `;
}

function exportPage() {
    // Create a print-friendly version
    const originalTitle = document.title;
    document.title = `Dashboard ${currentDashboard.toUpperCase()} - ${getMonthName(currentMonth)} ${currentYear}`;
    
    // Hide some elements for printing
    const elementsToHide = document.querySelectorAll('.dashboard-selector, .period-selector, .btn');
    elementsToHide.forEach(el => el.style.display = 'none');
    
    // Print
    window.print();
    
    // Restore elements
    elementsToHide.forEach(el => el.style.display = '');
    document.title = originalTitle;
}

function getMonthName(month) {
    const months = [
        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    return months[month - 1];
}
