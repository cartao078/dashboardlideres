// script.js

// CONFIGURAÇÃO DA API - SUA URL DO APPS SCRIPT
const API_BASE_URL = "https://script.google.com/macros/s/AKfycbwMISyxH0g_Qku9R0JjTUaR998yZIPDYdOViiUm3l7MJbl_WqowErW4GhwLLEYwgjLI9g/exec"; // Substitua pela URL do deployAPI()

// Variáveis globais
let dadosAtuais = null;
let chartMain = null;
let chartPie = null;
let apiURL = API_BASE_URL;

// Inicializar dashboard
function inicializarDashboard() {
    console.log("🚀 Inicializando Dashboard...");
    
    // Atualizar data atual
    atualizarDataAtual();
    
    // Preencher anos (últimos 3 anos)
    preencherAnos();
    
    // Definir mês atual
    const mesAtual = new Date().getMonth() + 1;
    document.getElementById('mesSelect').value = mesAtual;
    
    // Testar conexão com API
    testarConexaoAPI();
    
    // Carregar dados iniciais
    setTimeout(() => carregarDados(), 1000);
}

// Atualizar data atual no header
function atualizarDataAtual() {
    const agora = new Date();
    const options = { 
        weekday: 'long', 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    };
    document.getElementById('data-atual').textContent = 
        agora.toLocaleDateString('pt-BR', options);
}

// Preencher seletor de anos
function preencherAnos() {
    const anoSelect = document.getElementById('anoSelect');
    const anoAtual = new Date().getFullYear();
    
    // Limpar opções existentes
    anoSelect.innerHTML = '';
    
    // Adicionar últimos 3 anos e próximo ano
    for (let i = -1; i <= 1; i++) {
        const ano = anoAtual + i;
        const option = document.createElement('option');
        option.value = ano;
        option.textContent = ano;
        if (i === 0) option.selected = true;
        anoSelect.appendChild(option);
    }
}

// Testar conexão com API
async function testarConexaoAPI() {
    try {
        const response = await fetch(apiURL + '?endpoint=documentacao&mes=1&ano=2024');
        if (response.ok) {
            document.getElementById('status-api').className = 'badge bg-success ms-2';
            document.getElementById('status-api').textContent = '● ONLINE';
            document.getElementById('api-url').textContent = apiURL;
        } else {
            throw new Error('API não respondeu');
        }
    } catch (error) {
        console.error('Erro na conexão com API:', error);
        document.getElementById('status-api').className = 'badge bg-danger ms-2';
        document.getElementById('status-api').textContent = '● OFFLINE';
        
        // Tentar URL alternativa (para desenvolvimento)
        if (apiURL === API_BASE_URL) {
            apiURL = prompt('API não encontrada. Por favor, cole a URL da sua API (obtida com deployAPI()):', API_BASE_URL);
            if (apiURL) testarConexaoAPI();
        }
    }
}

// Mostrar modal de loading
function mostrarLoading(mensagem = 'Carregando dados...') {
    document.getElementById('loading-text').textContent = mensagem;
    const modal = new bootstrap.Modal(document.getElementById('loadingModal'));
    modal.show();
}

// Esconder modal de loading
function esconderLoading() {
    const modal = bootstrap.Modal.getInstance(document.getElementById('loadingModal'));
    if (modal) modal.hide();
}

// Mostrar erro
function mostrarErro(mensagem) {
    document.getElementById('error-message').textContent = mensagem;
    const modal = new bootstrap.Modal(document.getElementById('errorModal'));
    modal.show();
}

// Função principal para carregar dados
async function carregarDados() {
    try {
        mostrarLoading('Conectando com a API...');
        
        const mes = document.getElementById('mesSelect').value;
        const ano = document.getElementById('anoSelect').value;
        const dashboard = document.getElementById('dashboardSelect').value;
        
        let endpoint = '';
        let params = `?endpoint=${dashboard}&mes=${mes}&ano=${ano}`;
        
        // URL completa
        const url = apiURL + params;
        console.log(`📡 Fazendo request para: ${url}`);
        
        // Fazer requisição
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`Erro HTTP: ${response.status}`);
        }
        
        const data = await response.json();
        
        if (data.status === 'error') {
            throw new Error(data.error || 'Erro desconhecido na API');
        }
        
        // Salvar dados atuais
        dadosAtuais = data.data;
        
        // Atualizar interface
        atualizarInterface(dashboard, dadosAtuais);
        
        // Atualizar timestamp
        document.getElementById('ultima-atualizacao').textContent = 
            new Date().toLocaleTimeString('pt-BR');
        
        esconderLoading();
        
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        esconderLoading();
        mostrarErro(`Erro: ${error.message}`);
    }
}

// Atualizar interface baseada nos dados
function atualizarInterface(tipo, dados) {
    console.log(`🎨 Atualizando interface para: ${tipo}`, dados);
    
    // Atualizar título
    const meses = [
        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    const mesNome = meses[parseInt(document.getElementById('mesSelect').value) - 1];
    const ano = document.getElementById('anoSelect').value;
    
    document.getElementById('grafico-titulo').textContent = 
        `${tipo.charAt(0).toUpperCase() + tipo.slice(1)} - ${mesNome} ${ano}`;
    document.getElementById('tabela-titulo').textContent = 
        `Detalhamento - ${mesNome} ${ano}`;
    
    // Atualizar KPIs
    atualizarKPIs(tipo, dados);
    
    // Atualizar gráficos
    atualizarGraficos(tipo, dados);
    
    // Atualizar tabela
    atualizarTabela(tipo, dados);
}

// Atualizar KPIs
function atualizarKPIs(tipo, dados) {
    const kpiSection = document.getElementById('kpi-section');
    kpiSection.innerHTML = '';
    
    let kpis = [];
    
    switch(tipo) {
        case 'documentacao':
            kpis = [
                { titulo: 'Total Documentos', valor: dados.geral.total, cor: 'documentacao', icone: 'fa-file-alt' },
                { titulo: 'Aprovados', valor: dados.geral.aprovados, cor: 'documentacao', icone: 'fa-check-circle' },
                { titulo: 'Pendências', valor: dados.geral.pendencias, cor: 'documentacao', icone: 'fa-clock' },
                { titulo: '% Aprovação', valor: calcularPercentual(dados.geral.aprovados, dados.geral.total), cor: 'documentacao', icone: 'fa-percentage', sufixo: '%' }
            ];
            break;
            
        case 'app':
            kpis = [
                { titulo: 'Total', valor: dados.geral.total, cor: 'app', icone: 'fa-mobile-alt' },
                { titulo: 'Com App', valor: dados.geral.sim, cor: 'app', icone: 'fa-check' },
                { titulo: 'Sem App', valor: dados.geral.nao, cor: 'app', icone: 'fa-times' },
                { titulo: '% Com App', valor: calcularPercentual(dados.geral.sim, dados.geral.total), cor: 'app', icone: 'fa-percentage', sufixo: '%' }
            ];
            break;
            
        case 'adimplencia':
            kpis = [
                { titulo: 'Total Trocas', valor: dados.geral.totalTrocas, cor: 'adimplencia', icone: 'fa-exchange-alt' },
                { titulo: 'Aprovados', valor: dados.geral.aprovados, cor: 'adimplencia', icone: 'fa-check' },
                { titulo: 'Pendentes', valor: dados.geral.pendentes, cor: 'adimplencia', icone: 'fa-clock' },
                { titulo: '% Aprovação', valor: dados.geral.percentualAprovado || 0, cor: 'adimplencia', icone: 'fa-percentage', sufixo: '%' }
            ];
            break;
            
        case 'recorrencia':
            // Para recorrência, mostrar totais gerais
            const totalRetencao = Object.values(dados.retencao || {}).reduce((sum, c) => sum + (c.total3Meses?.totalRetido || 0), 0);
            const totalRefiliacao = Object.values(dados.refiliacao || {}).reduce((sum, c) => sum + (c.total3Meses?.totalRetido || 0), 0);
            
            kpis = [
                { titulo: 'Retenção (3M)', valor: totalRetencao, cor: 'recorrencia', icone: 'fa-user-friends' },
                { titulo: 'Refiliação (3M)', valor: totalRefiliacao, cor: 'recorrencia', icone: 'fa-users' },
                { titulo: 'Total Geral', valor: totalRetencao + totalRefiliacao, cor: 'recorrencia', icone: 'fa-chart-line' },
                { titulo: 'Período', valor: dados.periodo?.atual || 'Atual', cor: 'recorrencia', icone: 'fa-calendar' }
            ];
            break;
    }
    
    // Criar cards de KPI
    kpis.forEach(kpi => {
        const col = document.createElement('div');
        col.className = 'col-md-3 col-sm-6 mb-3';
        
        const card = document.createElement('div');
        card.className = `card kpi-card ${kpi.cor} fade-in`;
        card.innerHTML = `
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-start">
                    <div>
                        <h6 class="card-subtitle mb-1 opacity-75">
                            <i class="fas ${kpi.icone} me-1"></i>${kpi.titulo}
                        </h6>
                        <h2 class="kpi-value">${formatarNumero(kpi.valor)}${kpi.sufixo || ''}</h2>
                    </div>
                </div>
            </div>
        `;
        
        col.appendChild(card);
        kpiSection.appendChild(col);
    });
}

// Atualizar gráficos
function atualizarGraficos(tipo, dados) {
    // Destruir gráficos existentes
    if (chartMain) chartMain.destroy();
    if (chartPie) chartPie.destroy();
    
    const ctxMain = document.getElementById('mainChart').getContext('2d');
    const ctxPie = document.getElementById('pieChart').getContext('2d');
    
    let mainChartData = {};
    let pieChartData = {};
    
    switch(tipo) {
        case 'documentacao':
            mainChartData = {
                labels: ['Aprovados', 'Pendências', 'Reprovados', 'Expirado', 'Pendente', 'Não Enviado'],
                datasets: [{
                    label: 'Documentação',
                    data: [
                        dados.geral.aprovados,
                        dados.geral.pendencias,
                        dados.geral.reprovados || 0,
                        dados.geral.expirado || 0,
                        dados.geral.pendente || 0,
                        dados.geral.naoEnviado || 0
                    ],
                    backgroundColor: [
                        '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#f97316', '#64748b'
                    ]
                }]
            };
            
            pieChartData = {
                labels: ['Loja', 'Web/Televendas'],
                datasets: [{
                    data: [dados.vendasLoja.total, dados.vendasWeb.total],
                    backgroundColor: ['#1e3a8a', '#7c3aed']
                }]
            };
            break;
            
        case 'app':
            mainChartData = {
                labels: ['Com App', 'Sem App', 'Cancelado', 'Outros'],
                datasets: [{
                    label: 'Status App',
                    data: [
                        dados.geral.sim,
                        dados.geral.nao,
                        dados.geral.cancelado || 0,
                        dados.geral.outros || 0
                    ],
                    backgroundColor: ['#10b981', '#ef4444', '#f59e0b', '#64748b']
                }]
            };
            
            // Gráfico de pizza para distribuição por origem
            if (dados.consultorasRetencao && dados.consultorasRetencao.length > 0) {
                pieChartData = {
                    labels: ['App Principal', 'Kamyla', 'Jhane'],
                    datasets: [{
                        data: [
                            dados.geral.total - (dados.consultorasRetencao.reduce((sum, c) => sum + c.total, 0)),
                            dados.consultorasRetencao.find(c => c.nome === 'KAMYLA')?.total || 0,
                            dados.consultorasRetencao.find(c => c.nome === 'JHANE')?.total || 0
                        ],
                        backgroundColor: ['#1e3a8a', '#7c3aed', '#ec4899']
                    }]
                };
            }
            break;
            
        case 'adimplencia':
            mainChartData = {
                labels: ['Mens. OK', 'Mens. Aberto', 'Mens. Atraso'],
                datasets: [{
                    label: 'Mensalidades',
                    data: [
                        dados.geral.mensOk,
                        dados.geral.mensAberto,
                        dados.geral.mensAtraso
                    ],
                    backgroundColor: ['#10b981', '#f59e0b', '#ef4444']
                }]
            };
            
            pieChartData = {
                labels: ['Aprovados', 'Pendentes'],
                datasets: [{
                    data: [dados.geral.aprovados, dados.geral.pendentes],
                    backgroundColor: ['#10b981', '#ef4444']
                }]
            };
            break;
            
        case 'recorrencia':
            // Para recorrência, mostrar gráfico comparativo
            const retencaoLabels = Object.keys(dados.retencao || {});
            const retencaoDados = retencaoLabels.map(c => dados.retencao[c].total3Meses?.totalOK || 0);
            
            mainChartData = {
                labels: retencaoLabels,
                datasets: [{
                    label: 'Retenção - OK (3 meses)',
                    data: retencaoDados,
                    backgroundColor: '#10b981'
                }]
            };
            
            const refiliacaoLabels = Object.keys(dados.refiliacao || {});
            const refiliacaoDados = refiliacaoLabels.map(c => dados.refiliacao[c].total3Meses?.totalOK || 0);
            
            pieChartData = {
                labels: refiliacaoLabels,
                datasets: [{
                    label: 'Refiliação - OK (3 meses)',
                    data: refiliacaoDados,
                    backgroundColor: ['#f59e0b', '#8b5cf6', '#ec4899']
                }]
            };
            break;
    }
    
    // Criar gráfico principal (barra)
    chartMain = new Chart(ctxMain, {
        type: 'bar',
        data: mainChartData,
        options: {
            responsive: true,
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: true,
                    text: 'Desempenho Detalhado'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return formatarNumero(value);
                        }
                    }
                }
            }
        }
    });
    
    // Criar gráfico de pizza
    chartPie = new Chart(ctxPie, {
        type: 'pie',
        data: pieChartData,
        options: {
            responsive: true,
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: true,
                    text: 'Distribuição'
                }
            }
        }
    });
}

// Atualizar tabela
function atualizarTabela(tipo, dados) {
    const tabela = document.getElementById('detalhes-table');
    const thead = tabela.querySelector('thead');
    const tbody = tabela.querySelector('tbody');
    
    // Limpar tabela
    thead.innerHTML = '';
    tbody.innerHTML = '';
    
    let headers = [];
    let rowsData = [];
    
    switch(tipo) {
        case 'documentacao':
            headers = ['Consultor', 'Setor', 'Total', 'Aprovados', 'Pendências', '% Aprovação'];
            rowsData = dados.consultores.map(c => [
                c.nome,
                c.setor,
                c.total,
                c.aprovados,
                c.pendencias,
                calcularPercentual(c.aprovados, c.total) + '%'
            ]);
            break;
            
        case 'app':
            headers = ['Consultor', 'Setor', 'Total', 'Com App', 'Sem App', 'Cancelado', '% App'];
            rowsData = dados.consultores.map(c => [
                c.nome,
                c.setor,
                c.total,
                c.sim,
                c.nao,
                c.cancelado || 0,
                calcularPercentual(c.sim, c.total) + '%'
            ]);
            
            // Adicionar Kamyla e Jhane se existirem
            if (dados.consultorasRetencao && dados.consultorasRetencao.length > 0) {
                dados.consultorasRetencao.forEach(c => {
                    rowsData.push([
                        c.nome + ' (RETENÇÃO)',
                        c.setor,
                        c.total,
                        c.sim,
                        c.nao,
                        c.cancelado || 0,
                        calcularPercentual(c.sim, c.total) + '%'
                    ]);
                });
            }
            break;
            
        case 'adimplencia':
            headers = ['Consultor', 'Total Trocas', 'Mens. OK', 'Mens. Aberto', 'Mens. Atraso', 'Aprovados', 'Pendentes', '% Aprovação'];
            rowsData = dados.consultores.map(c => [
                c.nome,
                c.totalTrocas,
                c.mensOk,
                c.mensAberto,
                c.mensAtraso,
                c.aprovados,
                c.pendentes,
                (c.percentualAprovado || 0) + '%'
            ]);
            break;
            
        case 'recorrencia':
            headers = ['Consultor', 'Tipo', 'Total 3M', 'OK', 'Em Aberto', 'Total OK', 'Pendências', '% OK'];
            
            // Adicionar retenção
            Object.keys(dados.retencao || {}).forEach(c => {
                const d = dados.retencao[c].total3Meses;
                rowsData.push([
                    c,
                    'RETENÇÃO',
                    d.totalRetido,
                    d.ok,
                    d.emAberto,
                    d.totalOK || (d.ok + d.emAberto),
                    d.pendencias,
                    (d.percentualOK || 0) + '%'
                ]);
            });
            
            // Adicionar refiliação
            Object.keys(dados.refiliacao || {}).forEach(c => {
                const d = dados.refiliacao[c].total3Meses;
                rowsData.push([
                    c,
                    'REFILIAÇÃO',
                    d.totalRetido,
                    d.ok,
                    d.emAberto,
                    d.totalOK || (d.ok + d.emAberto),
                    d.pendencias,
                    (d.percentualOK || 0) + '%'
                ]);
            });
            break;
    }
    
    // Criar cabeçalho
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    
    // Criar linhas de dados
    rowsData.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach((cell, index) => {
            const td = document.createElement('td');
            
            // Formatar números
            if (typeof cell === 'number') {
                td.textContent = formatarNumero(cell);
                if (index > 1) { // A partir da terceira coluna (dados numéricos)
                    td.style.fontWeight = '500';
                }
            } else {
                td.textContent = cell;
            }
            
            // Destacar percentuais
            if (cell.toString().includes('%')) {
                const valor = parseFloat(cell);
                if (valor >= 90) td.className = 'text-success fw-bold';
                else if (valor >= 80) td.className = 'text-warning fw-bold';
                else td.className = 'text-danger fw-bold';
            }
            
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

// Funções auxiliares
function calcularPercentual(parte, total) {
    if (!total || total === 0) return 0;
    return Math.round((parte / total) * 100);
}

function formatarNumero(num) {
    if (typeof num !== 'number') return num;
    return num.toLocaleString('pt-BR');
}

// Testar API
function testarAPI() {
    mostrarLoading('Testando conexão com API...');
    testarConexaoAPI().finally(() => {
        setTimeout(() => {
            esconderLoading();
            alert('Teste de conexão concluído!');
        }, 1000);
    });
}

// Exportar para Excel
function exportarParaExcel() {
    if (!dadosAtuais) {
        alert('Nenhum dado carregado para exportar.');
        return;
    }
    
    const tipo = document.getElementById('dashboardSelect').value;
    const mes = document.getElementById('mesSelect').value;
    const ano = document.getElementById('anoSelect').value;
    const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    
    // Criar conteúdo CSV
    let csvContent = `Dashboard ${tipo.toUpperCase()} - ${meses[mes-1]} ${ano}\n\n`;
    
    // Adicionar KPIs
    csvContent += "KPIs:\n";
    const kpiCards = document.querySelectorAll('.kpi-card');
    kpiCards.forEach(card => {
        const titulo = card.querySelector('.card-subtitle').textContent.trim();
        const valor = card.querySelector('.kpi-value').textContent.trim();
        csvContent += `${titulo},${valor}\n`;
    });
    
    csvContent += "\nDetalhamento:\n";
    
    // Adicionar tabela
    const tabela = document.getElementById('detalhes-table');
    const linhas = tabela.querySelectorAll('tr');
    
    linhas.forEach(linha => {
        const celulas = [];
        linha.querySelectorAll('th, td').forEach(celula => {
            celulas.push(`"${celula.textContent.replace(/"/g, '""')}"`);
        });
        csvContent += celulas.join(',') + '\n';
    });
    
    // Criar arquivo e fazer download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `dashboard_${tipo}_${mes}_${ano}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Event Listeners
document.getElementById('dashboardSelect').addEventListener('change', carregarDados);
document.getElementById('mesSelect').addEventListener('change', carregarDados);
document.getElementById('anoSelect').addEventListener('change', carregarDados);

// Exportar funções para uso global
window.carregarDados = carregarDados;
window.testarAPI = testarAPI;
window.exportarParaExcel = exportarParaExcel;
