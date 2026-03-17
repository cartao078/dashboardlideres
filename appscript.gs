/**
 * API UNIFICADA - DASHBOARDS V20.6 - COM RECORRÊNCIA VENDEDOR (SEM FILTRO DE MÊS)
 * VERSÃO COMPLETA COM TODOS OS 6 DASHBOARDS:
 * 1. Documentação | 2. App | 3. Adimplência | 4. Recorrência | 5. Refuturiza | 6. Recorrência Vendedor
 */

// ============================================================================
// CONFIGURAÇÕES GLOBAIS
// ============================================================================

const ID_PLANILHA_DADOS = "1amlKkGxw-RDZ30-5smDHtoyv1LRhelCaQ9BAhzLte3E";
const ID_PLANILHA_DASHBOARDS = "1CWmgLuUMU6XhiMOy1mZ20Gl2vcuAdzVkHr7x_x3pU00";

const CONSULTORES_MAP = {
  "VENDAS": [
    "FRANCISCO ROGEAN ALVES NASCIMENTO", 
    "LAYANE MACHADO DA CUNHA", "RAFAEL DOS SANTOS DE JESUS", 
    "VANESSA CRISTINA MARTINS SOUZA", "THAYNARA ARAUJO DE OLIVEIRA",
    "MARIA MADALENA SILVA FURTADO", "JOISCIANE DE SOUSA SILVA", "MARCUS LUIZ ARAUJO CUNHA"
  ],
  "RECEPCAO": [
    "THIAGO DA SILVA CARDOSO", "MILENA VITORIA LEMOS DA SILVA", 
    "DANIELLE LIMA BRITO"
  ],
  "REFILIACAO": [
    "FABIANA DA SILVA", "INGRID CARVALHO RODRIGUES", 
    "JENIFFER THAYNNA LIMA DA ROCHA", "WANESSA EVELYN CARVALHO OLIVEIRA MORAIS"
  ]
};

// CONSULTORAS DE RETENÇÃO
const CONSULTORAS_RETENCAO = {
  "MARISE": "MARISE"
};

const CONSULTORES_RECORRENCIA_RETENCAO = ['MARISE'];
const CONSULTORES_RECORRENCIA_REFILIACAO = ['JENIFFER', 'INGRID', 'WANESSA'];

// CONSULTORES RECORRÊNCIA VENDEDOR
const CONSULTORES_RECORRENCIA_VENDAS = {
  "VENDAS": [
    "FRANCISCO",
    "LAYANE", 
    "MARIA MADALENA",
    "RAFAEL",
    "THAYNARA",
    "VANESSA",
    "JOISCIANE",
    "MARCUS"
  ],
  "REFILIACAO": [
    "JENIFFER",
    "INGRID",
    "WANESSA"
  ],
  "RECEPCAO": [
    "THIAGO",
    "DANIELLE"
  ]
};

const NOME_ABA_RECORRENCIA_FONTE = "RECORRENCIA";
const NOME_ABA_RECORRENCIA_VENDAS = "RECORRENCIA VENDAS";

const OUTROS_LISTA = ["KATIANE FERREIRA DOS SANTOS", "CARTÃO DE CEILANDIA", "FLÁVIA ARAÚJO MARQUES", "ALYNNE GABRIELLE DO NASCIMENTO RIBEIRO", 
                      "CLEONICE CONCEICAO DE MARIA", "EDUARDO VICENTE PEREIRA JUNIOR", "ESTEFFANNIE MENDES DA SILVA", "IGOR MARTINS DE SOUSA LIMA", 
                      "JHANE ELY SANTOS DE OLIVEIRA", "KAMYLA MORAES DE OLIVEIRA", "KATIA REGINA CAMELO DA SILVA", "RHYAN CARLOS FRANCO PEREIRA", 
                      "TAYNARA RODRIGUES DA CRUZ", "THARIK MENEZES SANTIAGO"];
const PREFIXO_AMOR = "AMOR SAUDE CEILANDIA - ";

const MESES = [
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
];

// ============================================================================
// FUNÇÃO PRINCIPAL - doGet() PARA API
// ============================================================================

function doGet(e) {
  try {
    const endpoint = e.parameter.endpoint || "documentacao";
    const mes = parseInt(e.parameter.mes) || new Date().getMonth() + 1;
    const ano = parseInt(e.parameter.ano) || new Date().getFullYear();

    let resultado;

    if (endpoint === "documentacao") {
      resultado = getDocumentacaoData(mes, ano);
    } else if (endpoint === "app") {
      resultado = getAppData(mes, ano);
    } else if (endpoint === "adimplencia") {
      resultado = getAdimplenciaData(mes, ano);
    } else if (endpoint === "recorrencia") {
      const mesRec = parseInt(e.parameter.mes) || new Date().getMonth() + 1;
      const anoRec = parseInt(e.parameter.ano) || new Date().getFullYear();
      resultado = processarDadosRecorrenciaPorMes(mesRec, anoRec, true);
    } else if (endpoint === "refuturiza") {
      resultado = getRefuturizaData(mes, ano);
    } else if (endpoint === "recorrencia_vendedor") {
      resultado = getRecorrenciaVendedorData(); // SEM parâmetros mes/ano
    } else {
      resultado = {
        status: "error",
        error: "Endpoint nao encontrado: " + endpoint,
        endpoints_disponiveis: ["documentacao", "app", "adimplencia", "recorrencia", "refuturiza", "recorrencia_vendedor"]
      };
    }

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================================
// MENU CUSTOMIZADO V20.6 - COM TODOS OS 6 DASHBOARDS FUNCIONANDO
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu("📊 DASHBOARDS V20.6")
    .addItem("🔄 Atualizar Todos (Mês Atual)", "atualizarTodosDashboards")
    .addSeparator()
    .addSubMenu(ui.createMenu("📁 Documentação")
      .addItem("Janeiro", "abrirDashboardDocJaneiro")
      .addItem("Fevereiro", "abrirDashboardDocFevereiro")
      .addItem("Março", "abrirDashboardDocMarco")
      .addItem("Abril", "abrirDashboardDocAbril")
      .addItem("Maio", "abrirDashboardDocMaio")
      .addItem("Junho", "abrirDashboardDocJunho")
      .addItem("Julho", "abrirDashboardDocJulho")
      .addItem("Agosto", "abrirDashboardDocAgosto")
      .addItem("Setembro", "abrirDashboardDocSetembro")
      .addItem("Outubro", "abrirDashboardDocOutubro")
      .addItem("Novembro", "abrirDashboardDocNovembro")
      .addItem("Dezembro", "abrirDashboardDocDezembro"))
    .addSubMenu(ui.createMenu("📱 App")
      .addItem("Janeiro", "abrirDashboardAppJaneiro")
      .addItem("Fevereiro", "abrirDashboardAppFevereiro")
      .addItem("Março", "abrirDashboardAppMarco")
      .addItem("Abril", "abrirDashboardAppAbril")
      .addItem("Maio", "abrirDashboardAppMaio")
      .addItem("Junho", "abrirDashboardAppJunho")
      .addItem("Julho", "abrirDashboardAppJulho")
      .addItem("Agosto", "abrirDashboardAppAgosto")
      .addItem("Setembro", "abrirDashboardAppSetembro")
      .addItem("Outubro", "abrirDashboardAppOutubro")
      .addItem("Novembro", "abrirDashboardAppNovembro")
      .addItem("Dezembro", "abrirDashboardAppDezembro"))
    .addSubMenu(ui.createMenu("💳 Adimplência")
      .addItem("Janeiro", "abrirDashboardAdimJaneiro")
      .addItem("Fevereiro", "abrirDashboardAdimFevereiro")
      .addItem("Março", "abrirDashboardAdimMarco")
      .addItem("Abril", "abrirDashboardAdimAbril")
      .addItem("Maio", "abrirDashboardAdimMaio")
      .addItem("Junho", "abrirDashboardAdimJunho")
      .addItem("Julho", "abrirDashboardAdimJulho")
      .addItem("Agosto", "abrirDashboardAdimAgosto")
      .addItem("Setembro", "abrirDashboardAdimSetembro")
      .addItem("Outubro", "abrirDashboardAdimOutubro")
      .addItem("Novembro", "abrirDashboardAdimNovembro")
      .addItem("Dezembro", "abrirDashboardAdimDezembro"))
    .addSubMenu(ui.createMenu("📈 Recorrência")
      .addItem("Janeiro", "abrirDashboardRecorrenciaJaneiro")
      .addItem("Fevereiro", "abrirDashboardRecorrenciaFevereiro")
      .addItem("Março", "abrirDashboardRecorrenciaMarco")
      .addItem("Abril", "abrirDashboardRecorrenciaAbril")
      .addItem("Maio", "abrirDashboardRecorrenciaMaio")
      .addItem("Junho", "abrirDashboardRecorrenciaJunho")
      .addItem("Julho", "abrirDashboardRecorrenciaJulho")
      .addItem("Agosto", "abrirDashboardRecorrenciaAgosto")
      .addItem("Setembro", "abrirDashboardRecorrenciaSetembro")
      .addItem("Outubro", "abrirDashboardRecorrenciaOutubro")
      .addItem("Novembro", "abrirDashboardRecorrenciaNovembro")
      .addItem("Dezembro", "abrirDashboardRecorrenciaDezembro")
      .addSeparator()
      .addItem("Atualizar Mês Atual", "gerarDashboardRecorrencia")
      .addItem("Histórico Detalhado", "gerarDashboardRecorrenciaHistorico"))
    .addSubMenu(ui.createMenu("🤝 Recorrência Vendedor")
      .addItem("📊 Abrir Dashboard (TODOS OS DADOS)", "abrirDashboardRecorrenciaVendedor")) // APENAS 1 OPÇÃO
    .addSubMenu(ui.createMenu("🔄 Refuturiza")
      .addItem("Janeiro", "abrirDashboardRefuturizaJaneiro")
      .addItem("Fevereiro", "abrirDashboardRefuturizaFevereiro")
      .addItem("Março", "abrirDashboardRefuturizaMarco")
      .addItem("Abril", "abrirDashboardRefuturizaAbril")
      .addItem("Maio", "abrirDashboardRefuturizaMaio")
      .addItem("Junho", "abrirDashboardRefuturizaJunho")
      .addItem("Julho", "abrirDashboardRefuturizaJulho")
      .addItem("Agosto", "abrirDashboardRefuturizaAgosto")
      .addItem("Setembro", "abrirDashboardRefuturizaSetembro")
      .addItem("Outubro", "abrirDashboardRefuturizaOutubro")
      .addItem("Novembro", "abrirDashboardRefuturizaNovembro")
      .addItem("Dezembro", "abrirDashboardRefuturizaDezembro"))
    .addSeparator()
    .addItem("🔧 Gerar Mês Passado", "gerarDashboardMesPassado")
    .addItem("⚙️ Configurações", "showConfigDialog")
    .addItem("🚀 Deploy API", "deployAPI")
    .addToUi();
}

// ============================================================================
// FUNÇÕES DE ABERTURA DE DASHBOARDS - CORRIGIDAS
// ============================================================================

// Documentação (12 funções)
function abrirDashboardDocJaneiro() { abrirDashboardPorMesETipo(1, "documentacao"); }
function abrirDashboardDocFevereiro() { abrirDashboardPorMesETipo(2, "documentacao"); }
function abrirDashboardDocMarco() { abrirDashboardPorMesETipo(3, "documentacao"); }
function abrirDashboardDocAbril() { abrirDashboardPorMesETipo(4, "documentacao"); }
function abrirDashboardDocMaio() { abrirDashboardPorMesETipo(5, "documentacao"); }
function abrirDashboardDocJunho() { abrirDashboardPorMesETipo(6, "documentacao"); }
function abrirDashboardDocJulho() { abrirDashboardPorMesETipo(7, "documentacao"); }
function abrirDashboardDocAgosto() { abrirDashboardPorMesETipo(8, "documentacao"); }
function abrirDashboardDocSetembro() { abrirDashboardPorMesETipo(9, "documentacao"); }
function abrirDashboardDocOutubro() { abrirDashboardPorMesETipo(10, "documentacao"); }
function abrirDashboardDocNovembro() { abrirDashboardPorMesETipo(11, "documentacao"); }
function abrirDashboardDocDezembro() { abrirDashboardPorMesETipo(12, "documentacao"); }

// App (12 funções)
function abrirDashboardAppJaneiro() { abrirDashboardPorMesETipo(1, "app"); }
function abrirDashboardAppFevereiro() { abrirDashboardPorMesETipo(2, "app"); }
function abrirDashboardAppMarco() { abrirDashboardPorMesETipo(3, "app"); }
function abrirDashboardAppAbril() { abrirDashboardPorMesETipo(4, "app"); }
function abrirDashboardAppMaio() { abrirDashboardPorMesETipo(5, "app"); }
function abrirDashboardAppJunho() { abrirDashboardPorMesETipo(6, "app"); }
function abrirDashboardAppJulho() { abrirDashboardPorMesETipo(7, "app"); }
function abrirDashboardAppAgosto() { abrirDashboardPorMesETipo(8, "app"); }
function abrirDashboardAppSetembro() { abrirDashboardPorMesETipo(9, "app"); }
function abrirDashboardAppOutubro() { abrirDashboardPorMesETipo(10, "app"); }
function abrirDashboardAppNovembro() { abrirDashboardPorMesETipo(11, "app"); }
function abrirDashboardAppDezembro() { abrirDashboardPorMesETipo(12, "app"); }

// Adimplência (12 funções)
function abrirDashboardAdimJaneiro() { abrirDashboardPorMesETipo(1, "adimplencia"); }
function abrirDashboardAdimFevereiro() { abrirDashboardPorMesETipo(2, "adimplencia"); }
function abrirDashboardAdimMarco() { abrirDashboardPorMesETipo(3, "adimplencia"); }
function abrirDashboardAdimAbril() { abrirDashboardPorMesETipo(4, "adimplencia"); }
function abrirDashboardAdimMaio() { abrirDashboardPorMesETipo(5, "adimplencia"); }
function abrirDashboardAdimJunho() { abrirDashboardPorMesETipo(6, "adimplencia"); }
function abrirDashboardAdimJulho() { abrirDashboardPorMesETipo(7, "adimplencia"); }
function abrirDashboardAdimAgosto() { abrirDashboardPorMesETipo(8, "adimplencia"); }
function abrirDashboardAdimSetembro() { abrirDashboardPorMesETipo(9, "adimplencia"); }
function abrirDashboardAdimOutubro() { abrirDashboardPorMesETipo(10, "adimplencia"); }
function abrirDashboardAdimNovembro() { abrirDashboardPorMesETipo(11, "adimplencia"); }
function abrirDashboardAdimDezembro() { abrirDashboardPorMesETipo(12, "adimplencia"); }

// Recorrência (12 funções)
function abrirDashboardRecorrenciaJaneiro() { abrirDashboardRecorrenciaPorMes(1); }
function abrirDashboardRecorrenciaFevereiro() { abrirDashboardRecorrenciaPorMes(2); }
function abrirDashboardRecorrenciaMarco() { abrirDashboardRecorrenciaPorMes(3); }
function abrirDashboardRecorrenciaAbril() { abrirDashboardRecorrenciaPorMes(4); }
function abrirDashboardRecorrenciaMaio() { abrirDashboardRecorrenciaPorMes(5); }
function abrirDashboardRecorrenciaJunho() { abrirDashboardRecorrenciaPorMes(6); }
function abrirDashboardRecorrenciaJulho() { abrirDashboardRecorrenciaPorMes(7); }
function abrirDashboardRecorrenciaAgosto() { abrirDashboardRecorrenciaPorMes(8); }
function abrirDashboardRecorrenciaSetembro() { abrirDashboardRecorrenciaPorMes(9); }
function abrirDashboardRecorrenciaOutubro() { abrirDashboardRecorrenciaPorMes(10); }
function abrirDashboardRecorrenciaNovembro() { abrirDashboardRecorrenciaPorMes(11); }
function abrirDashboardRecorrenciaDezembro() { abrirDashboardRecorrenciaPorMes(12); }

// Recorrência Vendedor (APENAS 1 função - SEM FILTRO DE MÊS)
function abrirDashboardRecorrenciaVendedor() { 
  abrirDashboardPorMesETipo(0, "recorrencia_vendedor"); // Usa 0 para indicar "todos os meses"
}

// Refuturiza (12 funções)
function abrirDashboardRefuturizaJaneiro() { abrirDashboardPorMesETipo(1, "refuturiza"); }
function abrirDashboardRefuturizaFevereiro() { abrirDashboardPorMesETipo(2, "refuturiza"); }
function abrirDashboardRefuturizaMarco() { abrirDashboardPorMesETipo(3, "refuturiza"); }
function abrirDashboardRefuturizaAbril() { abrirDashboardPorMesETipo(4, "refuturiza"); }
function abrirDashboardRefuturizaMaio() { abrirDashboardPorMesETipo(5, "refuturiza"); }
function abrirDashboardRefuturizaJunho() { abrirDashboardPorMesETipo(6, "refuturiza"); }
function abrirDashboardRefuturizaJulho() { abrirDashboardPorMesETipo(7, "refuturiza"); }
function abrirDashboardRefuturizaAgosto() { abrirDashboardPorMesETipo(8, "refuturiza"); }
function abrirDashboardRefuturizaSetembro() { abrirDashboardPorMesETipo(9, "refuturiza"); }
function abrirDashboardRefuturizaOutubro() { abrirDashboardPorMesETipo(10, "refuturiza"); }
function abrirDashboardRefuturizaNovembro() { abrirDashboardPorMesETipo(11, "refuturiza"); }
function abrirDashboardRefuturizaDezembro() { abrirDashboardPorMesETipo(12, "refuturiza"); }

// ============================================================================
// FUNÇÕES PRINCIPAIS DE ABERTURA - CORRIGIDAS
// ============================================================================

function abrirDashboardPorMesETipo(mes, tipo) {
  try {
    const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
    let sheetName = "";
    let sheet = null;
    
    if (tipo === "documentacao") {
      sheetName = "Dashboard Documentação";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) {
        const ano = new Date().getFullYear();
        getDocumentacaoData(mes, ano);
        sheet = ssDash.getSheetByName(sheetName);
      }
    } else if (tipo === "app") {
      sheetName = "Dashboard App";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) {
        const ano = new Date().getFullYear();
        getAppData(mes, ano);
        sheet = ssDash.getSheetByName(sheetName);
      }
    } else if (tipo === "adimplencia") {
      sheetName = "Dashboard Adimplência";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) {
        const ano = new Date().getFullYear();
        getAdimplenciaData(mes, ano);
        sheet = ssDash.getSheetByName(sheetName);
      }
    } else if (tipo === "recorrencia") {
      sheetName = "Dashboard Recorrência";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) {
        gerarDashboardRecorrencia();
        sheet = ssDash.getSheetByName(sheetName);
      }
    } else if (tipo === "recorrencia_vendedor") {
      sheetName = "Dashboard Recorrência Vendedor";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) {
        getRecorrenciaVendedorData(); // SEM parâmetros mes/ano
        sheet = ssDash.getSheetByName(sheetName);
      }
    } else if (tipo === "refuturiza") {
      sheetName = "Dashboard Refuturiza";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) {
        const ano = new Date().getFullYear();
        getRefuturizaData(mes, ano);
        sheet = ssDash.getSheetByName(sheetName);
      }
    }
    
    if (!sheet) {
      throw new Error(`Aba ${sheetName} não encontrada`);
    }
    
    const url = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_DASHBOARDS}/edit#gid=${sheet.getSheetId()}`;
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank'); google.script.host.close();</script>`),
      'Abrindo Dashboard...'
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Erro ao abrir dashboard: ${error.message}`);
  }
}

function abrirDashboardRecorrenciaPorMes(mes) {
  try {
    const ano = new Date().getFullYear();
    const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
    const nomeAba = `Recorrência ${MESES[mes-1]} ${ano}`;
    
    let sheet = ssDash.getSheetByName(nomeAba);
    
    if (!sheet) {
      const dados = processarDadosRecorrenciaPorMes(mes, ano, true);
      if (dados.status === "success") {
        renderizarPlanilhaRecorrenciaPorMes(dados.data, mes, ano);
        sheet = ssDash.getSheetByName(nomeAba);
      } else {
        throw new Error(dados.error);
      }
    }
    
    const url = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_DASHBOARDS}/edit#gid=${sheet.getSheetId()}`;
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank'); google.script.host.close();</script>`),
      'Abrindo Dashboard Recorrência...'
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Erro ao abrir dashboard: ${error.message}`);
  }
}

// ============================================================================
// NOVA FUNÇÃO: GERAR DASHBOARD PARA MÊS PASSADO ESPECÍFICO
// ============================================================================

function gerarDashboardMesPassado() {
  try {
    const ui = SpreadsheetApp.getUi();
    const anoAtual = new Date().getFullYear();
    
    // Criar interface para selecionar dashboard e mês
    const html = `
      <div style="padding: 20px; font-family: Arial;">
        <h2>🔧 Gerar Dashboard para Mês Específico</h2>
        
        <div style="margin: 15px 0;">
          <label for="dashboardType"><strong>Selecione o Dashboard:</strong></label><br>
          <select id="dashboardType" style="width: 100%; padding: 8px; margin: 10px 0;">
            <option value="documentacao">📁 Documentação</option>
            <option value="app">📱 App</option>
            <option value="adimplencia">💳 Adimplência</option>
            <option value="recorrencia">📈 Recorrência</option>
            <option value="recorrencia_vendedor">🤝 Recorrência Vendedor (TODOS OS MESES)</option>
            <option value="refuturiza">🔄 Refuturiza</option>
          </select>
        </div>
        
        <div style="margin: 15px 0;" id="monthYearSection">
          <label for="mes"><strong>Selecione o Mês:</strong></label><br>
          <select id="mes" style="width: 100%; padding: 8px; margin: 10px 0;">
            <option value="1">Janeiro</option>
            <option value="2">Fevereiro</option>
            <option value="3">Março</option>
            <option value="4">Abril</option>
            <option value="5">Maio</option>
            <option value="6">Junho</option>
            <option value="7">Julho</option>
            <option value="8">Agosto</option>
            <option value="9">Setembro</option>
            <option value="10">Outubro</option>
            <option value="11">Novembro</option>
            <option value="12">Dezembro</option>
          </select>
        </div>
        
        <div style="margin: 15px 0;" id="yearSection">
          <label for="ano"><strong>Ano:</strong></label><br>
          <input type="number" id="ano" value="${anoAtual}" style="width: 100%; padding: 8px; margin: 10px 0;">
        </div>
        
        <div style="margin: 20px 0; padding: 10px; background: #f0f0f0; border-radius: 5px;">
          <strong>📋 Resumo:</strong><br>
          Dashboard: <span id="dashboardPreview">Documentação</span><br>
          Mês/Ano: <span id="mesAnoPreview">Janeiro ${anoAtual}</span>
        </div>
        
        <button onclick="gerarDashboard()" style="background: #10b981; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">Gerar Dashboard</button>
        <button onclick="google.script.host.close()" style="background: #ef4444; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">Cancelar</button>
        
        <script>
          function updatePreview() {
            const dashboardType = document.getElementById('dashboardType').value;
            const mes = parseInt(document.getElementById('mes').value);
            const ano = document.getElementById('ano').value;
            const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
            
            const dashboardNames = {
              'documentacao': 'Documentação',
              'app': 'App',
              'adimplencia': 'Adimplência',
              'recorrencia': 'Recorrência',
              'recorrencia_vendedor': 'Recorrência Vendedor',
              'refuturiza': 'Refuturiza'
            };
            
            document.getElementById('dashboardPreview').textContent = dashboardNames[dashboardType];
            
            if (dashboardType === 'recorrencia_vendedor') {
              document.getElementById('mesAnoPreview').textContent = 'TODOS OS MESES (dados completos)';
            } else {
              document.getElementById('mesAnoPreview').textContent = meses[mes-1] + ' ' + ano;
            }
          }
          
          document.getElementById('dashboardType').addEventListener('change', function() {
            const isRecorrenciaVendedor = this.value === 'recorrencia_vendedor';
            document.getElementById('monthYearSection').style.display = isRecorrenciaVendedor ? 'none' : 'block';
            document.getElementById('yearSection').style.display = isRecorrenciaVendedor ? 'none' : 'block';
            updatePreview();
          });
          
          document.getElementById('mes').addEventListener('change', updatePreview);
          document.getElementById('ano').addEventListener('input', updatePreview);
          
          updatePreview();
          
          function gerarDashboard() {
            const dashboardType = document.getElementById('dashboardType').value;
            
            if (dashboardType === 'recorrencia_vendedor') {
              google.script.run
                .withSuccessHandler(function(result) {
                  alert(result);
                  google.script.host.close();
                })
                .withFailureHandler(function(error) {
                  alert('Erro: ' + error.message);
                })
                .processarDashboardEspecifico(dashboardType, 0, 0);
            } else {
              const mes = parseInt(document.getElementById('mes').value);
              const ano = parseInt(document.getElementById('ano').value);
              
              google.script.run
                .withSuccessHandler(function(result) {
                  alert(result);
                  google.script.host.close();
                })
                .withFailureHandler(function(error) {
                  alert('Erro: ' + error.message);
                })
                .processarDashboardEspecifico(dashboardType, mes, ano);
            }
          }
        </script>
      </div>
    `;
    
    ui.showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(400).setHeight(500),
      'Gerar Dashboard Específico'
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Erro: ${error.message}`);
  }
}

function processarDashboardEspecifico(tipo, mes, ano) {
  try {
    let resultado;
    
    switch(tipo) {
      case 'documentacao':
        resultado = getDocumentacaoData(mes, ano);
        break;
      case 'app':
        resultado = getAppData(mes, ano);
        break;
      case 'adimplencia':
        resultado = getAdimplenciaData(mes, ano);
        break;
      case 'recorrencia':
        resultado = processarDadosRecorrenciaPorMes(mes, ano, true);
        if (resultado.status === "success") {
          renderizarPlanilhaRecorrenciaPorMes(resultado.data, mes, ano);
        }
        break;
      case 'recorrencia_vendedor':
        resultado = getRecorrenciaVendedorData(); // SEM parâmetros mes/ano
        break;
      case 'refuturiza':
        resultado = getRefuturizaData(mes, ano);
        break;
      default:
        return `❌ Tipo de dashboard desconhecido: ${tipo}`;
    }
    
    if (resultado && resultado.status === "success") {
      if (tipo === 'recorrencia_vendedor') {
        return `✅ Dashboard Recorrência Vendedor (TODOS OS DADOS) gerado com sucesso!`;
      } else {
        return `✅ Dashboard ${tipo} de ${MESES[mes-1]} ${ano} gerado com sucesso!`;
      }
    } else if (resultado && resultado.status === "error") {
      return `❌ Erro: ${resultado.error}`;
    } else {
      if (tipo === 'recorrencia_vendedor') {
        return `✅ Dashboard Recorrência Vendedor (TODOS OS DADOS) processado com sucesso!`;
      } else {
        return `✅ Dashboard ${tipo} de ${MESES[mes-1]} ${ano} processado com sucesso!`;
      }
    }
    
  } catch (error) {
    return `❌ Erro: ${error.message}`;
  }
}

// ============================================================================
// FUNÇÃO AUXILIAR: NORMALIZAR STATUS DO APP
// ============================================================================

function normalizarAppStatus(status) {
  if (!status || status === "") return "OUTROS";
  
  const statusUpper = status.toString().toUpperCase().trim();
  
  if (["SIM", "S", "YES", "Y", "COM APP", "BAIXADO", "INSTALADO", "✅", "OK", "CONCLUÍDO", "CONCLUIDO"].includes(statusUpper)) {
    return "SIM";
  }
  
  if (["NÃO", "NAO", "N", "NO", "SEM APP", "NÃO BAIXADO", "NÃO INSTALADO", "❌", "NEGATIVO", "NA"].includes(statusUpper)) {
    return "NAO";
  }
  
  if (["CANCELADO", "CANCEL", "C", "CANCELADA", "CANCELADOS", "CANC"].includes(statusUpper)) {
    return "CANCELADO";
  }
  
  return "OUTROS";
}

// ============================================================================
// FUNÇÃO AUXILIAR: NORMALIZAR STATUS DE LIGAÇÃO (REFUTURIZA)
// ============================================================================

function normalizarStatusLigacao(status) {
  if (!status || status === "") return "SEM LIGAÇÃO";
  
  const statusUpper = status.toString().toUpperCase().trim();
  
  if (["SIM", "S", "YES", "Y", "COM LIGAÇÃO", "LIGADO", "REALIZADO", "CONCLUÍDO", "CONCLUIDO", "✅", "OK", "FEITO", "FEITA"].includes(statusUpper)) {
    return "COM LIGAÇÃO";
  }
  
  if (["NÃO", "NAO", "N", "NO", "SEM LIGAÇÃO", "NÃO LIGADO", "NÃO REALIZADO", "PENDENTE", "❌", "NEGATIVO", "NA"].includes(statusUpper)) {
    return "SEM LIGAÇÃO";
  }
  
  if (["CANCELADO", "CANCEL", "C", "CANCELADA", "CANCELADOS", "CANC", "DESISTENTE", "DESISTIU"].includes(statusUpper)) {
    return "CANCELADO";
  }
  
  if (statusUpper.includes("LIG") || statusUpper.includes("CALL") || statusUpper.includes("TELEFONE")) {
    return "COM LIGAÇÃO";
  }
  
  return "SEM LIGAÇÃO";
}

// ============================================================================
// ENDPOINT: DOCUMENTAÇÃO
// ============================================================================

function getDocumentacaoData(mes, ano) {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    const sheet = ssDados.getSheetByName("documentação");
    if (!sheet) return createErrorResponse("Sheet 'documentação' não encontrada");

    const data = sheet.getDataRange().getValues();
    const headers = data[1];
    const idx = {
      consultor: headers.indexOf("CONSULTOR"),
      data: headers.indexOf("DATA"),
      kyc: headers.indexOf("KYC")
    };

    const mesAtual = mes || new Date().getMonth() + 1;
    const anoAtual = ano || new Date().getFullYear();

    const consultoresData = {};
    let totalGeral = { total: 0, aprovados: 0, pendencias: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0 };
    let vendasLoja = { total: 0, aprovados: 0, pendencias: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0 };
    let vendasWeb = { total: 0, aprovados: 0, pendencias: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0 };

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const dataParsed = parseDate(row[idx.data]);
      if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;

      const consultorOriginal = String(row[idx.consultor] || "").trim();
      const consultorNormalizado = normalizarConsultorResidual(consultorOriginal);
      const kyc = String(row[idx.kyc] || "").trim().toUpperCase();
      
      let setor = "TELEVENDAS";
      if (consultorNormalizado === "WEB SITE") setor = "WEB SITE";
      else if (consultorNormalizado === "OUTROS") setor = "OUTROS";
      else if (consultorNormalizado === "TELEVENDAS") setor = "TELEVENDAS";
      else {
        for (const s in CONSULTORES_MAP) {
          if (CONSULTORES_MAP[s].some(c => consultorNormalizado.toUpperCase().includes(c.toUpperCase()))) {
            setor = s;
            break;
          }
        }
      }

      const isWebTelevendas = (consultorNormalizado === "WEB SITE" || consultorNormalizado === "TELEVENDAS" || consultorNormalizado === "OUTROS");

      if (!consultoresData[consultorNormalizado]) {
        consultoresData[consultorNormalizado] = { 
          nome: consultorNormalizado, 
          total: 0, aprovados: 0, pendencias: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0,
          setor: setor
        };
      }

      const c = consultoresData[consultorNormalizado];
      c.total++;
      totalGeral.total++;
      
      if (isWebTelevendas) vendasWeb.total++; else vendasLoja.total++;

      if (kyc === "APROVADO") {
        c.aprovados++;
        totalGeral.aprovados++;
        if (isWebTelevendas) vendasWeb.aprovados++; else vendasLoja.aprovados++;
      } else {
        c.pendencias++;
        totalGeral.pendencias++;
        if (isWebTelevendas) vendasWeb.pendencias++; else vendasLoja.pendencias++;
        
        if (kyc === "REPROVADO") { 
          c.reprovados++; totalGeral.reprovados++; 
          if (isWebTelevendas) vendasWeb.reprovados++; else vendasLoja.reprovados++;
        }
        else if (kyc === "EXPIRADO") { 
          c.expirado++; totalGeral.expirado++; 
          if (isWebTelevendas) vendasWeb.expirado++; else vendasLoja.expirado++;
        }
        else if (kyc === "PENDENTE") { 
          c.pendente++; totalGeral.pendente++; 
          if (isWebTelevendas) vendasWeb.pendente++; else vendasLoja.pendente++;
        }
        else { 
          c.naoEnviado++; totalGeral.naoEnviado++; 
          if (isWebTelevendas) vendasWeb.naoEnviado++; else vendasLoja.naoEnviado++;
        }
      }
    }

    const responseData = {
      mes: MESES[mesAtual - 1],
      ano: anoAtual,
      geral: totalGeral,
      vendasLoja: vendasLoja,
      vendasWeb: vendasWeb,
      consultores: Object.values(consultoresData)
    };

    criarDashboardDocumentacao(responseData);
    return createSuccessResponse(responseData);
  } catch (error) {
    return createErrorResponse(error.message);
  }
}

function criarDashboardDocumentacao(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Documentação");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Documentação");
  sheet.clear();

  sheet.getRange("A1:L1").merge().setValue("DASHBOARD DOCUMENTAÇÃO - " + dados.mes.toUpperCase() + " " + dados.ano)
    .setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  renderizarTresCardsPrincipais(sheet, dados, "DOC");

  const setores = ["VENDAS", "RECEPCAO", "REFILIACAO", "WEB SITE", "TELEVENDAS", "OUTROS"];
  let currentRow = 14;

  setores.forEach(setor => {
    const consultoresSetor = dados.consultores.filter(c => c.setor === setor);
    if (consultoresSetor.length === 0) return;

    sheet.getRange(currentRow, 1, 1, 12).merge().setValue("SETOR: " + setor)
      .setBackground("#4b5563").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("left");
    currentRow += 2;

    let col = 1;
    consultoresSetor.sort((a, b) => a.nome.localeCompare(b.nome)).forEach(c => {
      const p = c.total > 0 ? Math.round((c.aprovados / c.total) * 100) : 0;
      const card = [
        [c.nome, ""],
        ["TOTAL", c.total],
        ["APROVADOS", c.aprovados],
        ["PENDÊNCIAS", c.pendencias],
        ["  NÃO ENVIADO", c.naoEnviado],
        ["  EXPIRADO", c.expirado],
        ["  PENDENTE", c.pendente],
        ["% APROVADOS", p + "%"]
      ];
      
      sheet.getRange(currentRow, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
      sheet.getRange(currentRow, col, 1, 2).merge().setBackground("#059669").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      
      const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(currentRow + 7, col + 1).setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");

      col += 3;
      if (col > 9) { col = 1; currentRow += 10; }
    });
    
    if (col !== 1) currentRow += 10; else currentRow += 2;
  });
  
  for (let i = 1; i <= 12; i++) sheet.setColumnWidth(i, 160);
}

// ============================================================================
// ENDPOINT: APP
// ============================================================================

function getAppData(mes, ano) {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    
    const sheetPrincipal = ssDados.getSheetByName("app");
    if (!sheetPrincipal) return createErrorResponse("Sheet 'app' não encontrada");

    const dataPrincipal = sheetPrincipal.getDataRange().getValues();
    const headersPrincipal = dataPrincipal[1];
    const idxPrincipal = {
      consultor: headersPrincipal.indexOf("CONSULTOR"),
      data: headersPrincipal.indexOf("DATA"),
      app: headersPrincipal.indexOf("APP BAIXADO")
    };

    const mesAtual = mes || new Date().getMonth() + 1;
    const anoAtual = ano || new Date().getFullYear();

    const consultoresData = {};
    let totalGeral = { total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0 };
    let appLoja = { total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0 };
    let appWeb = { total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0 };

    for (let i = 2; i < dataPrincipal.length; i++) {
      const row = dataPrincipal[i];
      const dataParsed = parseDate(row[idxPrincipal.data]);
      if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;

      const consultorOriginal = String(row[idxPrincipal.consultor] || "").trim();
      const consultorNormalizado = normalizarConsultorResidual(consultorOriginal);
      
      const appRaw = String(row[idxPrincipal.app] || "").trim();
      const app = normalizarAppStatus(appRaw);

      let setor = "TELEVENDAS";
      if (consultorNormalizado === "WEB SITE") setor = "WEB SITE";
      else if (consultorNormalizado === "OUTROS") setor = "OUTROS";
      else if (consultorNormalizado === "TELEVENDAS") setor = "TELEVENDAS";
      else {
        for (const s in CONSULTORES_MAP) {
          if (CONSULTORES_MAP[s].some(c => consultorNormalizado.toUpperCase().includes(c.toUpperCase()))) {
            setor = s;
            break;
          }
        }
      }

      const isWebTelevendas = (consultorNormalizado === "WEB SITE" || consultorNormalizado === "TELEVENDAS" || consultorNormalizado === "OUTROS");

      if (!consultoresData[consultorNormalizado]) {
        consultoresData[consultorNormalizado] = { 
          nome: consultorNormalizado, 
          total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0,
          setor: setor,
          origem: "app"
        };
      }

      const c = consultoresData[consultorNormalizado];
      c.total++;
      totalGeral.total++;
      
      if (isWebTelevendas) appWeb.total++; else appLoja.total++;

      if (app === "SIM") { 
        c.sim++; totalGeral.sim++; 
        if (isWebTelevendas) appWeb.sim++; else appLoja.sim++;
      }
      else if (app === "NAO") { 
        c.nao++; totalGeral.nao++; 
        if (isWebTelevendas) appWeb.nao++; else appLoja.nao++;
      }
      else if (app === "CANCELADO") { 
        c.cancelado++; totalGeral.cancelado++; 
        if (isWebTelevendas) appWeb.cancelado++; else appLoja.cancelado++;
      }
      else {
        c.outros++; totalGeral.outros++;
        if (isWebTelevendas) appWeb.outros++; else appLoja.outros++;
      }
    }

    const retencaoData = {};
    
    const sheetRetencao = ssDados.getSheetByName("APP RETENÇÃO") || 
                         ssDados.getSheetByName("APP RETENCAO") ||
                         ssDados.getSheetByName("App Retenção");
    
    if (sheetRetencao) {
      const dataRetencao = sheetRetencao.getDataRange().getValues();
      const headersRetencao = dataRetencao[1];
      
      const idxRetencao = {
        consultor: headersRetencao.indexOf("CONSULTOR"),
        data: headersRetencao.indexOf("DATA"),
        app: headersRetencao.indexOf("APP BAIXADO")
      };
      
      if (idxRetencao.consultor !== -1 && idxRetencao.data !== -1 && idxRetencao.app !== -1) {
        for (let i = 2; i < dataRetencao.length; i++) {
          const row = dataRetencao[i];
          const dataParsed = parseDate(row[idxRetencao.data]);
          if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;
          
          const consultorOriginal = String(row[idxRetencao.consultor] || "").trim().toUpperCase();
          
          let consultora = null;
          for (const nome in CONSULTORAS_RETENCAO) {
            if (consultorOriginal.includes(nome.toUpperCase())) {
              consultora = nome;
              break;
            }
          }
          
          if (!consultora) continue;
          
          const appRaw = String(row[idxRetencao.app] || "").trim();
          const app = normalizarAppStatus(appRaw);
          
          if (!retencaoData[consultora]) {
            retencaoData[consultora] = { 
              nome: consultora, 
              total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0,
              setor: "RETENÇÃO",
              origem: "retencao"
            };
          }
          
          const r = retencaoData[consultora];
          r.total++;
          
          if (app === "SIM") { 
            r.sim++;
          }
          else if (app === "NAO") { 
            r.nao++;
          }
          else if (app === "CANCELADO") { 
            r.cancelado++;
          }
          else {
            r.outros++;
          }
        }
      }
    }

    const responseData = {
      mes: MESES[mesAtual - 1],
      ano: anoAtual,
      geral: totalGeral,
      appLoja: appLoja,
      appWeb: appWeb,
      consultores: Object.values(consultoresData),
      consultorasRetencao: Object.values(retencaoData)
    };

    criarDashboardApp(responseData);
    return createSuccessResponse(responseData);
  } catch (error) {
    return createErrorResponse(error.message);
  }
}

function criarDashboardApp(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard App");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard App");
  sheet.clear();

  sheet.getRange("A1:L1").merge().setValue("📱 DASHBOARD APP - " + dados.mes.toUpperCase() + " " + dados.ano)
    .setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  renderizarTresCardsPrincipais(sheet, dados, "APP");

  if (dados.consultorasRetencao && dados.consultorasRetencao.length > 0) {
    let currentRow = 13;
    
    sheet.getRange(currentRow, 1, 1, 12).merge()
      .setValue("👑 CONSULTORAS DE RETENÇÃO (DADOS SEPARADOS)")
      .setBackground("#7c3aed")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    currentRow += 2;
    
    let col = 1;
    dados.consultorasRetencao.sort((a, b) => a.nome.localeCompare(b.nome)).forEach(c => {
      const p = c.total > 0 ? Math.round((c.sim / c.total) * 100) : 0;
      const card = [
        [c.nome + " (RETENÇÃO)", ""],
        ["TOTAL", c.total],
        ["COM APP (SIM)", c.sim],
        ["SEM APP (NÃO)", c.nao],
        ["CANCELADO", c.cancelado],
        ["OUTROS", c.outros || 0],
        ["% COM APP", p + "%"],
        ["ORIGEM", c.origem === "retencao" ? "APP RETENÇÃO" : "APP"]
      ];
      
      sheet.getRange(currentRow, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
      
      sheet.getRange(currentRow, col, 1, 2).merge()
        .setBackground("#7c3aed")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(currentRow + 6, col + 1)
        .setBackground(cCor)
        .setFontColor("#ffffff")
        .setFontWeight("bold");
      
      col += 3;
      if (col > 9) { 
        col = 1; 
        currentRow += card.length + 1; 
      }
    });
    
    if (col !== 1) currentRow += 9; else currentRow += 2;
    
    sheet.getRange(currentRow, 1, 1, 12).setBackground("#e5e7eb");
    currentRow += 2;
  }

  const setores = ["VENDAS", "RECEPCAO", "REFILIACAO", "WEB SITE", "TELEVENDAS", "OUTROS"];
  
  let currentRow = sheet.getLastRow() + 2;

  setores.forEach(setor => {
    const consultoresSetor = dados.consultores.filter(c => 
      c.setor === setor && c.origem !== "retencao"
    );
    
    if (consultoresSetor.length === 0) return;

    sheet.getRange(currentRow, 1, 1, 12).merge().setValue("SETOR: " + setor)
      .setBackground("#4b5563").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("left");
    currentRow += 2;

    let col = 1;
    consultoresSetor.sort((a, b) => a.nome.localeCompare(b.nome)).forEach(c => {
      const p = c.total > 0 ? Math.round((c.sim / c.total) * 100) : 0;
      const card = [
        [c.nome, ""],
        ["TOTAL", c.total],
        ["SIM", c.sim],
        ["NÃO", c.nao],
        ["CANCELADO", c.cancelado],
        ["OUTROS", c.outros || 0],
        ["% APP", p + "%"]
      ];
      
      sheet.getRange(currentRow, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
      sheet.getRange(currentRow, col, 1, 2).merge().setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      
      const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(currentRow + 6, col + 1).setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");

      col += 3;
      if (col > 9) { col = 1; currentRow += 8; }
    });
    
    if (col !== 1) currentRow += 8; else currentRow += 2;
  });
  
  for (let i = 1; i <= 12; i++) sheet.setColumnWidth(i, 160);
}

// ============================================================================
// ENDPOINT: ADIMPLÊNCIA
// ============================================================================

function getAdimplenciaData(mes, ano) {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    const sheet = ssDados.getSheetByName("ADIMPLENCIA") || ssDados.getSheetByName("adimplência") || ssDados.getSheetByName("Adimplência");
    if (!sheet) return createErrorResponse("Sheet 'ADIMPLENCIA' não encontrada");
    const data = sheet.getDataRange().getValues();
    const headers = data[1];
    const idx = { consultor: headers.indexOf("Consultor"), data: headers.indexOf("Data"), kyc: headers.indexOf("KYC"), mensalidadeOk: headers.indexOf("MENSALIDADE OK"), statusBi: headers.indexOf("STATUS BI") };
    const mesAtual = mes || new Date().getMonth() + 1;
    const anoAtual = ano || new Date().getFullYear();
    const consultoresData = {};
    let totalGeral = { totalTrocas: 0, mensOk: 0, mensAberto: 0, mensAtraso: 0, aprovados: 0, pendentes: 0, totalBi: 0, foraBi: 0, okBi: 0 };
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const dataParsed = parseDate(row[idx.data]);
      if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;
      
      const consultor = normalizarConsultorOriginal(String(row[idx.consultor] || "").trim());
      const kyc = String(row[idx.kyc] || "").trim().toUpperCase();
      const mensalidade = String(row[idx.mensalidadeOk] || "").trim().toUpperCase();
      const statusBi = String(row[idx.statusBi] || "").trim().toUpperCase();
      
      if (!consultoresData[consultor]) consultoresData[consultor] = { nome: consultor, totalTrocas: 0, mensOk: 0, mensAberto: 0, mensAtraso: 0, aprovados: 0, pendentes: 0, totalBi: 0, foraBi: 0, okBi: 0 };
      const c = consultoresData[consultor];
      c.totalTrocas++; totalGeral.totalTrocas++;
      
      if (mensalidade === "OK") { c.mensOk++; totalGeral.mensOk++; }
      if (mensalidade === "EM ABERTO") { c.mensAberto++; totalGeral.mensAberto++; }
      if (mensalidade === "ATRASO") { c.mensAtraso++; totalGeral.mensAtraso++; }
      
      if (kyc === "APROVADO") { 
        c.aprovados++; totalGeral.aprovados++; 
      } else {
        c.pendentes++; totalGeral.pendentes++;
      }
      
      if (statusBi === "OK") { c.totalBi++; totalGeral.totalBi++; }
      if (statusBi === "FORA") { c.foraBi++; totalGeral.foraBi++; }
      if (statusBi === "OK" && mensalidade === "OK") { c.okBi++; totalGeral.okBi++; }
    }
    for (const c of Object.values(consultoresData)) c.percentualAprovado = c.totalTrocas > 0 ? Math.round((c.aprovados / c.totalTrocas) * 100) : 0;
    totalGeral.percentualAprovado = totalGeral.totalTrocas > 0 ? Math.round((totalGeral.aprovados / totalGeral.totalTrocas) * 100) : 0;
    const responseData = { mes: MESES[mesAtual - 1], ano: anoAtual, geral: totalGeral, consultores: Object.values(consultoresData).sort((a, b) => a.nome.localeCompare(b.nome)) };
    criarDashboardAdimplencia(responseData);
    return createSuccessResponse(responseData);
  } catch (error) { return createErrorResponse(error.message); }
}

function criarDashboardAdimplencia(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Adimplência");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Adimplência");
  sheet.clear();
  sheet.getRange("A1:L1").merge().setValue("DASHBOARD ADIMPLÊNCIA - " + dados.mes.toUpperCase() + " " + dados.ano).setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);
  const g = dados.geral;
  const cardGeral = [["EQUIPE - TOTAL", ""], ["TOTAL TROCAS", g.totalTrocas], ["MENS. OK", g.mensOk], ["MENS. ABERTO", g.mensAberto], ["MENS. ATRASO", g.mensAtraso], ["APROVADOS", g.aprovados], ["PENDENTES", g.pendentes], ["% APROVADOS", g.percentualAprovado + "%"], ["TOTAL BI", g.totalBi], ["FORA BI", g.foraBi], ["OK BI", g.okBi]];
  sheet.getRange(3, 1, cardGeral.length, 2).setValues(cardGeral).setBorder(true, true, true, true, true, true);
  sheet.getRange(3, 1, 1, 2).merge().setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(10, 2).setBackground(g.percentualAprovado >= 90 ? "#10b981" : g.percentualAprovado >= 80 ? "#f59e0b" : "#ef4444").setFontColor("#ffffff").setFontWeight("bold");
  let row = 16; let col = 1;
  dados.consultores.forEach((c, i) => {
    const card = [[c.nome, ""], ["TOTAL TROCAS", c.totalTrocas], ["MENS. OK", c.mensOk], ["MENS. ABERTO", c.mensAberto], ["MENS. ATRASO", c.mensAtraso], ["APROVADOS", c.aprovados], ["PENDENTES", c.pendentes], ["% APROVADOS", c.percentualAprovado + "%"], ["TOTAL BI", c.totalBi], ["FORA BI", c.foraBi], ["OK BI", c.okBi]];
    sheet.getRange(row, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
    sheet.getRange(row, col, 1, 2).merge().setBackground("#1e40af").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(row + 7, col + 1).setBackground(c.percentualAprovado >= 90 ? "#10b981" : c.percentualAprovado >= 80 ? "#f59e0b" : "#ef4444").setFontColor("#ffffff").setFontWeight("bold");
    col += 3; if (col > 9) { col = 1; row += 13; }
  });
  for (let i = 1; i <= 12; i++) sheet.setColumnWidth(i, 140);
}

// ============================================================================
// ENDPOINT: RECORRÊNCIA (SISTEMA POR MÊS FIXO)
// ============================================================================

function processarDadosRecorrenciaPorMes(mes, ano, isFixedMonth = false) {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    let sheetRec = ssDados.getSheetByName(NOME_ABA_RECORRENCIA_FONTE);
    
    if (!sheetRec) {
      const sheets = ssDados.getSheets();
      sheetRec = sheets.find(s => s.getName().toUpperCase().includes("RECORRENCIA"));
    }

    if (!sheetRec) throw new Error("Aba de dados não encontrada.");

    function formatToMonthYear(val) {
      if (!val) return "";
      let d;
      if (val instanceof Date) {
        d = val;
      } else {
        d = new Date(val);
      }
      if (isNaN(d.getTime())) return String(val).trim();
      
      const m = d.getMonth() + 1;
      const y = d.getFullYear();
      return (m < 10 ? '0' + m : m) + '/' + y;
    }

    const targetMonth = (mes < 10 ? '0' + mes : mes) + '/' + ano;
    
    // CALCULAR MESES ANTERIORES BASEADO NO MÊS ESPECIFICADO
    let historicoMonths = [];
    
    if (isFixedMonth) {
      // Para dados fixos do mês: usar 3 meses anteriores ao mês especificado
      for (let i = 1; i <= 3; i++) {
        let histMes = mes - i;
        let histAno = ano;
        if (histMes <= 0) {
          histMes += 12;
          histAno -= 1;
        }
        historicoMonths.push((histMes < 10 ? '0' + histMes : histMes) + '/' + histAno);
      }
    } else {
      // Para compatibilidade: usar meses anteriores ao mês atual
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      
      for (let i = 1; i <= 3; i++) {
        let histMes = currentMonth - i;
        let histAno = currentYear;
        if (histMes <= 0) {
          histMes += 12;
          histAno -= 1;
        }
        historicoMonths.push((histMes < 10 ? '0' + histMes : histMes) + '/' + histAno);
      }
    }

    const dataRec = sheetRec.getDataRange().getValues();
    const rows = dataRec.slice(2);

    function calculateMetrics(consultor, targetMonths, isCurrentMonth = false) {
      const filteredRows = rows.filter(row => {
        const rowConsultor = String(row[0]).trim().toUpperCase();
        const rowDateStr = formatToMonthYear(row[4]);
        return rowConsultor === consultor.toUpperCase() && targetMonths.includes(rowDateStr);
      });

      let totalRetido = filteredRows.length;
      let ok = 0, emAberto = 0, emAtraso = 0, cancelado = 0, pendenciasKYC = 0, pendenciasMensalidade = 0;

      filteredRows.forEach(row => {
        const mensalidade = String(row[2]).trim().toUpperCase();
        const kyc = String(row[3]).trim().toUpperCase();
        
        if (mensalidade === 'OK') ok++;
        else if (mensalidade === 'EM ABERTO') emAberto++;
        else if (mensalidade === 'EM ATRASO') emAtraso++;
        else if (mensalidade === 'CANCELADO') cancelado++;
        
        if (mensalidade !== 'OK' && mensalidade !== 'EM ABERTO' && mensalidade !== 'CANCELADO') {
          pendenciasMensalidade++;
        }
        
        if (kyc !== 'APROVADO' && mensalidade !== 'CANCELADO') {
          pendenciasKYC++;
        }
      });

      if (isCurrentMonth) {
        const totalPendencias = pendenciasKYC;
        const totalRetidosFinal = totalRetido - cancelado;
        const retençõesOK = totalRetidosFinal - totalPendencias;
        const percentualOK = totalRetidosFinal > 0 ? Math.round((retençõesOK / totalRetidosFinal) * 100) : 0;
        
        return { 
          totalRetido, 
          ok, 
          emAberto, 
          emAtraso, 
          cancelado, 
          pendenciasKYC, 
          totalPendencias, 
          totalRetidosFinal, 
          retençõesOK, 
          percentualOK 
        };
      } else {
        const totalOK = ok + emAberto;
        const totalPendencias = pendenciasMensalidade;
        const totalRetidosFinal = totalRetido - cancelado;
        const percentualOK = totalRetidosFinal > 0 ? Math.round((totalOK / totalRetidosFinal) * 100) : 0;
        
        return { 
          totalRetido, 
          ok,
          emAberto,
          totalOK,
          emAtraso, 
          pendencias: totalPendencias, 
          cancelado,
          totalRetidosFinal,
          percentualOK 
        };
      }
    }

    const consultoresRetencao = CONSULTORES_RECORRENCIA_RETENCAO;
    const consultoresRefiliacao = CONSULTORES_RECORRENCIA_REFILIACAO;
    
    const result = { 
      retencao: {}, 
      refiliacao: {}, 
      periodo: { 
        atual: targetMonth, 
        historico: historicoMonths,
        mes: mes,
        ano: ano,
        nomeMes: MESES[mes-1],
        isFixedMonth: isFixedMonth
      } 
    };

    consultoresRetencao.forEach(c => {
      result.retencao[c] = {
        atual: calculateMetrics(c, [targetMonth], true),
        historico: historicoMonths.map(m => ({ 
          mes: m, 
          dados: calculateMetrics(c, [m]) 
        })),
        total3Meses: calculateMetrics(c, historicoMonths)
      };
    });

    consultoresRefiliacao.forEach(c => {
      result.refiliacao[c] = {
        historico: historicoMonths.map(m => ({ 
          mes: m, 
          dados: calculateMetrics(c, [m]) 
        })),
        total3Meses: calculateMetrics(c, historicoMonths)
      };
    });

    return { status: "success", data: result };
  } catch (error) {
    return { status: "error", error: error.message };
  }
}

function processarDadosRecorrencia() {
  const now = new Date();
  const mes = now.getMonth() + 1;
  const ano = now.getFullYear();
  return processarDadosRecorrenciaPorMes(mes, ano, false);
}

function gerarDashboardRecorrencia() {
  const now = new Date();
  const mes = now.getMonth() + 1;
  const ano = now.getFullYear();
  
  const dados = processarDadosRecorrenciaPorMes(mes, ano, false);
  if (dados.status === "success") {
    renderizarPlanilhaRecorrenciaPorMes(dados.data, mes, ano);
    SpreadsheetApp.getUi().alert(`✅ Dashboard de Recorrência de ${MESES[mes-1]} ${ano} atualizado com sucesso!`);
  } else {
    SpreadsheetApp.getUi().alert("❌ Erro: " + dados.error);
  }
}

function gerarDashboardRecorrenciaHistorico() {
  const dados = processarDadosRecorrencia();
  if (dados.status === "success") {
    renderizarPlanilhaRecorrenciaDetalhada(dados.data);
    SpreadsheetApp.getUi().alert("✅ Dashboard Histórico de Recorrência atualizado!");
  } else {
    SpreadsheetApp.getUi().alert("❌ Erro: " + dados.error);
  }
}

function renderizarPlanilhaRecorrenciaPorMes(dados, mes, ano) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  const nomeAba = `Recorrência ${MESES[mes-1]} ${ano}`;
  
  let sheet = ssDash.getSheetByName(nomeAba);
  if (!sheet) sheet = ssDash.insertSheet(nomeAba);
  
  // Remove a aba antiga se existir com nome diferente
  const oldSheet = ssDash.getSheetByName("Dashboard Recorrência");
  if (oldSheet && oldSheet.getName() !== nomeAba) {
    ssDash.deleteSheet(oldSheet);
  }
  
  sheet.clear();

  sheet.getRange("A1:M1").merge().setValue(`📈 DASHBOARD RECORRÊNCIA - ${MESES[mes-1]} ${ano}`)
    .setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  if (dados.periodo.isFixedMonth) {
    sheet.getRange("A2:M2").merge().setValue(`📅 DADOS FIXOS DO MÊS ${dados.periodo.atual} (não atualiza automaticamente)`)
      .setBackground("#f59e0b").setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center");
  } else {
    sheet.getRange("A2:M2").merge().setValue(`🔄 DADOS DO MÊS ATUAL (atualiza automaticamente)`)
      .setBackground("#10b981").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  }

  let currentRow = 4;
  
  Object.keys(dados.retencao).forEach(c => {
    const d = dados.retencao[c];
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("CONSULTOR: " + c + " (RETENÇÃO)")
      .setBackground("#f1f5f9").setFontWeight("bold");
    
    currentRow++;
    
    const cardAtual = [
      ["MÊS ATUAL (" + dados.periodo.atual + ")", ""],
      ["TOTAL RETIDO", d.atual.totalRetido],
      ["CANCELADO", d.atual.cancelado],
      ["TOTAL RETIDOS FINAL", d.atual.totalRetidosFinal],
      ["MENSALIDADES OK", d.atual.ok],
      ["EM ABERTO", d.atual.emAberto],
      ["EM ATRASO", d.atual.emAtraso],
      ["PENDÊNCIAS KYC", d.atual.pendenciasKYC],
      ["TOTAL PENDÊNCIAS", d.atual.totalPendencias],
      ["RETENÇÕES OK", d.atual.retençõesOK],
      ["% OK", d.atual.percentualOK + "%"]
    ];
    sheet.getRange(currentRow, 1, cardAtual.length, 2).setValues(cardAtual).setBorder(true, true, true, true, true, true);
    sheet.getRange(currentRow, 1, 1, 2).merge().setBackground("#10b981").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    
    const cardTotal = [
      ["TOTAL 3 MESES ANTERIORES", ""],
      ["TOTAL RETIDO", d.total3Meses.totalRetido],
      ["CANCELADO", d.total3Meses.cancelado || 0],
      ["TOTAL RETIDOS FINAL", d.total3Meses.totalRetidosFinal],
      ["OK (Mensalidade OK)", d.total3Meses.ok],
      ["EM ABERTO", d.total3Meses.emAberto],
      ["TOTAL OK (OK + Aberto)", d.total3Meses.totalOK || (d.total3Meses.ok + d.total3Meses.emAberto)],
      ["EM ATRASO", d.total3Meses.emAtraso],
      ["TOTAL PENDÊNCIAS", d.total3Meses.pendencias],
      ["% OK (considera Aberto)", (d.total3Meses.percentualOK || 0) + "%"]
    ];
    sheet.getRange(currentRow, 4, cardTotal.length, 2).setValues(cardTotal).setBorder(true, true, true, true, true, true);
    sheet.getRange(currentRow, 4, 1, 2).merge().setBackground("#f59e0b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    
    const percentualCell = sheet.getRange(currentRow + cardTotal.length - 1, 5);
    const corPercentual = d.total3Meses.percentualOK >= 90 ? "#10b981" : d.total3Meses.percentualOK >= 80 ? "#f59e0b" : "#ef4444";
    percentualCell.setBackground(corPercentual).setFontColor("#ffffff").setFontWeight("bold");
    
    currentRow += Math.max(cardAtual.length, cardTotal.length) + 2;
  });

  Object.keys(dados.refiliacao).forEach(c => {
    const d = dados.refiliacao[c];
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("CONSULTOR: " + c + " (REFILIAÇÃO)")
      .setBackground("#f1f5f9").setFontWeight("bold");
    
    currentRow++;
    
    const cardTotal = [
      ["TOTAL 3 MESES ANTERIORES", ""],
      ["TOTAL REFILIAÇÃO", d.total3Meses.totalRetido],
      ["CANCELADO", d.total3Meses.cancelado || 0],
      ["TOTAL REFILIADOS FINAL", d.total3Meses.totalRetidosFinal],
      ["OK (Mensalidade OK)", d.total3Meses.ok],
      ["EM ABERTO", d.total3Meses.emAberto],
      ["TOTAL OK (OK + Aberto)", d.total3Meses.totalOK || (d.total3Meses.ok + d.total3Meses.emAberto)],
      ["EM ATRASO", d.total3Meses.emAtraso],
      ["TOTAL PENDÊNCIAS", d.total3Meses.pendencias],
      ["% OK (considera Aberto)", (d.total3Meses.percentualOK || 0) + "%"]
    ];
    sheet.getRange(currentRow, 1, cardTotal.length, 2).setValues(cardTotal).setBorder(true, true, true, true, true, true);
    sheet.getRange(currentRow, 1, 1, 2).merge().setBackground("#f59e0b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    
    const percentualCell = sheet.getRange(currentRow + cardTotal.length - 1, 2);
    const corPercentual = d.total3Meses.percentualOK >= 90 ? "#10b981" : d.total3Meses.percentualOK >= 80 ? "#f59e0b" : "#ef4444";
    percentualCell.setBackground(corPercentual).setFontColor("#ffffff").setFontWeight("bold");
    
    currentRow += cardTotal.length + 2;
  });
  
  for (let i = 1; i <= 13; i++) sheet.setColumnWidth(i, 150);
  
  sheet.setFrozenRows(2);
}

function renderizarPlanilhaRecorrenciaDetalhada(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Recorrência Histórico");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Recorrência Histórico");
  sheet.clear();

  sheet.getRange("A1:M1").merge().setValue("📊 DASHBOARD RECORRÊNCIA - HISTÓRICO DETALHADO")
    .setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  let currentRow = 3;
  
  Object.keys(dados.retencao).forEach(c => {
    const d = dados.retencao[c];
    
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("👑 " + c + " - RETENÇÃO")
      .setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
    currentRow++;
    
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("MÊS ATUAL: " + dados.periodo.atual)
      .setBackground("#10b981").setFontColor("#ffffff").setFontWeight("bold");
    currentRow++;
    
    const headersAtual = ["Métrica", "Valor"];
    sheet.getRange(currentRow, 1, 1, 2).setValues([headersAtual]).setBackground("#d1fae5");
    currentRow++;
    
    const dadosAtual = [
      ["Total Retido", d.atual.totalRetido],
      ["Cancelados", d.atual.cancelado],
      ["Total Retidos Final", d.atual.totalRetidosFinal],
      ["Mensalidades OK", d.atual.ok],
      ["Em Aberto", d.atual.emAberto],
      ["Em Atraso", d.atual.emAtraso],
      ["Pendências KYC", d.atual.pendenciasKYC],
      ["Total Pendências", d.atual.totalPendencias],
      ["Retenções OK", d.atual.retençõesOK],
      ["% OK", d.atual.percentualOK + "%"]
    ];
    sheet.getRange(currentRow, 1, dadosAtual.length, 2).setValues(dadosAtual);
    currentRow += dadosAtual.length + 2;
    
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("HISTÓRICO - ÚLTIMOS 3 MESES")
      .setBackground("#f59e0b").setFontColor("#ffffff").setFontWeight("bold");
    currentRow++;
    
    const headersHistorico = ["Mês", "Total", "OK", "Em Aberto", "Total OK", "Pendências", "% OK"];
    sheet.getRange(currentRow, 1, 1, 7).setValues([headersHistorico]).setBackground("#fef3c7");
    currentRow++;
    
    d.historico.forEach(h => {
      sheet.getRange(currentRow, 1, 1, 7).setValues([
        [
          h.mes, 
          h.dados.totalRetido, 
          h.dados.ok, 
          h.dados.emAberto, 
          h.dados.totalOK || (h.dados.ok + h.dados.emAberto),
          h.dados.pendencias, 
          h.dados.percentualOK + "%"
        ]
      ]);
      currentRow++;
    });
    
    currentRow += 2;
  });
  
  for (let i = 1; i <= 13; i++) sheet.setColumnWidth(i, 140);
}

// ============================================================================
// ENDPOINT: RECORRÊNCIA VENDEDOR (SEM FILTRO DE MÊS)
// ============================================================================

function getRecorrenciaVendedorData() { // REMOVIDOS os parâmetros mes, ano
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    let sheet = ssDados.getSheetByName(NOME_ABA_RECORRENCIA_VENDAS);
    
    if (!sheet) {
      const sheets = ssDados.getSheets();
      sheet = sheets.find(s => {
        const nome = s.getName().toUpperCase();
        return nome.includes("RECORRENCIA VENDAS") || 
               nome.includes("RECORRENCIA_VENDAS") ||
               nome.includes("RECORRENCIA-VENDAS");
      });
    }
    
    if (!sheet) {
      return createErrorResponse("Sheet 'RECORRENCIA VENDAS' não encontrada");
    }

    const data = sheet.getDataRange().getValues();
    
    // Verificar se tem cabeçalhos (assumindo que a primeira linha é cabeçalho)
    if (data.length < 2) {
      return createErrorResponse("Planilha 'RECORRENCIA VENDAS' está vazia");
    }
    
    // Baseado na sua descrição: A=matricula, B=filiado, C=consultor, D=data, E=status
    // Vamos usar índices fixos 0-4 conforme sua descrição
    const idx = {
      matricula: 0,    // Coluna A
      filiado: 1,      // Coluna B
      consultor: 2,    // Coluna C
      data: 3,         // Coluna D (data da filiação) - APENAS PARA REFERÊNCIA, NÃO FILTRA
      status: 4        // Coluna E
    };
    
    // Inicializar estrutura de dados
    const consultoresData = {};
    let totalGeral = {
      totalVendasPromocao: 0,
      totalOk: 0,
      totalEmAberto: 0,
      totalAtraso: 0,
      totalOutros: 0
    };
    
    // Data atual para o título do dashboard
    const dataAtual = new Date();
    const mesAtualNumero = dataAtual.getMonth() + 1;
    const anoAtual = dataAtual.getFullYear();
    
    // Processar TODAS as linhas de dados (SEM FILTRO DE MÊS)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Verificar se a linha tem dados básicos
      if (!row[idx.consultor] && !row[idx.matricula]) continue;
      
      const consultorOriginal = String(row[idx.consultor] || "").trim();
      if (!consultorOriginal) continue;
      
      // Normalizar nome do consultor (usar apenas primeiro nome)
      const consultorNome = normalizarNomeConsultorRecorrencia(consultorOriginal);
      
      // Determinar setor
      let setor = determinarSetorConsultor(consultorNome);
      
      // Verificar se o consultor está no mapeamento
      if (!consultorEstaNoMapeamento(consultorNome)) {
        // Se não estiver no mapeamento, não incluir
        continue;
      }
      
      // Inicializar dados do consultor se não existir
      if (!consultoresData[consultorNome]) {
        consultoresData[consultorNome] = {
          nome: consultorNome,
          setor: setor,
          totalVendasPromocao: 0,
          totalOk: 0,
          totalEmAberto: 0,
          totalAtraso: 0,
          totalOutros: 0
        };
      }
      
      const c = consultoresData[consultorNome];
      c.totalVendasPromocao++;
      totalGeral.totalVendasPromocao++;
      
      // Analisar status (coluna E)
      const status = String(row[idx.status] || "").trim().toUpperCase();
      
      if (status === "OK") {
        c.totalOk++;
        totalGeral.totalOk++;
      } else if (status.includes("ABERTO") || status === "EM ABERTO") {
        c.totalEmAberto++;
        totalGeral.totalEmAberto++;
      } else if (status.includes("ATRASO") || status === "EM ATRASO") {
        c.totalAtraso++;
        totalGeral.totalAtraso++;
      } else {
        c.totalOutros++;
        totalGeral.totalOutros++;
      }
    }
    
    // Calcular percentuais para cada consultor
    Object.values(consultoresData).forEach(c => {
      c.percentualVendasOk = c.totalVendasPromocao > 0 ? 
        Math.round(((c.totalOk + c.totalEmAberto) / c.totalVendasPromocao) * 100) : 0;
    });
    
    // Calcular percentual geral
    totalGeral.percentualVendasOk = totalGeral.totalVendasPromocao > 0 ?
      Math.round(((totalGeral.totalOk + totalGeral.totalEmAberto) / totalGeral.totalVendasPromocao) * 100) : 0;
    
    // Organizar dados por setor
    const dadosPorSetor = {};
    
    for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
      dadosPorSetor[setor] = Object.values(consultoresData)
        .filter(c => c.setor === setor)
        .sort((a, b) => b.totalVendasPromocao - a.totalVendasPromocao);
    }
    
    const responseData = {
      mes: "TODOS OS MESES", // MODIFICADO: Agora mostra "TODOS OS MESES"
      ano: "TODOS OS ANOS",  // MODIFICADO: Agora mostra "TODOS OS ANOS"
      mesAtual: MESES[mesAtualNumero - 1], // Mantém referência do mês atual para título
      anoAtual: anoAtual, // Mantém referência do ano atual para título
      geral: totalGeral,
      consultores: Object.values(consultoresData)
        .sort((a, b) => b.totalVendasPromocao - a.totalVendasPromocao),
      dadosPorSetor: dadosPorSetor,
      totalConsultores: Object.keys(consultoresData).length,
      totalRegistros: data.length - 1 // Total de linhas processadas (exceto cabeçalho)
    };
    
    criarDashboardRecorrenciaVendedor(responseData);
    return createSuccessResponse(responseData);
    
  } catch (error) {
    return createErrorResponse(`Erro no processamento: ${error.message}`);
  }
}

function normalizarNomeConsultorRecorrencia(nome) {
  if (!nome) return "DESCONHECIDO";
  
  const nomeUpper = nome.toUpperCase().trim();
  
  // 1. Correspondência exata (case-insensitive) — retorna o valor do mapa
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    for (const consultor of CONSULTORES_RECORRENCIA_VENDAS[setor]) {
      const consultorUpper = consultor.toUpperCase();
      if (nomeUpper === consultorUpper || nomeUpper.includes(consultorUpper)) {
        return consultor; // retorna o valor ORIGINAL do mapa (para comparação posterior)
      }
    }
  }
  
  // 2. Correspondência pelo primeiro nome
  const primeiroNome = nomeUpper.split(" ")[0];
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    for (const consultor of CONSULTORES_RECORRENCIA_VENDAS[setor]) {
      const primeiroNomeConsultor = consultor.toUpperCase().split(" ")[0];
      if (primeiroNome === primeiroNomeConsultor) {
        return consultor; // retorna o valor ORIGINAL do mapa
      }
    }
  }
  
  return nomeUpper; // não encontrado
}

function determinarSetorConsultor(consultorNome) {
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    if (CONSULTORES_RECORRENCIA_VENDAS[setor].includes(consultorNome)) {
      return setor;
    }
  }
  return "OUTROS";
}

function consultorEstaNoMapeamento(consultorNome) {
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    if (CONSULTORES_RECORRENCIA_VENDAS[setor].includes(consultorNome)) {
      return true;
    }
  }
  return false;
}

function criarDashboardRecorrenciaVendedor(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Recorrência Vendedor");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Recorrência Vendedor");
  sheet.clear();
  
  console.log("=== INICIANDO DASHBOARD ===");
  console.log("Total consultores:", dados.consultores ? dados.consultores.length : 0);
  
  // 1. TÍTULO
  sheet.getRange("A1:L1").merge()
    .setValue(`🤝 DASHBOARD RECORRÊNCIA VENDEDOR - DADOS COMPLETOS`)
    .setBackground("#0d9488")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(14);
  
  // 2. CARD GERAL
  const g = dados.geral || { totalVendasPromocao: 0, totalOk: 0, totalEmAberto: 0, totalAtraso: 0, totalOutros: 0 };
  const percentualGeral = g.totalVendasPromocao > 0 ? 
    Math.round(((g.totalOk + g.totalEmAberto) / g.totalVendasPromocao) * 100) : 0;
  
  const cardGeral = [
    ["📈 TOTAL GERAL", ""],
    ["TOTAL VENDAS", g.totalVendasPromocao],
    ["OK", g.totalOk],
    ["EM ABERTO", g.totalEmAberto],
    ["ATRASO", g.totalAtraso],
    ["OUTROS", g.totalOutros],
    ["", ""],
    ["% OK+ABERTO", percentualGeral + "%"]
  ];
  
  sheet.getRange(3, 1, cardGeral.length, 2)
    .setValues(cardGeral)
    .setBorder(true, true, true, true, true, true);
  
  sheet.getRange(3, 1, 1, 2).merge()
    .setBackground("#0f766e")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  // Colorir percentual
  const corPercentual = percentualGeral >= 90 ? "#10b981" : 
                       percentualGeral >= 80 ? "#f59e0b" : "#ef4444";
  sheet.getRange(10, 2)
    .setBackground(corPercentual)
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  
  let currentRow = 14;
  
  // 3. DEFINIR OS 3 SETORES ESPERADOS
  const setores = [
    { nome: "VENDAS", cor: "#1e3a8a" },
    { nome: "REFILIACAO", cor: "#7c3aed" },
    { nome: "RECEPCAO", cor: "#059669" }
  ];
  
  // 4. CRIAR SETORES MANUALMENTE (sem depender de dados.dadosPorSetor)
  setores.forEach(setorInfo => {
    const setor = setorInfo.nome;
    const corSetor = setorInfo.cor;
    
    // Filtrar consultores deste setor
    const consultoresSetor = dados.consultores ? 
      dados.consultores.filter(c => c.setor === setor) : [];
    
    console.log(`${setor}: ${consultoresSetor.length} consultores`);
    
    // Título do setor
    sheet.getRange(currentRow, 1, 1, 12).merge()
      .setValue(`👥 SETOR: ${setor} (${consultoresSetor.length} consultor${consultoresSetor.length !== 1 ? 'es' : ''})`)
      .setBackground(corSetor)
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    currentRow++;
    
    if (consultoresSetor.length === 0) {
      sheet.getRange(currentRow, 1, 1, 12).merge()
        .setValue("ℹ️ Nenhum dado encontrado")
        .setBackground("#f0f0f0")
        .setFontColor("#666666")
        .setHorizontalAlignment("center");
      
      currentRow += 2;
      return;
    }
    
    // Ordenar por vendas (maior primeiro)
    consultoresSetor.sort((a, b) => b.totalVendasPromocao - a.totalVendasPromocao);
    
    let col = 1;
    let consultorRow = currentRow;
    const CARD_HEIGHT = 9; // 8 linhas de dados + 1 linha de espaçamento
    
    // CRIAR CARDS DOS CONSULTORES
    consultoresSetor.forEach(consultor => {
      const percentualConsultor = consultor.totalVendasPromocao > 0 ? 
        Math.round(((consultor.totalOk + consultor.totalEmAberto) / consultor.totalVendasPromocao) * 100) : 0;
      
      const cardConsultor = [
        [consultor.nome, ""],
        ["TOTAL VENDAS", consultor.totalVendasPromocao],
        ["OK", consultor.totalOk],
        ["EM ABERTO", consultor.totalEmAberto],
        ["ATRASO", consultor.totalAtraso],
        ["OUTROS", consultor.totalOutros || 0],
        ["", ""],
        ["% OK", percentualConsultor + "%"]
      ];
      
      // Inserir card
      sheet.getRange(consultorRow, col, cardConsultor.length, 2)
        .setValues(cardConsultor)
        .setBorder(true, true, true, true, true, true);
      
      sheet.getRange(consultorRow, col, 1, 2).merge()
        .setBackground(corSetor)
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      const corPercentualConsultor = percentualConsultor >= 90 ? "#10b981" : 
                                    percentualConsultor >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(consultorRow + 7, col + 1)
        .setBackground(corPercentualConsultor)
        .setFontColor("#ffffff")
        .setFontWeight("bold");
      
      // Avançar coluna — máx 4 cards por linha (colunas 1, 4, 7, 10)
      col += 3;
      if (col > 10) {
        col = 1;
        consultorRow += CARD_HEIGHT;
      }
    });
    
    // CORREÇÃO: sempre avançar currentRow pela última linha usada
    // Se col === 1, já avançou para a próxima linha no loop; caso contrário, precisa avançar agora
    currentRow = (col === 1) ? consultorRow : consultorRow + CARD_HEIGHT;
    currentRow += 2; // espaço entre setores
  });
  
  // 5. RESUMO FINAL
  sheet.getRange(currentRow, 1, 1, 12).merge()
    .setValue("📊 RESUMO FINAL")
    .setBackground("#4b5563")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  currentRow++;
  
  const resumo = [
    ["Total Consultores:", dados.totalConsultores || 0],
    ["Total Vendas:", g.totalVendasPromocao],
    ["Registros Processados:", dados.totalRegistros || 0],
    ["Média por Consultor:", dados.totalConsultores > 0 ? 
      Math.round(g.totalVendasPromocao / dados.totalConsultores) : 0]
  ];
  
  sheet.getRange(currentRow, 1, resumo.length, 2)
    .setValues(resumo)
    .setBorder(true, true, true, true, true, true);
  
  sheet.getRange(currentRow, 1, 1, 2).merge()
    .setBackground("#0d9488")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  
  // 6. AJUSTES FINAIS
  for (let i = 1; i <= 12; i++) {
    sheet.setColumnWidth(i, 160);
  }
  
  sheet.setFrozenRows(2);
  
  console.log("=== DASHBOARD FINALIZADO ===");
}

function corPorSetorRecorrenciaVendedor(setor) {
  const cores = {
    "VENDAS": "#1e3a8a",      // Azul escuro
    "REFILIACAO": "#7c3aed",  // Roxo
    "RECEPCAO": "#059669",    // Verde
    "OUTROS": "#4b5563"       // Cinza
  };
  return cores[setor] || "#4b5563";
}

// ============================================================================
// ENDPOINT: REFUTURIZA (CORRIGIDO - SEM FILTRO EM CÉLULAS MESCLADAS)
// ============================================================================

function getRefuturizaData(mes, ano) {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    
    let sheet = ssDados.getSheetByName("REFUTURIZA");
    
    if (!sheet) {
      const sheets = ssDados.getSheets();
      sheet = sheets.find(s => {
        const nome = s.getName().toUpperCase();
        return nome.includes("REFUTURIZA") || nome.includes("REFUTURISA") || nome.includes("REFUTUR");
      });
    }
    
    if (!sheet) return createErrorResponse("Sheet 'REFUTURIZA' não encontrada");

    const data = sheet.getDataRange().getValues();
    
    const idx = {
      matricula: 0,
      nomeAderente: 1,
      nomeConsultor: 2,
      dataFiliacao: 3,
      statusLigacao: 4
    };

    const mesAtual = mes || new Date().getMonth() + 1;
    const anoAtual = ano || new Date().getFullYear();

    const consultoresData = {};
    let totalGeral = { total: 0, comLigacao: 0, semLigacao: 0, cancelado: 0 };

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      
      if (!row[idx.matricula] && !row[idx.nomeAderente] && !row[idx.nomeConsultor]) continue;
      
      if (row[idx.dataFiliacao]) {
        const dataParsed = parseDate(row[idx.dataFiliacao]);
        if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;
      }

      const consultorOriginal = String(row[idx.nomeConsultor] || "").trim();
      
      if (!consultorOriginal) continue;
      
      const consultor = consultorOriginal;
      
      const statusRaw = String(row[idx.statusLigacao] || "").trim();
      const status = normalizarStatusLigacao(statusRaw);

      if (!consultoresData[consultor]) {
        consultoresData[consultor] = { 
          nome: consultor, 
          total: 0, 
          comLigacao: 0, 
          semLigacao: 0, 
          cancelado: 0 
        };
      }

      const c = consultoresData[consultor];
      c.total++;
      totalGeral.total++;

      if (status === "COM LIGAÇÃO") { 
        c.comLigacao++; 
        totalGeral.comLigacao++; 
      }
      else if (status === "SEM LIGAÇÃO") { 
        c.semLigacao++; 
        totalGeral.semLigacao++; 
      }
      else if (status === "CANCELADO") { 
        c.cancelado++; 
        totalGeral.cancelado++; 
      }
      else {
        c.semLigacao++; 
        totalGeral.semLigacao++; 
      }
    }

    const consultoresComVendas = Object.values(consultoresData)
      .filter(c => c.total > 0)
      .sort((a, b) => {
        if (b.total !== a.total) return b.total - a.total;
        return a.nome.localeCompare(b.nome);
      });

    const responseData = {
      mes: MESES[mesAtual - 1],
      ano: anoAtual,
      geral: totalGeral,
      consultores: consultoresComVendas
    };

    criarDashboardRefuturiza(responseData);
    return createSuccessResponse(responseData);
  } catch (error) {
    return createErrorResponse(error.message);
  }
}

function criarDashboardRefuturiza(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Refuturiza");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Refuturiza");
  sheet.clear();

  sheet.getRange("A1:L1").merge().setValue("🔄 DASHBOARD REFUTURIZA - " + dados.mes.toUpperCase() + " " + dados.ano)
    .setBackground("#7c3aed")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(14);

  const row = 3;
  const g = dados.geral;
  const p = g.total > 0 ? Math.round((g.comLigacao / g.total) * 100) : 0;
  
  const cardGeral = [
    ["REFUTURIZA - TOTAL DA LOJA", ""],
    ["TOTAL", g.total],
    ["COM LIGAÇÃO", g.comLigacao],
    ["SEM LIGAÇÃO", g.semLigacao],
    ["CANCELADO", g.cancelado],
    ["", ""],
    ["% COM LIGAÇÃO", p + "%"]
  ];
  
  sheet.getRange(row, 1, cardGeral.length, 2)
    .setValues(cardGeral)
    .setBorder(true, true, true, true, true, true);
  
  sheet.getRange(row, 1, 1, 2)
    .merge()
    .setBackground("#7c3aed")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  const percentualCell = sheet.getRange(row + 6, 2);
  const corPercentual = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
  percentualCell
    .setBackground(corPercentual)
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  let currentRow = row + cardGeral.length + 3;
  
  if (dados.consultores.length > 0) {
    sheet.getRange(currentRow, 1, 1, 12)
      .merge()
      .setValue("CONSULTORES COM VENDAS")
      .setBackground("#4b5563")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    currentRow += 2;
    
    let col = 1;
    let consultorRow = currentRow;
    
    dados.consultores.forEach(c => {
      const pConsultor = c.total > 0 ? Math.round((c.comLigacao / c.total) * 100) : 0;
      
      const card = [
        [c.nome, ""],
        ["TOTAL", c.total],
        ["COM LIGAÇÃO", c.comLigacao],
        ["SEM LIGAÇÃO", c.semLigacao],
        ["CANCELADO", c.cancelado],
        ["", ""],
        ["% COM LIGAÇÃO", pConsultor + "%"]
      ];
      
      sheet.getRange(consultorRow, col, card.length, 2)
        .setValues(card)
        .setBorder(true, true, true, true, true, true);
      
      sheet.getRange(consultorRow, col, 1, 2)
        .merge()
        .setBackground("#7c3aed")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      const corPercentualConsultor = pConsultor >= 90 ? "#10b981" : pConsultor >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(consultorRow + 6, col + 1)
        .setBackground(corPercentualConsultor)
        .setFontColor("#ffffff")
        .setFontWeight("bold");
      
      col += 3;
      if (col > 10) {
        col = 1;
        consultorRow += card.length + 1;
      }
    });
    
    if (col !== 1) {
      currentRow = consultorRow + 8;
    } else {
      currentRow = consultorRow + 1;
    }
  } else {
    sheet.getRange(currentRow, 1, 1, 12)
      .merge()
      .setValue("ℹ️ Nenhum consultor com vendas neste período")
      .setBackground("#f59e0b")
      .setFontColor("#000000")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    currentRow += 2;
  }
  
  currentRow += 2;
  
  const resumoFinal = [
    ["📈 RESUMO FINAL", ""],
    ["Total de Consultores:", dados.consultores.length],
    ["Total de Vendas:", g.total],
    ["Média por Consultor:", dados.consultores.length > 0 ? Math.round(g.total / dados.consultores.length) : 0],
    ["Melhor %:", dados.consultores.length > 0 ? 
      Math.max(...dados.consultores.map(c => c.total > 0 ? Math.round((c.comLigacao / c.total) * 100) : 0)) + "%" : "0%"],
    ["Pior %:", dados.consultores.length > 0 ? 
      Math.min(...dados.consultores.map(c => c.total > 0 ? Math.round((c.comLigacao / c.total) * 100) : 0)) + "%" : "0%"]
  ];
  
  sheet.getRange(currentRow, 1, resumoFinal.length, 2)
    .setValues(resumoFinal)
    .setBorder(true, true, true, true, true, true);
  
  sheet.getRange(currentRow, 1, 1, 2)
    .merge()
    .setBackground("#000000")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  for (let i = 1; i <= 12; i++) {
    sheet.setColumnWidth(i, 160);
  }
  
  sheet.setFrozenRows(1);
  
  // CORREÇÃO: NÃO criar filtro se houver células mescladas
  try {
    const ultimaLinha = sheet.getLastRow();
    const rangeA1 = sheet.getRange("A1");
    if (rangeA1.isPartOfMerge()) {
      sheet.getRange(2, 1, ultimaLinha - 1, 12).createFilter();
    } else {
      sheet.getRange(1, 1, ultimaLinha, 12).createFilter();
    }
  } catch (e) {
    console.log("Não foi possível criar filtro: " + e.message);
  }
}

// ============================================================================
// FUNÇÕES AUXILIARES
// ============================================================================

function normalizarConsultorResidual(nome) {
  if (!nome) return "TELEVENDAS";
  let n = nome.toUpperCase();
  if (n.includes("WEB SITE")) return "WEB SITE";
  
  for (const consultora in CONSULTORAS_RETENCAO) {
    if (n.includes(consultora.toUpperCase())) {
      return "OUTROS";
    }
  }
  
  if (OUTROS_LISTA.some(o => n.includes(o.toUpperCase()))) return "OUTROS";
  let nomeLimpo = n.replace(PREFIXO_AMOR, "").trim();
  let primeiroNome = nomeLimpo.split(" ")[0];
  for (const setor in CONSULTORES_MAP) {
    const encontrado = CONSULTORES_MAP[setor].find(c => 
      nomeLimpo === c.toUpperCase() || primeiroNome === c.toUpperCase().split(" ")[0]
    );
    if (encontrado) return encontrado;
  }
  return "TELEVENDAS";
}

function normalizarConsultorOriginal(nome) {
  if (!nome) return "OUTROS";
  let n = nome.toUpperCase();
  if (n.startsWith(PREFIXO_AMOR)) {
    const primeiroNome = n.replace(PREFIXO_AMOR, "").split(" ")[0];
    for (const setor in CONSULTORES_MAP) {
      const encontrado = CONSULTORES_MAP[setor].find(c => c.toUpperCase().startsWith(primeiroNome));
      if (encontrado) return encontrado;
    }
  }
  return nome;
}

function renderizarTresCardsPrincipais(sheet, dados, tipo) {
  const row = 3;
  const g = dados.geral;
  const l = tipo === "DOC" ? dados.vendasLoja : dados.appLoja;
  const w = tipo === "DOC" ? dados.vendasWeb : dados.appWeb;
  
  const pG = g.total > 0 ? Math.round((tipo === "DOC" ? (g.aprovados || 0) : (g.sim || 0)) / g.total * 100) : 0;
  const pL = l.total > 0 ? Math.round((tipo === "DOC" ? (l.aprovados || 0) : (l.sim || 0)) / l.total * 100) : 0;
  const pW = w.total > 0 ? Math.round((tipo === "DOC" ? (w.aprovados || 0) : (w.sim || 0)) / w.total * 100) : 0;
  
  const labels = tipo === "DOC" 
    ? ["TOTAL", "APROVADOS", "PENDÊNCIAS", "  NÃO ENVIADO", "  EXPIRADO", "  PENDENTE", "% APROVADOS"] 
    : ["TOTAL", "COM APP (SIM)", "SEM APP (NÃO)", "CANCELADO", "OUTROS", "", "% COM APP"];
  
  const card1 = [[tipo === "DOC" ? "VENDAS TOTAIS" : "APP - TOTAL", ""], ...labels.map((lab, i) => {
    if (!lab) return ["", ""];
    if (i === 0) return [lab, g.total || 0];
    if (tipo === "DOC") {
      if (i === 1) return [lab, g.aprovados || 0];
      if (i === 2) return [lab, g.pendencias || 0];
      if (i === 3) return [lab, g.naoEnviado || 0];
      if (i === 4) return [lab, g.expirado || 0];
      if (i === 5) return [lab, g.pendente || 0];
      return [lab, pG + "%"];
    } else {
      if (i === 1) return [lab, g.sim || 0];
      if (i === 2) return [lab, g.nao || 0];
      if (i === 3) return [lab, g.cancelado || 0];
      if (i === 4) return [lab, g.outros || 0];
      return [lab, pG + "%"];
    }
  })];
  
  const card2 = [[tipo === "DOC" ? "VENDAS LOJA" : "APP - LOJA", ""], ...labels.map((lab, i) => {
    if (!lab) return ["", ""];
    if (i === 0) return [lab, l.total || 0];
    if (tipo === "DOC") {
      if (i === 1) return [lab, l.aprovados || 0];
      if (i === 2) return [lab, l.pendencias || 0];
      if (i === 3) return [lab, l.naoEnviado || 0];
      if (i === 4) return [lab, l.expirado || 0];
      if (i === 5) return [lab, l.pendente || 0];
      return [lab, pL + "%"];
    } else {
      if (i === 1) return [lab, l.sim || 0];
      if (i === 2) return [lab, l.nao || 0];
      if (i === 3) return [lab, l.cancelado || 0];
      if (i === 4) return [lab, l.outros || 0];
      return [lab, pL + "%"];
    }
  })];
  
  const card3 = [[tipo === "DOC" ? "VENDAS WEB/TELEVENDAS" : "APP - WEB/TELEVENDAS", ""], ...labels.map((lab, i) => {
    if (!lab) return ["", ""];
    if (i === 0) return [lab, w.total || 0];
    if (tipo === "DOC") {
      if (i === 1) return [lab, w.aprovados || 0];
      if (i === 2) return [lab, w.pendencias || 0];
      if (i === 3) return [lab, w.naoEnviado || 0];
      if (i === 4) return [lab, w.expirado || 0];
      if (i === 5) return [lab, w.pendente || 0];
      return [lab, pW + "%"];
    } else {
      if (i === 1) return [lab, w.sim || 0];
      if (i === 2) return [lab, w.nao || 0];
      if (i === 3) return [lab, w.cancelado || 0];
      if (i === 4) return [lab, w.outros || 0];
      return [lab, pW + "%"];
    }
  })];
  
  sheet.getRange(row, 1, card1.length, 2).setValues(card1).setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 1, 1, 2).merge().setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  
  const percentRow = tipo === "DOC" ? row + 7 : row + 7;
  sheet.getRange(percentRow, 2).setBackground(pG >= 90 ? "#10b981" : pG >= 80 ? "#f59e0b" : "#ef4444").setFontColor("#ffffff").setFontWeight("bold");
  
  sheet.getRange(row, 4, card2.length, 2).setValues(card2).setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 4, 1, 2).merge().setBackground("#78350f").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(percentRow, 5).setBackground(pL >= 90 ? "#10b981" : pL >= 80 ? "#f59e0b" : "#ef4444").setFontColor("#ffffff").setFontWeight("bold");
  
  sheet.getRange(row, 7, card3.length, 2).setValues(card3).setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 7, 1, 2).merge().setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(percentRow, 8).setBackground(pW >= 90 ? "#10b981" : pW >= 80 ? "#f59e0b" : "#ef4444").setFontColor("#ffffff").setFontWeight("bold");
}

function parseDate(dataStr) {
  if (!dataStr) return { mes: 0, ano: 0 };
  
  const d = new Date(dataStr);
  if (!isNaN(d.getTime())) return { mes: d.getMonth() + 1, ano: d.getFullYear() };
  
  const str = String(dataStr).trim();
  
  const partesBarra = str.split("/");
  if (partesBarra.length >= 2) {
    const mes = parseInt(partesBarra[1], 10);
    let ano;
    
    if (partesBarra.length >= 3) {
      ano = parseInt(partesBarra[2], 10);
      if (ano < 100) {
        ano += 2000;
      }
    } else {
      ano = new Date().getFullYear();
    }
    
    if (!isNaN(mes) && !isNaN(ano)) {
      return { mes: mes, ano: ano };
    }
  }
  
  const partesMesAno = str.split("/");
  if (partesMesAno.length === 2) {
    const mes = parseInt(partesMesAno[0], 10);
    const ano = parseInt(partesMesAno[1], 10);
    
    if (!isNaN(mes) && !isNaN(ano)) {
      const anoFinal = ano < 100 ? ano + 2000 : ano;
      return { mes: mes, ano: anoFinal };
    }
  }
  
  const numeros = str.match(/\d+/g);
  if (numeros && numeros.length >= 2) {
    const mes = parseInt(numeros[1], 10);
    const ano = numeros.length >= 3 ? parseInt(numeros[2], 10) : new Date().getFullYear();
    
    if (!isNaN(mes) && mes >= 1 && mes <= 12) {
      const anoFinal = ano < 100 ? ano + 2000 : ano;
      return { mes: mes, ano: anoFinal };
    }
  }
  
  return { mes: 0, ano: 0 };
}

function createSuccessResponse(data) { 
  return { 
    status: "success", 
    timestamp: new Date().toISOString(), 
    data: data
  }; 
}

function createErrorResponse(msg) { 
  return { 
    status: "error", 
    timestamp: new Date().toISOString(), 
    error: msg 
  }; 
}

function atualizarTodosDashboards() {
  const mes = new Date().getMonth() + 1;
  const ano = new Date().getFullYear();
  getDocumentacaoData(mes, ano);
  getAppData(mes, ano);
  getAdimplenciaData(mes, ano);
  getRefuturizaData(mes, ano);
  getRecorrenciaVendedorData(); // SEM parâmetros mes/ano
  gerarDashboardRecorrencia();
  SpreadsheetApp.getUi().alert("✅ Todos os 6 dashboards foram atualizados!");
}

// ============================================================================
// FUNÇÕES ADICIONAIS
// ============================================================================

function showConfigDialog() {
  const html = `
    <div style="padding: 20px; font-family: Arial;">
      <h2>⚙️ Configurações V20.6</h2>
      <p><strong>🎯 MODIFICAÇÕES V20.6:</strong></p>
      <ol>
        <li><strong>🤝 RECORRÊNCIA VENDEDOR: SEM FILTRO DE MÊS</strong><br>
           Agora soma <strong>TODOS OS DADOS</strong> da planilha completa</li>
        <li><strong>✅ Menu simplificado</strong><br>
           Apenas 1 opção para abrir o dashboard completo</li>
        <li><strong>📊 Dados Completos</strong><br>
           Análise histórica de todos os registros disponíveis</li>
      </ol>
      
      <p><strong>📊 Dashboards Disponíveis (6):</strong></p>
      <ul>
        <li>📁 Documentação (por mês)</li>
        <li>📱 App (por mês)</li>
        <li>💳 Adimplência (por mês)</li>
        <li>📈 Recorrência (por mês fixo)</li>
        <li>🤝 Recorrência Vendedor (<strong>TODOS OS DADOS</strong>)</li>
        <li>🔄 Refuturiza (por mês)</li>
      </ul>
      
      <p><strong>🔧 Funcionalidades Principais:</strong></p>
      <ul>
        <li><strong>Gerar Mês Passado</strong>: Interface para criar qualquer mês/ano</li>
        <li><strong>API completa</strong>: Todos endpoints funcionando</li>
        <li><strong>Sistema estável</strong>: Sem erros de filtro ou células mescladas</li>
        <li><strong>Recorrência Vendedor</strong>: Análise completa sem filtro temporal</li>
      </ul>
      
      <button onclick="google.script.host.close()" style="background: #10b981; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">Fechar</button>
    </div>
  `;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(550),
    'Configurações V20.6'
  );
}

function deployAPI() {
  const webAppUrl = ScriptApp.getService().getUrl();
  const html = `
    <div style="padding: 20px; font-family: Arial;">
      <h2>🚀 API V20.6 Implantada!</h2>
      <p>Sua API está disponível em:</p>
      <div style="background: #f0f0f0; padding: 10px; border-radius: 5px; margin: 10px 0;">
        <code>${webAppUrl}</code>
      </div>
      
      <p><strong>📋 Exemplos de uso (6 endpoints):</strong></p>
      <div style="background: #f8fafc; padding: 10px; border-radius: 5px; margin: 10px 0; font-family: monospace;">
        <div>${webAppUrl}?endpoint=documentacao&mes=1&ano=2024</div>
        <div>${webAppUrl}?endpoint=app&mes=2&ano=2024</div>
        <div>${webAppUrl}?endpoint=adimplencia&mes=3&ano=2024</div>
        <div>${webAppUrl}?endpoint=recorrencia&mes=4&ano=2024</div>
        <div>${webAppUrl}?endpoint=<strong>recorrencia_vendedor</strong> (sem parâmetros)</div>
        <div>${webAppUrl}?endpoint=refuturiza&mes=6&ano=2024</div>
      </div>
      
      <p><strong>🎯 Recorrência Vendedor (Dados Completos)</strong></p>
      <div style="background: #f0f9ff; padding: 10px; border-radius: 5px; margin: 10px 0;">
        <strong>⚠️ AGORA SEM FILTRO DE MÊS:</strong><br>
        • Processa <strong>TODOS OS DADOS</strong> da planilha<br>
        • Análise histórica completa<br>
        • Ideal para acompanhamento geral<br>
        • Endpoint: <code>${webAppUrl}?endpoint=recorrencia_vendedor</code>
      </div>
      
      <button onclick="google.script.host.close()" style="background: #10b981; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">Fechar</button>
    </div>
  `;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(650).setHeight(500),
    'API Implantada V20.6'
  );
}

// ============================================================================
// FUNÇÃO PARA O HTML MANTER O MÊS SELECIONADO
// ============================================================================

function getDashboardDataForHTML(tipo, mes, ano) {
  // Esta função é chamada pelo HTML para obter dados de qualquer mês
  try {
    let resultado;
    
    switch(tipo) {
      case 'documentacao':
        resultado = getDocumentacaoData(mes, ano);
        break;
      case 'app':
        resultado = getAppData(mes, ano);
        break;
      case 'adimplencia':
        resultado = getAdimplenciaData(mes, ano);
        break;
      case 'recorrencia':
        resultado = processarDadosRecorrenciaPorMes(mes, ano, true);
        break;
      case 'recorrencia_vendedor':
        resultado = getRecorrenciaVendedorData(); // SEM parâmetros mes/ano
        break;
      case 'refuturiza':
        resultado = getRefuturizaData(mes, ano);
        break;
      default:
        return { status: "error", error: "Tipo de dashboard desconhecido" };
    }
    
    // Adicionar informações do período para o HTML
    if (resultado.status === "success" && resultado.data) {
      resultado.data.tipo = tipo;
      resultado.data.mesNumero = mes;
      resultado.data.anoNumero = ano;
    }
    
    return resultado;
    
  } catch (error) {
    return { status: "error", error: error.message };
  }
}

// ============================================================================
// TESTES
// ============================================================================

function testarTodosDashboardsMesesPassados() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ano = new Date().getFullYear();
    
    let resultados = [];
    
    // Testar Documentação
    for (let mes = 1; mes <= 3; mes++) {
      const resultado = getDocumentacaoData(mes, ano);
      resultados.push(`Documentação ${MESES[mes-1]}: ${resultado.status === "success" ? "✅" : "❌"}`);
    }
    
    // Testar App
    for (let mes = 1; mes <= 3; mes++) {
      const resultado = getAppData(mes, ano);
      resultados.push(`App ${MESES[mes-1]}: ${resultado.status === "success" ? "✅" : "❌"}`);
    }
    
    // Testar Adimplência
    for (let mes = 1; mes <= 3; mes++) {
      const resultado = getAdimplenciaData(mes, ano);
      resultados.push(`Adimplência ${MESES[mes-1]}: ${resultado.status === "success" ? "✅" : "❌"}`);
    }
    
    // Testar Recorrência
    for (let mes = 1; mes <= 3; mes++) {
      const resultado = processarDadosRecorrenciaPorMes(mes, ano, true);
      resultados.push(`Recorrência ${MESES[mes-1]}: ${resultado.status === "success" ? "✅" : "❌"}`);
    }
    
    // Testar Recorrência Vendedor (APENAS 1 TESTE - SEM MÊS)
    const resultadoRecVendedor = getRecorrenciaVendedorData();
    resultados.push(`Recorrência Vendedor (TODOS OS DADOS): ${resultadoRecVendedor.status === "success" ? "✅" : "❌"}`);
    
    // Testar Refuturiza
    for (let mes = 1; mes <= 3; mes++) {
      const resultado = getRefuturizaData(mes, ano);
      resultados.push(`Refuturiza ${MESES[mes-1]}: ${resultado.status === "success" ? "✅" : "❌"}`);
    }
    
    const mensagem = "🧪 TESTE DE TODOS OS DASHBOARDS\n\n" + resultados.join("\n");
    
    ui.alert(mensagem);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("❌ Erro no teste: " + error.message);
  }
}

function testarCalculoOkBi() {
  const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
  const sheet = ssDados.getSheetByName("ADIMPLENCIA");
  const data = sheet.getDataRange().getValues();
  const headers = data[1];
  
  const idx = {
    consultor: headers.indexOf("Consultor"),
    data: headers.indexOf("Data"),
    kyc: headers.indexOf("KYC"),
    mensalidadeOk: headers.indexOf("MENSALIDADE OK"),
    statusBi: headers.indexOf("STATUS BI")
  };
  
  let totalOkBi = 0;
  let exemplos = [];
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const consultor = String(row[idx.consultor] || "").trim();
    const dataStr = String(row[idx.data] || "").trim();
    const kyc = String(row[idx.kyc] || "").trim();
    const mensalidade = String(row[idx.mensalidadeOk] || "").trim();
    const statusBi = String(row[idx.statusBi] || "").trim();
    
    // Verificar se é janeiro
    if (dataStr.includes("/01") || dataStr.includes("01/")) {
      if (consultor === "FLÁVIA") {
        if (statusBi === "OK" && mensalidade === "OK") {
          totalOkBi++;
          exemplos.push(`Matrícula: ${row[2]}, Data: ${dataStr}, KYC: ${kyc}`);
        }
      }
    }
  }
  
  console.log(`Total okBi para FLÁVIA em janeiro: ${totalOkBi}`);
  console.log("Exemplos:", exemplos.slice(0, 5));
  
  SpreadsheetApp.getUi().alert(`okBi para FLÁVIA em janeiro: ${totalOkBi}\n(Esperado: ~55)`);
}

function debugDetalhadoOkBi() {
  const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
  const sheet = ssDados.getSheetByName("ADIMPLENCIA");
  const data = sheet.getDataRange().getValues();
  const headers = data[1];
  
  // Log para ver os cabeçalhos
  console.log("Cabeçalhos encontrados:");
  headers.forEach((header, index) => {
    console.log(`${index}: "${header}"`);
  });
  
  const idx = {
    consultor: headers.indexOf("Consultor"),
    data: headers.indexOf("Data"),
    kyc: headers.indexOf("KYC"),
    mensalidadeOk: headers.indexOf("MENSALIDADE OK"),
    statusBi: headers.indexOf("STATUS BI")
  };
  
  console.log("\nÍndices encontrados:");
  console.log("Consultor:", idx.consultor);
  console.log("Data:", idx.data);
  console.log("KYC:", idx.kyc);
  console.log("MENSALIDADE OK:", idx.mensalidadeOk);
  console.log("STATUS BI:", idx.statusBi);
  
  let contadores = {
    totalFlavia: 0,
    janeiroFlavia: 0,
    statusBiOk: 0,
    mensalidadeOk: 0,
    ambosOk: 0,
    exemplosProblema: []
  };
  
  for (let i = 2; i < Math.min(data.length, 20); i++) { // Vamos ver apenas 20 primeiras linhas
    const row = data[i];
    const consultor = String(row[idx.consultor] || "").trim();
    const dataStr = String(row[idx.data] || "").trim();
    const kyc = String(row[idx.kyc] || "").trim();
    const mensalidade = String(row[idx.mensalidadeOk] || "").trim();
    const statusBi = String(row[idx.statusBi] || "").trim();
    
    // Verificar se é Flávia
    if (consultor.toUpperCase().includes("FLÁVIA") || consultor.toUpperCase().includes("FLAVIA")) {
      contadores.totalFlavia++;
      
      // Verificar se é janeiro
      const isJaneiro = dataStr.includes("/01") || dataStr.includes("01/") || 
                       dataStr.includes("01-") || dataStr.includes("-01");
      
      if (isJaneiro) {
        contadores.janeiroFlavia++;
        
        if (statusBi.toUpperCase() === "OK") {
          contadores.statusBiOk++;
        }
        
        if (mensalidade.toUpperCase() === "OK") {
          contadores.mensalidadeOk++;
        }
        
        if (statusBi.toUpperCase() === "OK" && mensalidade.toUpperCase() === "OK") {
          contadores.ambosOk++;
        } else {
          contadores.exemplosProblema.push({
            linha: i + 1,
            consultor: consultor,
            data: dataStr,
            statusBi: statusBi,
            mensalidade: mensalidade,
            kyc: kyc,
            matricula: row[2] || ""
          });
        }
      }
    }
  }
  
  console.log("\n=== RESULTADOS DO DEBUG ===");
  console.log("Total Flávia (todas datas):", contadores.totalFlavia);
  console.log("Flávia em janeiro:", contadores.janeiroFlavia);
  console.log("Status BI = 'OK':", contadores.statusBiOk);
  console.log("Mensalidade OK = 'OK':", contadores.mensalidadeOk);
  console.log("AMBOS 'OK' (okBi):", contadores.ambosOk);
  
  console.log("\n=== EXEMPLOS QUE NÃO ENCAIXAM ===");
  contadores.exemplosProblema.slice(0, 5).forEach(exemplo => {
    console.log(`Linha ${exemplo.linha}: StatusBI="${exemplo.statusBi}", Mensalidade="${exemplo.mensalidade}"`);
  });
  
  let mensagem = `🔍 DEBUG DETALHADO OK_BI\n\n`;
  mensagem += `Total registros Flávia: ${contadores.totalFlavia}\n`;
  mensagem += `Flávia em janeiro: ${contadores.janeiroFlavia}\n`;
  mensagem += `Status BI = "OK": ${contadores.statusBiOk}\n`;
  mensagem += `Mensalidade OK = "OK": ${contadores.mensalidadeOk}\n`;
  mensagem += `AMBOS "OK" (okBi): ${contadores.ambosOk}\n\n`;
  
  if (contadores.exemplosProblema.length > 0) {
    mensagem += `📋 Exemplos não contados:\n`;
    contadores.exemplosProblema.slice(0, 3).forEach(exemplo => {
      mensagem += `• Linha ${exemplo.linha}: BI="${exemplo.statusBi}", Mens="${exemplo.mensalidade}"\n`;
    });
  }
  
  SpreadsheetApp.getUi().alert(mensagem);
}

// ============================================================================
// TESTE ESPECÍFICO PARA RECORRÊNCIA VENDEDOR (DADOS COMPLETOS)
// ============================================================================

function testarRecorrenciaVendedor() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const resultado = getRecorrenciaVendedorData();
    
    if (resultado.status === "success") {
      const dados = resultado.data;
      let mensagem = `🧪 TESTE RECORRÊNCIA VENDEDOR (DADOS COMPLETOS)\n\n`;
      mensagem += `✅ Dashboard gerado com sucesso!\n\n`;
      mensagem += `📊 Métricas Gerais (TODOS OS MESES):\n`;
      mensagem += `• Total Vendas Promoção: ${dados.geral.totalVendasPromocao}\n`;
      mensagem += `• Total OK: ${dados.geral.totalOk}\n`;
      mensagem += `• Total Em Aberto: ${dados.geral.totalEmAberto}\n`;
      mensagem += `• Total Atraso: ${dados.geral.totalAtraso}\n`;
      mensagem += `• Total Outros: ${dados.geral.totalOutros || 0}\n`;
      mensagem += `• % Vendas OK: ${dados.geral.percentualVendasOk}%\n\n`;
      mensagem += `👥 Consultores Encontrados: ${dados.totalConsultores}\n`;
      mensagem += `📋 Total de Registros: ${dados.totalRegistros}\n`;
      
      if (dados.consultores.length > 0) {
        mensagem += `\n🏆 Top 3 Consultores:\n`;
        dados.consultores.slice(0, 3).forEach((c, i) => {
          mensagem += `${i+1}. ${c.nome}: ${c.totalVendasPromocao} vendas (${c.percentualVendasOk}% OK)\n`;
        });
      }
      
      mensagem += `\n🎯 Setores Ativos:\n`;
      for (const setor in dados.dadosPorSetor) {
        if (dados.dadosPorSetor[setor].length > 0) {
          const totalSetor = dados.dadosPorSetor[setor].reduce((sum, c) => sum + c.totalVendasPromocao, 0);
          mensagem += `• ${setor}: ${dados.dadosPorSetor[setor].length} consultor${dados.dadosPorSetor[setor].length !== 1 ? 'es' : ''} (${totalSetor} vendas)\n`;
        }
      }
      
      ui.alert(mensagem);
    } else {
      ui.alert(`❌ Erro: ${resultado.error}`);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("❌ Erro no teste: " + error.message);
  }
}
function debugRecorrenciaVendedor() {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    let sheet = ssDados.getSheetByName(NOME_ABA_RECORRENCIA_VENDAS);
    
    if (!sheet) {
      const sheets = ssDados.getSheets();
      sheet = sheets.find(s => {
        const nome = s.getName().toUpperCase();
        return nome.includes("RECORRENCIA VENDAS") || 
               nome.includes("RECORRENCIA_VENDAS") ||
               nome.includes("RECORRENCIA-VENDAS");
      });
    }
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Sheet 'RECORRENCIA VENDAS' não encontrada");
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Coletar todos os nomes de consultores únicos
    const consultoresUnicos = new Set();
    const consultoresCompletos = {};
    
    for (let i = 1; i < Math.min(data.length, 100); i++) { // Verificar até 100 linhas
      const row = data[i];
      const consultorOriginal = String(row[2] || "").trim();
      
      if (consultorOriginal) {
        consultoresUnicos.add(consultorOriginal);
        
        // Normalizar e ver setor
        const consultorNormalizado = normalizarNomeConsultorRecorrencia(consultorOriginal);
        const setor = determinarSetorConsultor(consultorNormalizado);
        
        if (!consultoresCompletos[consultorNormalizado]) {
          consultoresCompletos[consultorNormalizado] = {
            original: consultorOriginal,
            normalizado: consultorNormalizado,
            setor: setor,
            noMapeamento: consultorEstaNoMapeamento(consultorNormalizado)
          };
        }
      }
    }
    
    // Criar mensagem de debug
    let mensagem = "🔍 DEBUG RECORRÊNCIA VENDEDOR\n\n";
    mensagem += `Total de consultores únicos encontrados: ${consultoresUnicos.size}\n\n`;
    
    mensagem += "📋 CONSULTORES DETECTADOS:\n";
    for (const [normalizado, info] of Object.entries(consultoresCompletos)) {
      const status = info.noMapeamento ? "✅" : "❌";
      mensagem += `${status} "${info.original}" → "${normalizado}" (Setor: ${info.setor})\n`;
    }
    
    mensagem += "\n🎯 CONSULTORES ESPERADOS:\n";
    for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
      mensagem += `${setor}: ${CONSULTORES_RECORRENCIA_VENDAS[setor].join(", ")}\n`;
    }
    
    SpreadsheetApp.getUi().alert(mensagem);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("Erro no debug: " + error.message);
  }
}
function debugEstruturaDadosRecorrenciaVendedor() {
  try {
    const resultado = getRecorrenciaVendedorData();
    
    if (resultado.status === "success") {
      const dados = resultado.data;
      
      let mensagem = "🔍 DEBUG ESTRUTURA DE DADOS\n\n";
      
      mensagem += `Total consultores: ${dados.consultores.length}\n\n`;
      
      mensagem += "📋 TODOS OS CONSULTORES:\n";
      dados.consultores.forEach((c, i) => {
        mensagem += `${i+1}. ${c.nome} (Setor: ${c.setor}) - ${c.totalVendasPromocao} vendas\n`;
      });
      
      mensagem += "\n🎯 DADOS POR SETOR:\n";
      if (dados.dadosPorSetor) {
        for (const setor in dados.dadosPorSetor) {
          mensagem += `${setor}: ${dados.dadosPorSetor[setor].length} consultores\n`;
          dados.dadosPorSetor[setor].forEach(c => {
            mensagem += `   • ${c.nome}: ${c.totalVendasPromocao} vendas\n`;
          });
        }
      } else {
        mensagem += "❌ dadosPorSetor é undefined ou null!\n";
      }
      
      // Verificar setores manualmente
      mensagem += "\n🔍 VERIFICAÇÃO MANUAL POR SETOR:\n";
      const setores = ["VENDAS", "REFILIACAO", "RECEPCAO"];
      setores.forEach(setor => {
        const consultoresSetor = dados.consultores.filter(c => c.setor === setor);
        mensagem += `${setor}: ${consultoresSetor.length} consultores\n`;
      });
      
      SpreadsheetApp.getUi().alert(mensagem);
      
      // Log adicional no console
      console.log("Estrutura completa de dados:", dados);
      console.log("dadosPorSetor existe?", !!dados.dadosPorSetor);
      if (dados.dadosPorSetor) {
        console.log("Keys de dadosPorSetor:", Object.keys(dados.dadosPorSetor));
      }
      
    } else {
      SpreadsheetApp.getUi().alert(`❌ Erro: ${resultado.error}`);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("Erro no debug: " + error.message);
  }
}
// ============================================================================
// SYNC SUPABASE - Adicione este bloco ao seu arquivo .gs existente
// Substitui a URL e chave pelas suas do Supabase
// ============================================================================

// ⚠️ CONFIGURE AQUI — pegue no Supabase: Settings → API
const SUPABASE_URL  = 'https://vycjtmjvkwvxunxtkdyi.supabase.co';
const SUPABASE_KEY  = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ5Y2p0bWp2a3d2eHVueHRrZHlpIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3MTYwNjY5NiwiZXhwIjoyMDg3MTgyNjk2fQ.sT4DiipgLk4CfZqa-FP1bATcb6wreg0b_usNXkyBpbk'; // Settings → API → service_role (secret)

// ============================================================================
// TRIGGER DE EDIÇÃO — configure uma vez só executando: configurarTriggerEdicao()
// ============================================================================

/**
 * Execute esta função UMA VEZ pelo menu Apps Script para instalar o trigger.
 * Depois disso toda edição na planilha de dados dispara o sync automático.
 */
function configurarTriggerEdicao() {
  // Remove triggers antigos para não duplicar
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onEdicaoPlanilha') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Cria novo trigger na planilha de DADOS (onde os consultores editam)
  ScriptApp.newTrigger('onEdicaoPlanilha')
    .forSpreadsheet(ID_PLANILHA_DADOS)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Trigger instalado!\n\n' +
    'Agora toda edição na planilha de dados vai sincronizar automaticamente com o Supabase.\n\n' +
    'Delay estimado: 2-5 segundos após salvar.'
  );
}

/**
 * Chamado automaticamente a cada edição na planilha de dados.
 * Detecta qual aba foi editada e sincroniza só aquele endpoint.
 */
function onEdicaoPlanilha(e) {
  if (!e || !e.range) return;

  const abaEditada = e.range.getSheet().getName().toLowerCase();
  const mes  = new Date().getMonth() + 1;
  const ano  = new Date().getFullYear();

  // Sincroniza apenas o endpoint relacionado à aba editada
  if (abaEditada.includes('documenta')) {
    syncDocumentacao(mes, ano);
  } else if (abaEditada.includes('app retencao') || abaEditada.includes('app retenção')) {
    syncApp(mes, ano);
  } else if (abaEditada === 'app') {
    syncApp(mes, ano);
  } else if (abaEditada.includes('adimpl')) {
    syncAdimplencia(mes, ano);
  } else if (abaEditada === 'recorrencia') {
    syncRecorrencia(mes, ano);
  } else if (abaEditada.includes('recorrencia vendas')) {
    syncRecorrenciaVendedor();
  } else if (abaEditada.includes('refuturiza')) {
    syncRefuturiza(mes, ano);
  }
  // Se editou outra aba (configurações, etc.), não faz nada
}

// ============================================================================
// FUNÇÕES DE SYNC — cada uma busca os dados e envia pro Supabase
// ============================================================================

function syncDocumentacao(mes, ano) {
  try {
    const resultado = getDocumentacaoData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;

    supabaseUpsert('documentacao', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total: d.geral.total,
      geral_aprovados: d.geral.aprovados,
      geral_pendencias: d.geral.pendencias,
      geral_reprovados: d.geral.reprovados,
      geral_expirado: d.geral.expirado,
      geral_pendente: d.geral.pendente,
      geral_nao_enviado: d.geral.naoEnviado,
      loja_total: d.vendasLoja.total,
      loja_aprovados: d.vendasLoja.aprovados,
      loja_pendencias: d.vendasLoja.pendencias,
      loja_reprovados: d.vendasLoja.reprovados,
      loja_expirado: d.vendasLoja.expirado,
      loja_pendente: d.vendasLoja.pendente,
      loja_nao_enviado: d.vendasLoja.naoEnviado,
      web_total: d.vendasWeb.total,
      web_aprovados: d.vendasWeb.aprovados,
      web_pendencias: d.vendasWeb.pendencias,
      web_reprovados: d.vendasWeb.reprovados,
      web_expirado: d.vendasWeb.expirado,
      web_pendente: d.vendasWeb.pendente,
      web_nao_enviado: d.vendasWeb.naoEnviado,
      consultores: JSON.stringify(d.consultores),
      atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] documentacao ' + mes + '/' + ano + ' OK');
  } catch(err) {
    console.error('[sync] documentacao ERRO: ' + err.message);
  }
}

function syncApp(mes, ano) {
  try {
    const resultado = getAppData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;

    supabaseUpsert('app_dashboard', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total: d.geral.total,
      geral_sim: d.geral.sim,
      geral_nao: d.geral.nao,
      geral_cancelado: d.geral.cancelado,
      geral_outros: d.geral.outros || 0,
      loja_total: d.appLoja.total,
      loja_sim: d.appLoja.sim,
      loja_nao: d.appLoja.nao,
      loja_cancelado: d.appLoja.cancelado,
      loja_outros: d.appLoja.outros || 0,
      web_total: d.appWeb.total,
      web_sim: d.appWeb.sim,
      web_nao: d.appWeb.nao,
      web_cancelado: d.appWeb.cancelado,
      web_outros: d.appWeb.outros || 0,
      consultores: JSON.stringify(d.consultores),
      consultoras_retencao: JSON.stringify(d.consultorasRetencao || []),
      atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] app ' + mes + '/' + ano + ' OK');
  } catch(err) {
    console.error('[sync] app ERRO: ' + err.message);
  }
}

function syncAdimplencia(mes, ano) {
  try {
    const resultado = getAdimplenciaData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;
    const g = d.geral;

    supabaseUpsert('adimplencia', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total_trocas: g.totalTrocas,
      geral_mens_ok: g.mensOk,
      geral_mens_aberto: g.mensAberto,
      geral_mens_atraso: g.mensAtraso,
      geral_aprovados: g.aprovados,
      geral_pendentes: g.pendentes,
      geral_total_bi: g.totalBi,
      geral_fora_bi: g.foraBi,
      geral_ok_bi: g.okBi,
      geral_percentual_aprovado: g.percentualAprovado,
      consultores: JSON.stringify(d.consultores),
      atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] adimplencia ' + mes + '/' + ano + ' OK');
  } catch(err) {
    console.error('[sync] adimplencia ERRO: ' + err.message);
  }
}

function syncRecorrencia(mes, ano) {
  try {
    const resultado = processarDadosRecorrenciaPorMes(mes, ano, true);
    if (resultado.status !== 'success') return;

    supabaseUpsert('recorrencia', {
      mes: mes, ano: ano,
      mes_nome: MESES[mes - 1],
      dados: JSON.stringify(resultado.data),
      atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] recorrencia ' + mes + '/' + ano + ' OK');
  } catch(err) {
    console.error('[sync] recorrencia ERRO: ' + err.message);
  }
}

function syncRecorrenciaVendedor() {
  try {
    const resultado = getRecorrenciaVendedorData();
    if (resultado.status !== 'success') return;
    const d = resultado.data;
    const g = d.geral;

    // Recorrência vendedor tem apenas 1 linha — usa UPDATE simples
    supabaseUpdate('recorrencia_vendedor', {
      geral_total_vendas: g.totalVendasPromocao,
      geral_total_ok: g.totalOk,
      geral_total_em_aberto: g.totalEmAberto,
      geral_total_atraso: g.totalAtraso,
      geral_total_outros: g.totalOutros || 0,
      geral_percentual_ok: g.percentualVendasOk,
      consultores: JSON.stringify(d.consultores),
      dados_por_setor: JSON.stringify(d.dadosPorSetor),
      total_consultores: d.totalConsultores,
      total_registros: d.totalRegistros,
      atualizado_em: new Date().toISOString()
    }, 'id=eq.1');
    console.log('[sync] recorrencia_vendedor OK');
  } catch(err) {
    console.error('[sync] recorrencia_vendedor ERRO: ' + err.message);
  }
}

function syncRefuturiza(mes, ano) {
  try {
    const resultado = getRefuturizaData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;

    supabaseUpsert('refuturiza', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total: d.geral.total,
      geral_com_ligacao: d.geral.comLigacao,
      geral_sem_ligacao: d.geral.semLigacao,
      geral_cancelado: d.geral.cancelado,
      consultores: JSON.stringify(d.consultores),
      atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] refuturiza ' + mes + '/' + ano + ' OK');
  } catch(err) {
    console.error('[sync] refuturiza ERRO: ' + err.message);
  }
}

// ============================================================================
// SYNC COMPLETO — sincroniza todos os endpoints do mês atual de uma vez
// Use para inicialização ou re-sync manual
// ============================================================================
function syncTudoAgora() {
  const mes = new Date().getMonth() + 1;
  const ano = new Date().getFullYear();
  const ui  = SpreadsheetApp.getUi();

  ui.alert('🔄 Iniciando sync completo...\n\nAguarde, isso pode levar ~30 segundos.');

  syncDocumentacao(mes, ano);
  syncApp(mes, ano);
  syncAdimplencia(mes, ano);
  syncRecorrencia(mes, ano);
  syncRecorrenciaVendedor();
  syncRefuturiza(mes, ano);

  ui.alert(
    '✅ Sync completo!\n\n' +
    'Todos os dados do mês ' + MESES[mes - 1] + '/' + ano + ' foram enviados ao Supabase.\n\n' +
    'A partir de agora, toda edição nas planilhas sincroniza automaticamente.'
  );
}

// ============================================================================
// HELPERS HTTP — funções que chamam a API REST do Supabase
// ============================================================================

/**
 * UPSERT: insere ou atualiza baseado em conflict_columns (ex: "mes,ano")
 */
function supabaseUpsert(tabela, dados, conflictColumns) {
  const url = SUPABASE_URL + '/rest/v1/' + tabela;
  const options = {
    method: 'post',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates,return=minimal'
    },
    payload: JSON.stringify(dados),
    muteHttpExceptions: true
  };

  // Adiciona o parâmetro de conflict para upsert
  const urlComConflict = url + '?on_conflict=' + conflictColumns;
  const response = UrlFetchApp.fetch(urlComConflict, options);
  const code = response.getResponseCode();

  if (code !== 200 && code !== 201 && code !== 204) {
    throw new Error('Supabase ' + tabela + ' retornou ' + code + ': ' + response.getContentText());
  }
}

/**
 * UPDATE simples com filtro (ex: "id=eq.1")
 */
function supabaseUpdate(tabela, dados, filtro) {
  const url = SUPABASE_URL + '/rest/v1/' + tabela + '?' + filtro;
  const options = {
    method: 'patch',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(dados),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code !== 200 && code !== 204) {
    throw new Error('Supabase update ' + tabela + ' retornou ' + code + ': ' + response.getContentText());
  }
}
function syncHistoricoCompleto() {
  const ano = 2026;
  const meses = [2]; // adiciona mais meses conforme precisar
  
  meses.forEach(mes => {
    syncDocumentacao(mes, ano);
    syncApp(mes, ano);
    syncAdimplencia(mes, ano);
    syncRecorrencia(mes, ano);
    syncRefuturiza(mes, ano);
    Utilities.sleep(1000); // pausa 1s entre meses para não sobrecarregar
  });
  
  syncRecorrenciaVendedor(); // esse é sem filtro de mês, roda só uma vez
  
  SpreadsheetApp.getUi().alert('✅ Histórico completo sincronizado!');
}
