/**
 * API UNIFICADA - DASHBOARDS V20.6 - COM RECORRÊNCIA VENDEDOR (SEM FILTRO DE MÊS)
 * VERSÃO COMPLETA COM TODOS OS 6 DASHBOARDS:
 * 1. Documentação | 2. App | 3. Adimplência | 4. Recorrência | 5. Refuturiza | 6. Recorrência Vendedor
 *
 * ATUALIZAÇÃO V20.6.2 - DOCUMENTAÇÃO:
 * - Card do consultor redesenhado: nome abraça 4 colunas (banner único)
 * - PROMOÇÃO e NORMAL ficam à direita do card principal, vinculados ao dono
 * - Seção Promoção/Normal só é renderizada quando o consultor tem promo > 0
 * - Espaço de 1 coluna entre cards para melhor visualização (2 cards por linha)
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
  "JACKSON": "JACKSON", "ISAAC": "ISAAC"
};

const CONSULTORES_RECORRENCIA_RETENCAO = ['JACKSON', 'ISAAC'];
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

// ── CAMPANHA 14° SALÁRIO ────────────────────────────────────────────────────
const ID_PLANILHA_CAMPANHA14 = "1bSB3p5NeaA6CMNUJ9C43wXoI6uGMuAe-3w82LA5FnWQ";
const MESES_CAMPANHA         = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov"];
const META_MESES_CAMPANHA    = 9; // 9 de 11 meses = 80%
const SETORES_CAMPANHA       = [
  "Recepção","Refiliação","Homologação","Conciliação",
  "Adimplência","Vendas","Cashback",
  "Líder de Vendas","Líder de Adimplência","Líder de Conciliação","Coordenador"
];
// ───────────────────────────────────────────────────────────────────────────

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
    } else if (endpoint === "campanha14") {
      resultado = getCampanha14Data();
    } else {
      resultado = {
        status: "error",
        error: "Endpoint nao encontrado: " + endpoint,
        endpoints_disponiveis: ["documentacao", "app", "adimplencia", "recorrencia", "refuturiza", "recorrencia_vendedor", "campanha14"]
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
      .addItem("📊 Abrir Dashboard (TODOS OS DADOS)", "abrirDashboardRecorrenciaVendedor"))
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
// FUNÇÕES DE ABERTURA DE DASHBOARDS
// ============================================================================

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

function abrirDashboardRecorrenciaVendedor() {
  abrirDashboardPorMesETipo(0, "recorrencia_vendedor");
}

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
// FUNÇÕES PRINCIPAIS DE ABERTURA
// ============================================================================

function abrirDashboardPorMesETipo(mes, tipo) {
  try {
    const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
    let sheetName = "";
    let sheet = null;

    if (tipo === "documentacao") {
      sheetName = "Dashboard Documentação";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) { const ano = new Date().getFullYear(); getDocumentacaoData(mes, ano); sheet = ssDash.getSheetByName(sheetName); }
    } else if (tipo === "app") {
      sheetName = "Dashboard App";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) { const ano = new Date().getFullYear(); getAppData(mes, ano); sheet = ssDash.getSheetByName(sheetName); }
    } else if (tipo === "adimplencia") {
      sheetName = "Dashboard Adimplência";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) { const ano = new Date().getFullYear(); getAdimplenciaData(mes, ano); sheet = ssDash.getSheetByName(sheetName); }
    } else if (tipo === "recorrencia") {
      sheetName = "Dashboard Recorrência";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) { gerarDashboardRecorrencia(); sheet = ssDash.getSheetByName(sheetName); }
    } else if (tipo === "recorrencia_vendedor") {
      sheetName = "Dashboard Recorrência Vendedor";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) { getRecorrenciaVendedorData(); sheet = ssDash.getSheetByName(sheetName); }
    } else if (tipo === "refuturiza") {
      sheetName = "Dashboard Refuturiza";
      sheet = ssDash.getSheetByName(sheetName);
      if (!sheet) { const ano = new Date().getFullYear(); getRefuturizaData(mes, ano); sheet = ssDash.getSheetByName(sheetName); }
    }

    if (!sheet) throw new Error(`Aba ${sheetName} não encontrada`);

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
// GERAR DASHBOARD PARA MÊS PASSADO ESPECÍFICO
// ============================================================================

function gerarDashboardMesPassado() {
  try {
    const ui = SpreadsheetApp.getUi();
    const anoAtual = new Date().getFullYear();
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
            <option value="1">Janeiro</option><option value="2">Fevereiro</option>
            <option value="3">Março</option><option value="4">Abril</option>
            <option value="5">Maio</option><option value="6">Junho</option>
            <option value="7">Julho</option><option value="8">Agosto</option>
            <option value="9">Setembro</option><option value="10">Outubro</option>
            <option value="11">Novembro</option><option value="12">Dezembro</option>
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
            const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
            const dashboardNames = { 'documentacao':'Documentação','app':'App','adimplencia':'Adimplência','recorrencia':'Recorrência','recorrencia_vendedor':'Recorrência Vendedor','refuturiza':'Refuturiza' };
            document.getElementById('dashboardPreview').textContent = dashboardNames[dashboardType];
            document.getElementById('mesAnoPreview').textContent = dashboardType === 'recorrencia_vendedor' ? 'TODOS OS MESES (dados completos)' : meses[mes-1] + ' ' + ano;
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
              google.script.run.withSuccessHandler(function(result) { alert(result); google.script.host.close(); }).withFailureHandler(function(error) { alert('Erro: ' + error.message); }).processarDashboardEspecifico(dashboardType, 0, 0);
            } else {
              const mes = parseInt(document.getElementById('mes').value);
              const ano = parseInt(document.getElementById('ano').value);
              google.script.run.withSuccessHandler(function(result) { alert(result); google.script.host.close(); }).withFailureHandler(function(error) { alert('Erro: ' + error.message); }).processarDashboardEspecifico(dashboardType, mes, ano);
            }
          }
        </script>
      </div>
    `;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(500), 'Gerar Dashboard Específico');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Erro: ${error.message}`);
  }
}

function processarDashboardEspecifico(tipo, mes, ano) {
  try {
    let resultado;
    switch(tipo) {
      case 'documentacao': resultado = getDocumentacaoData(mes, ano); break;
      case 'app': resultado = getAppData(mes, ano); break;
      case 'adimplencia': resultado = getAdimplenciaData(mes, ano); break;
      case 'recorrencia':
        resultado = processarDadosRecorrenciaPorMes(mes, ano, true);
        if (resultado.status === "success") renderizarPlanilhaRecorrenciaPorMes(resultado.data, mes, ano);
        break;
      case 'recorrencia_vendedor': resultado = getRecorrenciaVendedorData(); break;
      case 'refuturiza': resultado = getRefuturizaData(mes, ano); break;
      default: return `❌ Tipo de dashboard desconhecido: ${tipo}`;
    }
    if (resultado && resultado.status === "success") {
      return tipo === 'recorrencia_vendedor'
        ? `✅ Dashboard Recorrência Vendedor (TODOS OS DADOS) gerado com sucesso!`
        : `✅ Dashboard ${tipo} de ${MESES[mes-1]} ${ano} gerado com sucesso!`;
    } else if (resultado && resultado.status === "error") {
      return `❌ Erro: ${resultado.error}`;
    } else {
      return tipo === 'recorrencia_vendedor'
        ? `✅ Dashboard Recorrência Vendedor processado com sucesso!`
        : `✅ Dashboard ${tipo} de ${MESES[mes-1]} ${ano} processado com sucesso!`;
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
  if (["SIM", "S", "YES", "Y", "COM APP", "BAIXADO", "INSTALADO", "✅", "OK", "CONCLUÍDO", "CONCLUIDO"].includes(statusUpper)) return "SIM";
  if (["NÃO", "NAO", "N", "NO", "SEM APP", "NÃO BAIXADO", "NÃO INSTALADO", "❌", "NEGATIVO", "NA"].includes(statusUpper)) return "NAO";
  if (["CANCELADO", "CANCEL", "C", "CANCELADA", "CANCELADOS", "CANC"].includes(statusUpper)) return "CANCELADO";
  return "OUTROS";
}

// ============================================================================
// FUNÇÃO AUXILIAR: NORMALIZAR STATUS DE LIGAÇÃO (REFUTURIZA)
// ============================================================================

function normalizarStatusLigacao(status) {
  if (!status || status === "") return "SEM LIGAÇÃO";
  const statusUpper = status.toString().toUpperCase().trim();
  if (["SIM", "S", "YES", "Y", "COM LIGAÇÃO", "LIGADO", "REALIZADO", "CONCLUÍDO", "CONCLUIDO", "✅", "OK", "FEITO", "FEITA"].includes(statusUpper)) return "COM LIGAÇÃO";
  if (["NÃO", "NAO", "N", "NO", "SEM LIGAÇÃO", "NÃO LIGADO", "NÃO REALIZADO", "PENDENTE", "❌", "NEGATIVO", "NA"].includes(statusUpper)) return "SEM LIGAÇÃO";
  if (["CANCELADO", "CANCEL", "C", "CANCELADA", "CANCELADOS", "CANC", "DESISTENTE", "DESISTIU"].includes(statusUpper)) return "CANCELADO";
  if (statusUpper.includes("LIG") || statusUpper.includes("CALL") || statusUpper.includes("TELEFONE")) return "COM LIGAÇÃO";
  return "SEM LIGAÇÃO";
}

// ============================================================================
// ENDPOINT: DOCUMENTAÇÃO — V21.0
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
      data:      headers.indexOf("DATA"),
      kyc:       headers.indexOf("KYC"),
      promocao:  headers.indexOf("PROMOÇÃO")
    };

    const mesAtual = mes || new Date().getMonth() + 1;
    const anoAtual = ano || new Date().getFullYear();

    function novoTotalizador() {
      return {
        total: 0, cancelados: 0, aprovados: 0,
        pendencias: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0,
        promo:  { total: 0, cancelados: 0, aprovados: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0 },
        normal: { total: 0, cancelados: 0, aprovados: 0, reprovados: 0, expirado: 0, pendente: 0, naoEnviado: 0 }
      };
    }

    const consultoresData = {};
    let totalGeral = novoTotalizador();
    let vendasLoja = novoTotalizador();
    let vendasWeb  = novoTotalizador();

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const dataParsed = parseDate(row[idx.data]);
      if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;

      const consultorOriginal    = String(row[idx.consultor] || "").trim();
      const consultorNormalizado = normalizarConsultorResidual(consultorOriginal);
      const kyc = String(row[idx.kyc] || "").trim().toUpperCase();

      let isPromo = false;
      if (idx.promocao !== -1) {
        const promoVal = String(row[idx.promocao] || "").trim().toUpperCase();
        isPromo = ["SIM", "S", "YES", "Y", "1", "TRUE"].includes(promoVal);
      }
      const bucket = isPromo ? "promo" : "normal";

      let setor = "TELEVENDAS";
      if (consultorNormalizado === "WEB SITE")    setor = "WEB SITE";
      else if (consultorNormalizado === "OUTROS") setor = "OUTROS";
      else if (consultorNormalizado === "TELEVENDAS") setor = "TELEVENDAS";
      else {
        for (const s in CONSULTORES_MAP) {
          if (CONSULTORES_MAP[s].some(c => consultorNormalizado.toUpperCase().includes(c.toUpperCase()))) { setor = s; break; }
        }
      }

      const isWebTelevendas = (consultorNormalizado === "WEB SITE" || consultorNormalizado === "TELEVENDAS" || consultorNormalizado === "OUTROS");
      const dest = isWebTelevendas ? vendasWeb : vendasLoja;

      if (!consultoresData[consultorNormalizado]) {
        consultoresData[consultorNormalizado] = Object.assign(novoTotalizador(), { nome: consultorNormalizado, setor: setor });
      }
      const c = consultoresData[consultorNormalizado];

      c.total++; totalGeral.total++; dest.total++;
      c[bucket].total++; totalGeral[bucket].total++; dest[bucket].total++;

      if (kyc === "CANCELADO") {
        c.cancelados++; totalGeral.cancelados++; dest.cancelados++;
        c[bucket].cancelados++; totalGeral[bucket].cancelados++; dest[bucket].cancelados++;
        continue;
      }

      if (kyc === "APROVADO") {
        c.aprovados++;           totalGeral.aprovados++;           dest.aprovados++;
        c[bucket].aprovados++;   totalGeral[bucket].aprovados++;   dest[bucket].aprovados++;
      } else {
        c.pendencias++; totalGeral.pendencias++; dest.pendencias++;
        if (kyc === "REPROVADO") {
          c.reprovados++;          totalGeral.reprovados++;          dest.reprovados++;
          c[bucket].reprovados++;  totalGeral[bucket].reprovados++;  dest[bucket].reprovados++;
        } else if (kyc === "EXPIRADO") {
          c.expirado++;            totalGeral.expirado++;            dest.expirado++;
          c[bucket].expirado++;    totalGeral[bucket].expirado++;    dest[bucket].expirado++;
        } else if (kyc === "PENDENTE") {
          c.pendente++;            totalGeral.pendente++;            dest.pendente++;
          c[bucket].pendente++;    totalGeral[bucket].pendente++;    dest[bucket].pendente++;
        } else {
          c.naoEnviado++;          totalGeral.naoEnviado++;          dest.naoEnviado++;
          c[bucket].naoEnviado++;  totalGeral[bucket].naoEnviado++;  dest[bucket].naoEnviado++;
        }
      }
    }

    const responseData = {
      mes: MESES[mesAtual - 1],
      ano: anoAtual,
      geral:      totalGeral,
      vendasLoja: vendasLoja,
      vendasWeb:  vendasWeb,
      consultores: Object.values(consultoresData),
      temColunasPromo: idx.promocao !== -1
    };

    criarDashboardDocumentacao(responseData);
    return createSuccessResponse(responseData);
  } catch (error) {
    return createErrorResponse(error.message);
  }
}

// ── Helper: % aprovação sobre base líquida (sem cancelados) ──────────────────
function calcPctDoc(aprovados, total, cancelados) {
  const base = total - cancelados;
  return base > 0 ? Math.round((aprovados / base) * 100) : 0;
}

// ============================================================================
// DASHBOARD DOCUMENTAÇÃO — V20.6.1 (LAYOUT REDESENHADO)
// ============================================================================

function criarDashboardDocumentacao(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Documentação");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Documentação");
  sheet.clear();

  // ── Título ─────────────────────────────────────────────────────────────────
  sheet.getRange("A1:L1").merge()
    .setValue("DASHBOARD DOCUMENTAÇÃO - " + dados.mes.toUpperCase() + " " + dados.ano)
    .setBackground("#000000").setFontColor("#ffffff")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  // ── 3 cards principais (Geral / Loja / Web) ────────────────────────────────
  renderizarTresCardsPrincipaisDoc(sheet, dados);

  // ── Seção comparativo Promoção vs Normal ───────────────────────────────────
  let currentRow = 22;
  if (dados.temColunasPromo) {
    currentRow = renderizarCardsPromoDoc(sheet, dados, currentRow);
  }

  // ── Cards por setor ────────────────────────────────────────────────────────
  const setores = ["VENDAS", "RECEPCAO", "REFILIACAO", "WEB SITE", "TELEVENDAS", "OUTROS"];

  setores.forEach(setor => {
    const consultoresSetor = dados.consultores.filter(c => c.setor === setor);
    if (consultoresSetor.length === 0) return;

    sheet.getRange(currentRow, 1, 1, 12).merge()
      .setValue("SETOR: " + setor)
      .setBackground("#4b5563").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("left");
    currentRow += 2;

    let col = 1;
    let maxCardHeight = 0;

    consultoresSetor.sort((a, b) => a.nome.localeCompare(b.nome)).forEach(c => {
      const p = calcPctDoc(c.aprovados, c.total, c.cancelados);

      // Só mostra seção de promoção quando o consultor realmente tem vendas promo
      const hasPromo = dados.temColunasPromo && c.promo.total > 0;

      const cardHeight = renderizarCardConsultorDoc(sheet, c, p, hasPromo, currentRow, col);
      maxCardHeight = Math.max(maxCardHeight, cardHeight);

      // Cada card ocupa 4 colunas de dados + 1 coluna de espaço = 5 por slot
      // 2 cards por linha: col 1 e col 6 (col 5 fica em branco como separador)
      col += 5;
      if (col > 9) {
        col = 1;
        currentRow += maxCardHeight + 3; // +3 linhas de respiro entre linhas de cards
        maxCardHeight = 0;
      }
    });

    if (col !== 1) currentRow += maxCardHeight + 3;
    currentRow += 3; // espaço extra entre setores
  });

  sheet.setColumnWidths(1, 12, 160);
  sheet.setFrozenRows(1);
}

// ============================================================================
// NOVA FUNÇÃO: renderizarCardConsultorDoc
//
// Layout quando hasPromo = true (4 colunas):
// ┌──────────────────────────────────────────────────────────┐  ← banner nome
// │ TOTAL         │  20  │ ▶ PROMOÇÃO (cabeçalho)           │
// │ CANCELADOS    │   0  │ Total        │  5                 │
// │ BASE LÍQUIDA  │  20  │ Cancelados   │  0                 │
// │ APROVADOS     │  16  │ Base líquida │  5                 │
// │ PENDÊNCIAS    │   4  │ Aprovados    │  4                 │
// │   NÃO ENVIADO │   3  │ Pendências   │  1                 │
// │   EXPIRADO    │   1  │ % Aprovados  │ 80%                │
// │   PENDENTE    │   0  ├──────────────────────────────────┤
// │   REPROVADO   │   0  │ ▶ NORMAL (cabeçalho)             │
// │ % APROVADOS   │  80% │ Total        │ 15                 │
// │               │      │ Cancelados   │  0                 │
// │               │      │ Base líquida │ 15                 │
// │               │      │ Aprovados    │ 12                 │
// │               │      │ Pendências   │  3                 │
// │               │      │ % Aprovados  │ 80%                │
// └──────────────────────────────────────────────────────────┘
//
// Layout quando hasPromo = false (2 colunas — sem coluna direita):
// ┌──────────────┐
// │ nome         │  ← banner 2 colunas
// │ TOTAL  │  20 │
// │ ...         │
// └──────────────┘
//
// Retorna a altura total do card (linhas utilizadas, incluindo o banner)
// ============================================================================

function renderizarCardConsultorDoc(sheet, c, p, hasPromo, row, col) {
  const cardWidth = hasPromo ? 4 : 2;

  // ── Banner do nome (abrange TODAS as colunas do card) ────────────────────
  sheet.getRange(row, col, 1, cardWidth).merge()
    .setValue(c.nome)
    .setBackground("#059669").setFontColor("#ffffff")
    .setFontWeight("bold").setHorizontalAlignment("center");

  // ── Dados principais (cols col e col+1) ─────────────────────────────────
  const mainData = [
    ["TOTAL",          c.total],
    ["CANCELADOS",     c.cancelados],
    ["BASE LÍQUIDA",   c.total - c.cancelados],
    ["APROVADOS",      c.aprovados],
    ["PENDÊNCIAS",     c.pendencias],
    ["  NÃO ENVIADO",  c.naoEnviado],
    ["  EXPIRADO",     c.expirado],
    ["  PENDENTE",     c.pendente],
    ["  REPROVADO",    c.reprovados],
    ["% APROVADOS",    p + "%"]
  ];

  sheet.getRange(row + 1, col, mainData.length, 2)
    .setValues(mainData).setBorder(true, true, true, true, true, true);

  // Cancelados — destaque vermelho suave
  sheet.getRange(row + 2, col + 1)
    .setBackground("#fee2e2").setFontColor("#991b1b").setFontWeight("bold");

  // % com cor de desempenho
  const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
  sheet.getRange(row + 10, col + 1)
    .setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");

  // Altura base: 1 (banner) + 10 (dados principais)
  let cardHeight = 11;

  // ── Seção Promoção / Normal (só quando hasPromo = true) ──────────────────
  if (hasPromo) {
    const pP = calcPctDoc(c.promo.aprovados, c.promo.total, c.promo.cancelados);
    const pN = calcPctDoc(c.normal.aprovados, c.normal.total, c.normal.cancelados);

    const promoData = [
      ["Total",        c.promo.total],
      ["Cancelados",   c.promo.cancelados],
      ["Base líquida", c.promo.total - c.promo.cancelados],
      ["Aprovados",    c.promo.aprovados],
      ["Pendências",   c.promo.reprovados + c.promo.expirado + c.promo.pendente + c.promo.naoEnviado],
      ["% Aprovados",  pP + "%"]
    ];

    const normalData = [
      ["Total",        c.normal.total],
      ["Cancelados",   c.normal.cancelados],
      ["Base líquida", c.normal.total - c.normal.cancelados],
      ["Aprovados",    c.normal.aprovados],
      ["Pendências",   c.normal.reprovados + c.normal.expirado + c.normal.pendente + c.normal.naoEnviado],
      ["% Aprovados",  pN + "%"]
    ];

    // Posições relativas ao row:
    //   row+1 : cabeçalho PROMOÇÃO (nas cols col+2 e col+3)
    //   row+2 a row+7 : dados promo (6 linhas)
    //   row+8 : cabeçalho NORMAL
    //   row+9 a row+14: dados normal (6 linhas)

    const promoHeaderRow  = row + 1;
    const promoDataRow    = row + 2;
    const normalHeaderRow = row + 2 + promoData.length;    // row + 8
    const normalDataRow   = normalHeaderRow + 1;           // row + 9

    // ── PROMOÇÃO ─────────────────────────────────────────────────────────
    sheet.getRange(promoHeaderRow, col + 2, 1, 2).merge()
      .setValue("▶ PROMOÇÃO")
      .setBackground("#7c3aed").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange(promoDataRow, col + 2, promoData.length, 2)
      .setValues(promoData).setBorder(true, true, true, true, true, true);

    const cP = pP >= 90 ? "#10b981" : pP >= 80 ? "#f59e0b" : "#ef4444";
    sheet.getRange(promoDataRow + promoData.length - 1, col + 3)
      .setBackground(cP).setFontColor("#ffffff").setFontWeight("bold");

    // ── NORMAL ────────────────────────────────────────────────────────────
    sheet.getRange(normalHeaderRow, col + 2, 1, 2).merge()
      .setValue("▶ NORMAL")
      .setBackground("#1e3a8a").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");

    sheet.getRange(normalDataRow, col + 2, normalData.length, 2)
      .setValues(normalData).setBorder(true, true, true, true, true, true);

    const cN = pN >= 90 ? "#10b981" : pN >= 80 ? "#f59e0b" : "#ef4444";
    sheet.getRange(normalDataRow + normalData.length - 1, col + 3)
      .setBackground(cN).setFontColor("#ffffff").setFontWeight("bold");

    // Altura com promo: banner(1) + promoHeader(1) + promoData(6) + normalHeader(1) + normalData(6) = 15
    const promoNormalHeight = 1 + promoData.length + 1 + normalData.length; // = 14
    cardHeight = Math.max(cardHeight, 1 + promoNormalHeight);               // = 15
  }

  return cardHeight;
}

// ── Cards principais (Geral / Loja / Web) para Documentação ──────────────────
function renderizarTresCardsPrincipaisDoc(sheet, dados) {
  const row = 3;
  const blocos = [
    { label: "TOTAL GERAL",       data: dados.geral,      col: 1, cor: "#000000" },
    { label: "VENDAS LOJA",       data: dados.vendasLoja, col: 4, cor: "#78350f" },
    { label: "WEB / TELEVENDAS",  data: dados.vendasWeb,  col: 7, cor: "#1e3a8a" }
  ];

  blocos.forEach(b => {
    const d  = b.data;
    const p  = calcPctDoc(d.aprovados, d.total, d.cancelados);
    const pP = dados.temColunasPromo ? calcPctDoc(d.promo.aprovados,  d.promo.total,  d.promo.cancelados)  : null;
    const pN = dados.temColunasPromo ? calcPctDoc(d.normal.aprovados, d.normal.total, d.normal.cancelados) : null;

    const card = [
      [b.label,         ""],
      ["TOTAL",         d.total],
      ["CANCELADOS",    d.cancelados],
      ["BASE LÍQUIDA",  d.total - d.cancelados],
      ["APROVADOS",     d.aprovados],
      ["PENDÊNCIAS",    d.pendencias],
      ["  NÃO ENVIADO", d.naoEnviado],
      ["  EXPIRADO",    d.expirado],
      ["  PENDENTE",    d.pendente],
      ["  REPROVADO",   d.reprovados],
      ["% APROVADOS",   p + "%"],
      ["", ""],
      ["── PROMOÇÃO ──", dados.temColunasPromo ? d.promo.total  + " vendas" : "sem coluna"],
      ["  Aprovados",   dados.temColunasPromo ? d.promo.aprovados  : "–"],
      ["  % Aprovados", dados.temColunasPromo ? pP + "%"           : "–"],
      ["── NORMAL ──",  dados.temColunasPromo ? d.normal.total + " vendas" : "sem coluna"],
      ["  Aprovados",   dados.temColunasPromo ? d.normal.aprovados : "–"],
      ["  % Aprovados", dados.temColunasPromo ? pN + "%"           : "–"]
    ];

    sheet.getRange(row, b.col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
    sheet.getRange(row, b.col, 1, 2).merge().setBackground(b.cor).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(row + 2, b.col + 1).setBackground("#fee2e2").setFontColor("#991b1b").setFontWeight("bold");
    const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
    sheet.getRange(row + 10, b.col + 1).setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");

    if (dados.temColunasPromo) {
      const cP = pP >= 90 ? "#10b981" : pP >= 80 ? "#f59e0b" : "#ef4444";
      const cN = pN >= 90 ? "#10b981" : pN >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(row + 12, b.col, 1, 2).setBackground("#f3e8ff").setFontWeight("bold");
      sheet.getRange(row + 15, b.col, 1, 2).setBackground("#e0f2fe").setFontWeight("bold");
      sheet.getRange(row + 14, b.col + 1).setBackground(cP).setFontColor("#ffffff").setFontWeight("bold");
      sheet.getRange(row + 17, b.col + 1).setBackground(cN).setFontColor("#ffffff").setFontWeight("bold");
    }
  });
}

// ── Seção comparativo Promoção vs Normal ─────────────────────────────────────
function renderizarCardsPromoDoc(sheet, dados, startRow) {
  sheet.getRange(startRow, 1, 1, 12).merge()
    .setValue("📊 COMPARATIVO: PROMOÇÃO vs NORMAL")
    .setBackground("#7c3aed").setFontColor("#ffffff")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(12);
  startRow++;

  const blocos = [
    { label: "PROMOÇÃO — GERAL",  data: dados.geral.promo,       col: 1,  cor: "#7c3aed" },
    { label: "PROMOÇÃO — LOJA",   data: dados.vendasLoja.promo,  col: 4,  cor: "#6d28d9" },
    { label: "NORMAL — GERAL",    data: dados.geral.normal,      col: 7,  cor: "#1e3a8a" },
    { label: "NORMAL — LOJA",     data: dados.vendasLoja.normal, col: 10, cor: "#1e40af" }
  ];

  blocos.forEach(b => {
    const d = b.data;
    const p = calcPctDoc(d.aprovados, d.total, d.cancelados);
    const card = [
      [b.label,         ""],
      ["TOTAL",         d.total],
      ["CANCELADOS",    d.cancelados],
      ["BASE LÍQUIDA",  d.total - d.cancelados],
      ["APROVADOS",     d.aprovados],
      ["REPROVADOS",    d.reprovados],
      ["EXPIRADO",      d.expirado],
      ["PENDENTE",      d.pendente],
      ["NÃO ENVIADO",   d.naoEnviado],
      ["% APROVADOS",   p + "%"]
    ];
    sheet.getRange(startRow, b.col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
    sheet.getRange(startRow, b.col, 1, 2).merge().setBackground(b.cor).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(startRow + 2, b.col + 1).setBackground("#fee2e2").setFontColor("#991b1b").setFontWeight("bold");
    const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
    sheet.getRange(startRow + 9, b.col + 1).setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");
  });

  return startRow + 12;
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
          if (CONSULTORES_MAP[s].some(c => consultorNormalizado.toUpperCase().includes(c.toUpperCase()))) { setor = s; break; }
        }
      }

      const isWebTelevendas = (consultorNormalizado === "WEB SITE" || consultorNormalizado === "TELEVENDAS" || consultorNormalizado === "OUTROS");

      if (!consultoresData[consultorNormalizado]) {
        consultoresData[consultorNormalizado] = { nome: consultorNormalizado, total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0, setor: setor, origem: "app" };
      }

      const c = consultoresData[consultorNormalizado];
      c.total++; totalGeral.total++;
      if (isWebTelevendas) appWeb.total++; else appLoja.total++;

      if (app === "SIM") { c.sim++; totalGeral.sim++; if (isWebTelevendas) appWeb.sim++; else appLoja.sim++; }
      else if (app === "NAO") { c.nao++; totalGeral.nao++; if (isWebTelevendas) appWeb.nao++; else appLoja.nao++; }
      else if (app === "CANCELADO") { c.cancelado++; totalGeral.cancelado++; if (isWebTelevendas) appWeb.cancelado++; else appLoja.cancelado++; }
      else { c.outros++; totalGeral.outros++; if (isWebTelevendas) appWeb.outros++; else appLoja.outros++; }
    }

    const retencaoData = {};
    const sheetRetencao = ssDados.getSheetByName("APP RETENÇÃO") || ssDados.getSheetByName("APP RETENCAO") || ssDados.getSheetByName("App Retenção");

    if (sheetRetencao) {
      const dataRetencao = sheetRetencao.getDataRange().getValues();
      const headersRetencao = dataRetencao[1];
      const idxRetencao = { consultor: headersRetencao.indexOf("CONSULTOR"), data: headersRetencao.indexOf("DATA"), app: headersRetencao.indexOf("APP BAIXADO") };

      if (idxRetencao.consultor !== -1 && idxRetencao.data !== -1 && idxRetencao.app !== -1) {
        for (let i = 2; i < dataRetencao.length; i++) {
          const row = dataRetencao[i];
          const dataParsed = parseDate(row[idxRetencao.data]);
          if (dataParsed.mes !== mesAtual || dataParsed.ano !== anoAtual) continue;
          const consultorOriginal = String(row[idxRetencao.consultor] || "").trim().toUpperCase();
          let consultora = null;
          for (const nome in CONSULTORAS_RETENCAO) {
            if (consultorOriginal.includes(nome.toUpperCase())) { consultora = nome; break; }
          }
          if (!consultora) continue;
          const appRaw = String(row[idxRetencao.app] || "").trim();
          const app = normalizarAppStatus(appRaw);
          if (!retencaoData[consultora]) {
            retencaoData[consultora] = { nome: consultora, total: 0, sim: 0, nao: 0, cancelado: 0, outros: 0, setor: "RETENÇÃO", origem: "retencao" };
          }
          const r = retencaoData[consultora];
          r.total++;
          if (app === "SIM") r.sim++;
          else if (app === "NAO") r.nao++;
          else if (app === "CANCELADO") r.cancelado++;
          else r.outros++;
        }
      }
    }

    const responseData = {
      mes: MESES[mesAtual - 1], ano: anoAtual,
      geral: totalGeral, appLoja: appLoja, appWeb: appWeb,
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
    sheet.getRange(currentRow, 1, 1, 12).merge().setValue("👑 CONSULTORAS DE RETENÇÃO (DADOS SEPARADOS)")
      .setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    currentRow += 2;
    let col = 1;
    dados.consultorasRetencao.sort((a, b) => a.nome.localeCompare(b.nome)).forEach(c => {
      const p = c.total > 0 ? Math.round((c.sim / c.total) * 100) : 0;
      const card = [[c.nome + " (RETENÇÃO)", ""], ["TOTAL", c.total], ["COM APP (SIM)", c.sim], ["SEM APP (NÃO)", c.nao], ["CANCELADO", c.cancelado], ["OUTROS", c.outros || 0], ["% COM APP", p + "%"], ["ORIGEM", c.origem === "retencao" ? "APP RETENÇÃO" : "APP"]];
      sheet.getRange(currentRow, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
      sheet.getRange(currentRow, col, 1, 2).merge().setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(currentRow + 6, col + 1).setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");
      col += 3;
      if (col > 9) { col = 1; currentRow += card.length + 1; }
    });
    if (col !== 1) currentRow += 9; else currentRow += 2;
    sheet.getRange(currentRow, 1, 1, 12).setBackground("#e5e7eb");
    currentRow += 2;
  }

  const setores = ["VENDAS", "RECEPCAO", "REFILIACAO", "WEB SITE", "TELEVENDAS", "OUTROS"];
  let currentRow = sheet.getLastRow() + 2;

  setores.forEach(setor => {
    const consultoresSetor = dados.consultores.filter(c => c.setor === setor && c.origem !== "retencao");
    if (consultoresSetor.length === 0) return;
    sheet.getRange(currentRow, 1, 1, 12).merge().setValue("SETOR: " + setor)
      .setBackground("#4b5563").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("left");
    currentRow += 2;
    let col = 1;
    consultoresSetor.sort((a, b) => a.nome.localeCompare(b.nome)).forEach(c => {
      const p = c.total > 0 ? Math.round((c.sim / c.total) * 100) : 0;
      const card = [[c.nome, ""], ["TOTAL", c.total], ["SIM", c.sim], ["NÃO", c.nao], ["CANCELADO", c.cancelado], ["OUTROS", c.outros || 0], ["% APP", p + "%"]];
      sheet.getRange(currentRow, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
      sheet.getRange(currentRow, col, 1, 2).merge().setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      const cCor = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(currentRow + 6, col + 1).setBackground(cCor).setFontColor("#ffffff").setFontWeight("bold");
      col += 3;
      if (col > 9) { col = 1; currentRow += 8; }
    });
    if (col !== 1) currentRow += 8; else currentRow += 2;
  });

  sheet.setColumnWidths(1, 12, 160);
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
      if (kyc === "APROVADO") { c.aprovados++; totalGeral.aprovados++; } else { c.pendentes++; totalGeral.pendentes++; }
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
  sheet.setColumnWidths(1, 12, 140);
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
      if (val instanceof Date) { d = val; } else { d = new Date(val); }
      if (isNaN(d.getTime())) return String(val).trim();
      const m = d.getMonth() + 1;
      const y = d.getFullYear();
      return (m < 10 ? '0' + m : m) + '/' + y;
    }

    const targetMonth = (mes < 10 ? '0' + mes : mes) + '/' + ano;
    let historicoMonths = [];

    if (isFixedMonth) {
      for (let i = 1; i <= 3; i++) {
        let histMes = mes - i; let histAno = ano;
        if (histMes <= 0) { histMes += 12; histAno -= 1; }
        historicoMonths.push((histMes < 10 ? '0' + histMes : histMes) + '/' + histAno);
      }
    } else {
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      for (let i = 1; i <= 3; i++) {
        let histMes = currentMonth - i; let histAno = currentYear;
        if (histMes <= 0) { histMes += 12; histAno -= 1; }
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
        if (mensalidade !== 'OK' && mensalidade !== 'EM ABERTO' && mensalidade !== 'CANCELADO') pendenciasMensalidade++;
        if (kyc !== 'APROVADO' && mensalidade !== 'CANCELADO') pendenciasKYC++;
      });

      if (isCurrentMonth) {
        const totalPendencias = pendenciasKYC;
        const totalRetidosFinal = totalRetido - cancelado;
        const retençõesOK = totalRetidosFinal - totalPendencias;
        const percentualOK = totalRetidosFinal > 0 ? Math.round((retençõesOK / totalRetidosFinal) * 100) : 0;
        return { totalRetido, ok, emAberto, emAtraso, cancelado, pendenciasKYC, totalPendencias, totalRetidosFinal, retençõesOK, percentualOK };
      } else {
        const totalOK = ok + emAberto;
        const totalPendencias = pendenciasMensalidade;
        const totalRetidosFinal = totalRetido - cancelado;
        const percentualOK = totalRetidosFinal > 0 ? Math.round((totalOK / totalRetidosFinal) * 100) : 0;
        return { totalRetido, ok, emAberto, totalOK, emAtraso, pendencias: totalPendencias, cancelado, totalRetidosFinal, percentualOK };
      }
    }

    const consultoresRetencao = CONSULTORES_RECORRENCIA_RETENCAO;
    const consultoresRefiliacao = CONSULTORES_RECORRENCIA_REFILIACAO;
    const result = { retencao: {}, refiliacao: {}, periodo: { atual: targetMonth, historico: historicoMonths, mes: mes, ano: ano, nomeMes: MESES[mes-1], isFixedMonth: isFixedMonth } };

    consultoresRetencao.forEach(c => {
      result.retencao[c] = {
        atual: calculateMetrics(c, [targetMonth], true),
        historico: historicoMonths.map(m => ({ mes: m, dados: calculateMetrics(c, [m]) })),
        total3Meses: calculateMetrics(c, historicoMonths)
      };
    });

    consultoresRefiliacao.forEach(c => {
      result.refiliacao[c] = {
        historico: historicoMonths.map(m => ({ mes: m, dados: calculateMetrics(c, [m]) })),
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
  return processarDadosRecorrenciaPorMes(now.getMonth() + 1, now.getFullYear(), false);
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
  const oldSheet = ssDash.getSheetByName("Dashboard Recorrência");
  if (oldSheet && oldSheet.getName() !== nomeAba) ssDash.deleteSheet(oldSheet);
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
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("CONSULTOR: " + c + " (RETENÇÃO)").setBackground("#f1f5f9").setFontWeight("bold");
    currentRow++;

    const cardAtual = [["MÊS ATUAL (" + dados.periodo.atual + ")", ""], ["TOTAL RETIDO", d.atual.totalRetido], ["CANCELADO", d.atual.cancelado], ["TOTAL RETIDOS FINAL", d.atual.totalRetidosFinal], ["MENSALIDADES OK", d.atual.ok], ["EM ABERTO", d.atual.emAberto], ["EM ATRASO", d.atual.emAtraso], ["PENDÊNCIAS KYC", d.atual.pendenciasKYC], ["TOTAL PENDÊNCIAS", d.atual.totalPendencias], ["RETENÇÕES OK", d.atual.retençõesOK], ["% OK", d.atual.percentualOK + "%"]];
    sheet.getRange(currentRow, 1, cardAtual.length, 2).setValues(cardAtual).setBorder(true, true, true, true, true, true);
    sheet.getRange(currentRow, 1, 1, 2).merge().setBackground("#10b981").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    const cardTotal = [["TOTAL 3 MESES ANTERIORES", ""], ["TOTAL RETIDO", d.total3Meses.totalRetido], ["CANCELADO", d.total3Meses.cancelado || 0], ["TOTAL RETIDOS FINAL", d.total3Meses.totalRetidosFinal], ["OK (Mensalidade OK)", d.total3Meses.ok], ["EM ABERTO", d.total3Meses.emAberto], ["TOTAL OK (OK + Aberto)", d.total3Meses.totalOK || (d.total3Meses.ok + d.total3Meses.emAberto)], ["EM ATRASO", d.total3Meses.emAtraso], ["TOTAL PENDÊNCIAS", d.total3Meses.pendencias], ["% OK (considera Aberto)", (d.total3Meses.percentualOK || 0) + "%"]];
    sheet.getRange(currentRow, 4, cardTotal.length, 2).setValues(cardTotal).setBorder(true, true, true, true, true, true);
    sheet.getRange(currentRow, 4, 1, 2).merge().setBackground("#f59e0b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    const percentualCell = sheet.getRange(currentRow + cardTotal.length - 1, 5);
    const corPercentual = d.total3Meses.percentualOK >= 90 ? "#10b981" : d.total3Meses.percentualOK >= 80 ? "#f59e0b" : "#ef4444";
    percentualCell.setBackground(corPercentual).setFontColor("#ffffff").setFontWeight("bold");

    currentRow += Math.max(cardAtual.length, cardTotal.length) + 2;
  });

  Object.keys(dados.refiliacao).forEach(c => {
    const d = dados.refiliacao[c];
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("CONSULTOR: " + c + " (REFILIAÇÃO)").setBackground("#f1f5f9").setFontWeight("bold");
    currentRow++;

    const cardTotal = [["TOTAL 3 MESES ANTERIORES", ""], ["TOTAL REFILIAÇÃO", d.total3Meses.totalRetido], ["CANCELADO", d.total3Meses.cancelado || 0], ["TOTAL REFILIADOS FINAL", d.total3Meses.totalRetidosFinal], ["OK (Mensalidade OK)", d.total3Meses.ok], ["EM ABERTO", d.total3Meses.emAberto], ["TOTAL OK (OK + Aberto)", d.total3Meses.totalOK || (d.total3Meses.ok + d.total3Meses.emAberto)], ["EM ATRASO", d.total3Meses.emAtraso], ["TOTAL PENDÊNCIAS", d.total3Meses.pendencias], ["% OK (considera Aberto)", (d.total3Meses.percentualOK || 0) + "%"]];
    sheet.getRange(currentRow, 1, cardTotal.length, 2).setValues(cardTotal).setBorder(true, true, true, true, true, true);
    sheet.getRange(currentRow, 1, 1, 2).merge().setBackground("#f59e0b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    const percentualCell = sheet.getRange(currentRow + cardTotal.length - 1, 2);
    const corPercentual = d.total3Meses.percentualOK >= 90 ? "#10b981" : d.total3Meses.percentualOK >= 80 ? "#f59e0b" : "#ef4444";
    percentualCell.setBackground(corPercentual).setFontColor("#ffffff").setFontWeight("bold");

    currentRow += cardTotal.length + 2;
  });

  sheet.setColumnWidths(1, 13, 150);
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
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("👑 " + c + " - RETENÇÃO").setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
    currentRow++;
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("MÊS ATUAL: " + dados.periodo.atual).setBackground("#10b981").setFontColor("#ffffff").setFontWeight("bold");
    currentRow++;
    sheet.getRange(currentRow, 1, 1, 2).setValues([["Métrica", "Valor"]]).setBackground("#d1fae5");
    currentRow++;
    const dadosAtual = [["Total Retido", d.atual.totalRetido], ["Cancelados", d.atual.cancelado], ["Total Retidos Final", d.atual.totalRetidosFinal], ["Mensalidades OK", d.atual.ok], ["Em Aberto", d.atual.emAberto], ["Em Atraso", d.atual.emAtraso], ["Pendências KYC", d.atual.pendenciasKYC], ["Total Pendências", d.atual.totalPendencias], ["Retenções OK", d.atual.retençõesOK], ["% OK", d.atual.percentualOK + "%"]];
    sheet.getRange(currentRow, 1, dadosAtual.length, 2).setValues(dadosAtual);
    currentRow += dadosAtual.length + 2;
    sheet.getRange(currentRow, 1, 1, 13).merge().setValue("HISTÓRICO - ÚLTIMOS 3 MESES").setBackground("#f59e0b").setFontColor("#ffffff").setFontWeight("bold");
    currentRow++;
    sheet.getRange(currentRow, 1, 1, 7).setValues([["Mês", "Total", "OK", "Em Aberto", "Total OK", "Pendências", "% OK"]]).setBackground("#fef3c7");
    currentRow++;
    d.historico.forEach(h => {
      sheet.getRange(currentRow, 1, 1, 7).setValues([[h.mes, h.dados.totalRetido, h.dados.ok, h.dados.emAberto, h.dados.totalOK || (h.dados.ok + h.dados.emAberto), h.dados.pendencias, h.dados.percentualOK + "%"]]);
      currentRow++;
    });
    currentRow += 2;
  });

  sheet.setColumnWidths(1, 13, 140);
}

// ============================================================================
// ENDPOINT: RECORRÊNCIA VENDEDOR (SEM FILTRO DE MÊS)
// ============================================================================

function getRecorrenciaVendedorData() {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    let sheet = ssDados.getSheetByName(NOME_ABA_RECORRENCIA_VENDAS);
    if (!sheet) {
      const sheets = ssDados.getSheets();
      sheet = sheets.find(s => {
        const nome = s.getName().toUpperCase();
        return nome.includes("RECORRENCIA VENDAS") || nome.includes("RECORRENCIA_VENDAS") || nome.includes("RECORRENCIA-VENDAS");
      });
    }
    if (!sheet) return createErrorResponse("Sheet 'RECORRENCIA VENDAS' não encontrada");

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return createErrorResponse("Planilha 'RECORRENCIA VENDAS' está vazia");

    const idx = { matricula: 0, filiado: 1, consultor: 2, data: 3, status: 4 };
    const consultoresData = {};
    let totalGeral = { totalVendasPromocao: 0, totalOk: 0, totalEmAberto: 0, totalAtraso: 0, totalOutros: 0 };
    const dataAtual = new Date();
    const mesAtualNumero = dataAtual.getMonth() + 1;
    const anoAtual = dataAtual.getFullYear();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idx.consultor] && !row[idx.matricula]) continue;
      const consultorOriginal = String(row[idx.consultor] || "").trim();
      if (!consultorOriginal) continue;
      const consultorNome = normalizarNomeConsultorRecorrencia(consultorOriginal);
      const setor = determinarSetorConsultor(consultorNome);
      if (!consultorEstaNoMapeamento(consultorNome)) continue;
      if (!consultoresData[consultorNome]) {
        consultoresData[consultorNome] = { nome: consultorNome, setor: setor, totalVendasPromocao: 0, totalOk: 0, totalEmAberto: 0, totalAtraso: 0, totalOutros: 0 };
      }
      const c = consultoresData[consultorNome];
      c.totalVendasPromocao++; totalGeral.totalVendasPromocao++;
      const status = String(row[idx.status] || "").trim().toUpperCase();
      if (status === "OK") { c.totalOk++; totalGeral.totalOk++; }
      else if (status.includes("ABERTO") || status === "EM ABERTO") { c.totalEmAberto++; totalGeral.totalEmAberto++; }
      else if (status.includes("ATRASO") || status === "EM ATRASO") { c.totalAtraso++; totalGeral.totalAtraso++; }
      else { c.totalOutros++; totalGeral.totalOutros++; }
    }

    Object.values(consultoresData).forEach(c => {
      c.percentualVendasOk = c.totalVendasPromocao > 0 ? Math.round(((c.totalOk + c.totalEmAberto) / c.totalVendasPromocao) * 100) : 0;
    });
    totalGeral.percentualVendasOk = totalGeral.totalVendasPromocao > 0 ?
      Math.round(((totalGeral.totalOk + totalGeral.totalEmAberto) / totalGeral.totalVendasPromocao) * 100) : 0;

    const dadosPorSetor = {};
    for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
      dadosPorSetor[setor] = Object.values(consultoresData).filter(c => c.setor === setor).sort((a, b) => b.totalVendasPromocao - a.totalVendasPromocao);
    }

    const responseData = {
      mes: "TODOS OS MESES", ano: "TODOS OS ANOS",
      mesAtual: MESES[mesAtualNumero - 1], anoAtual: anoAtual,
      geral: totalGeral,
      consultores: Object.values(consultoresData).sort((a, b) => b.totalVendasPromocao - a.totalVendasPromocao),
      dadosPorSetor: dadosPorSetor,
      totalConsultores: Object.keys(consultoresData).length,
      totalRegistros: data.length - 1
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
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    for (const consultor of CONSULTORES_RECORRENCIA_VENDAS[setor]) {
      const consultorUpper = consultor.toUpperCase();
      if (nomeUpper === consultorUpper || nomeUpper.includes(consultorUpper)) return consultor;
    }
  }
  const primeiroNome = nomeUpper.split(" ")[0];
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    for (const consultor of CONSULTORES_RECORRENCIA_VENDAS[setor]) {
      const primeiroNomeConsultor = consultor.toUpperCase().split(" ")[0];
      if (primeiroNome === primeiroNomeConsultor) return consultor;
    }
  }
  return nomeUpper;
}

function determinarSetorConsultor(consultorNome) {
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    if (CONSULTORES_RECORRENCIA_VENDAS[setor].includes(consultorNome)) return setor;
  }
  return "OUTROS";
}

function consultorEstaNoMapeamento(consultorNome) {
  for (const setor in CONSULTORES_RECORRENCIA_VENDAS) {
    if (CONSULTORES_RECORRENCIA_VENDAS[setor].includes(consultorNome)) return true;
  }
  return false;
}

function criarDashboardRecorrenciaVendedor(dados) {
  const ssDash = SpreadsheetApp.openById(ID_PLANILHA_DASHBOARDS);
  let sheet = ssDash.getSheetByName("Dashboard Recorrência Vendedor");
  if (!sheet) sheet = ssDash.insertSheet("Dashboard Recorrência Vendedor");
  sheet.clear();

  sheet.getRange("A1:L1").merge().setValue(`🤝 DASHBOARD RECORRÊNCIA VENDEDOR - DADOS COMPLETOS`)
    .setBackground("#0d9488").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  const g = dados.geral || { totalVendasPromocao: 0, totalOk: 0, totalEmAberto: 0, totalAtraso: 0, totalOutros: 0 };
  const percentualGeral = g.totalVendasPromocao > 0 ? Math.round(((g.totalOk + g.totalEmAberto) / g.totalVendasPromocao) * 100) : 0;

  const cardGeral = [["📈 TOTAL GERAL", ""], ["TOTAL VENDAS", g.totalVendasPromocao], ["OK", g.totalOk], ["EM ABERTO", g.totalEmAberto], ["ATRASO", g.totalAtraso], ["OUTROS", g.totalOutros], ["", ""], ["% OK+ABERTO", percentualGeral + "%"]];
  sheet.getRange(3, 1, cardGeral.length, 2).setValues(cardGeral).setBorder(true, true, true, true, true, true);
  sheet.getRange(3, 1, 1, 2).merge().setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  const corPercentual = percentualGeral >= 90 ? "#10b981" : percentualGeral >= 80 ? "#f59e0b" : "#ef4444";
  sheet.getRange(10, 2).setBackground(corPercentual).setFontColor("#ffffff").setFontWeight("bold");

  let currentRow = 14;
  const setores = [{ nome: "VENDAS", cor: "#1e3a8a" }, { nome: "REFILIACAO", cor: "#7c3aed" }, { nome: "RECEPCAO", cor: "#059669" }];

  setores.forEach(setorInfo => {
    const setor = setorInfo.nome;
    const corSetor = setorInfo.cor;
    const consultoresSetor = dados.consultores ? dados.consultores.filter(c => c.setor === setor) : [];

    sheet.getRange(currentRow, 1, 1, 12).merge()
      .setValue(`👥 SETOR: ${setor} (${consultoresSetor.length} consultor${consultoresSetor.length !== 1 ? 'es' : ''})`)
      .setBackground(corSetor).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    currentRow++;

    if (consultoresSetor.length === 0) {
      sheet.getRange(currentRow, 1, 1, 12).merge().setValue("ℹ️ Nenhum dado encontrado").setBackground("#f0f0f0").setFontColor("#666666").setHorizontalAlignment("center");
      currentRow += 2; return;
    }

    consultoresSetor.sort((a, b) => b.totalVendasPromocao - a.totalVendasPromocao);
    let col = 1; let consultorRow = currentRow; const CARD_HEIGHT = 9;

    consultoresSetor.forEach(consultor => {
      const percentualConsultor = consultor.totalVendasPromocao > 0 ? Math.round(((consultor.totalOk + consultor.totalEmAberto) / consultor.totalVendasPromocao) * 100) : 0;
      const cardConsultor = [[consultor.nome, ""], ["TOTAL VENDAS", consultor.totalVendasPromocao], ["OK", consultor.totalOk], ["EM ABERTO", consultor.totalEmAberto], ["ATRASO", consultor.totalAtraso], ["OUTROS", consultor.totalOutros || 0], ["", ""], ["% OK", percentualConsultor + "%"]];
      sheet.getRange(consultorRow, col, cardConsultor.length, 2).setValues(cardConsultor).setBorder(true, true, true, true, true, true);
      sheet.getRange(consultorRow, col, 1, 2).merge().setBackground(corSetor).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      const corPercentualConsultor = percentualConsultor >= 90 ? "#10b981" : percentualConsultor >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(consultorRow + 7, col + 1).setBackground(corPercentualConsultor).setFontColor("#ffffff").setFontWeight("bold");
      col += 3;
      if (col > 10) { col = 1; consultorRow += CARD_HEIGHT; }
    });

    currentRow = (col === 1) ? consultorRow : consultorRow + CARD_HEIGHT;
    currentRow += 2;
  });

  sheet.getRange(currentRow, 1, 1, 12).merge().setValue("📊 RESUMO FINAL").setBackground("#4b5563").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  currentRow++;
  const resumo = [["Total Consultores:", dados.totalConsultores || 0], ["Total Vendas:", g.totalVendasPromocao], ["Registros Processados:", dados.totalRegistros || 0], ["Média por Consultor:", dados.totalConsultores > 0 ? Math.round(g.totalVendasPromocao / dados.totalConsultores) : 0]];
  sheet.getRange(currentRow, 1, resumo.length, 2).setValues(resumo).setBorder(true, true, true, true, true, true);
  sheet.getRange(currentRow, 1, 1, 2).merge().setBackground("#0d9488").setFontColor("#ffffff").setFontWeight("bold");

  sheet.setColumnWidths(1, 12, 160);
  sheet.setFrozenRows(2);
}

// ============================================================================
// ENDPOINT: REFUTURIZA
// ============================================================================

function getRefuturizaData(mes, ano) {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    let sheet = ssDados.getSheetByName("REFUTURIZA");
    if (!sheet) {
      const sheets = ssDados.getSheets();
      sheet = sheets.find(s => { const nome = s.getName().toUpperCase(); return nome.includes("REFUTURIZA") || nome.includes("REFUTURISA") || nome.includes("REFUTUR"); });
    }
    if (!sheet) return createErrorResponse("Sheet 'REFUTURIZA' não encontrada");

    const data = sheet.getDataRange().getValues();
    const idx = { matricula: 0, nomeAderente: 1, nomeConsultor: 2, dataFiliacao: 3, statusLigacao: 4 };
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
      const status = normalizarStatusLigacao(String(row[idx.statusLigacao] || "").trim());

      if (!consultoresData[consultorOriginal]) {
        consultoresData[consultorOriginal] = { nome: consultorOriginal, total: 0, comLigacao: 0, semLigacao: 0, cancelado: 0 };
      }
      const c = consultoresData[consultorOriginal];
      c.total++; totalGeral.total++;
      if (status === "COM LIGAÇÃO") { c.comLigacao++; totalGeral.comLigacao++; }
      else if (status === "SEM LIGAÇÃO") { c.semLigacao++; totalGeral.semLigacao++; }
      else if (status === "CANCELADO") { c.cancelado++; totalGeral.cancelado++; }
      else { c.semLigacao++; totalGeral.semLigacao++; }
    }

    const consultoresComVendas = Object.values(consultoresData).filter(c => c.total > 0).sort((a, b) => b.total !== a.total ? b.total - a.total : a.nome.localeCompare(b.nome));
    const responseData = { mes: MESES[mesAtual - 1], ano: anoAtual, geral: totalGeral, consultores: consultoresComVendas };
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
    .setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(14);

  const row = 3;
  const g = dados.geral;
  const p = g.total > 0 ? Math.round((g.comLigacao / g.total) * 100) : 0;
  const cardGeral = [["REFUTURIZA - TOTAL DA LOJA", ""], ["TOTAL", g.total], ["COM LIGAÇÃO", g.comLigacao], ["SEM LIGAÇÃO", g.semLigacao], ["CANCELADO", g.cancelado], ["", ""], ["% COM LIGAÇÃO", p + "%"]];
  sheet.getRange(row, 1, cardGeral.length, 2).setValues(cardGeral).setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 1, 1, 2).merge().setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  const corPercentual = p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444";
  sheet.getRange(row + 6, 2).setBackground(corPercentual).setFontColor("#ffffff").setFontWeight("bold");

  let currentRow = row + cardGeral.length + 3;

  if (dados.consultores.length > 0) {
    sheet.getRange(currentRow, 1, 1, 12).merge().setValue("CONSULTORES COM VENDAS").setBackground("#4b5563").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    currentRow += 2;
    let col = 1; let consultorRow = currentRow;
    dados.consultores.forEach(c => {
      const pConsultor = c.total > 0 ? Math.round((c.comLigacao / c.total) * 100) : 0;
      const card = [[c.nome, ""], ["TOTAL", c.total], ["COM LIGAÇÃO", c.comLigacao], ["SEM LIGAÇÃO", c.semLigacao], ["CANCELADO", c.cancelado], ["", ""], ["% COM LIGAÇÃO", pConsultor + "%"]];
      sheet.getRange(consultorRow, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
      sheet.getRange(consultorRow, col, 1, 2).merge().setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
      const corPercentualConsultor = pConsultor >= 90 ? "#10b981" : pConsultor >= 80 ? "#f59e0b" : "#ef4444";
      sheet.getRange(consultorRow + 6, col + 1).setBackground(corPercentualConsultor).setFontColor("#ffffff").setFontWeight("bold");
      col += 3;
      if (col > 10) { col = 1; consultorRow += card.length + 1; }
    });
    if (col !== 1) currentRow = consultorRow + 8; else currentRow = consultorRow + 1;
  } else {
    sheet.getRange(currentRow, 1, 1, 12).merge().setValue("ℹ️ Nenhum consultor com vendas neste período").setBackground("#f59e0b").setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center");
    currentRow += 2;
  }

  currentRow += 2;
  const resumoFinal = [["📈 RESUMO FINAL", ""], ["Total de Consultores:", dados.consultores.length], ["Total de Vendas:", g.total], ["Média por Consultor:", dados.consultores.length > 0 ? Math.round(g.total / dados.consultores.length) : 0], ["Melhor %:", dados.consultores.length > 0 ? Math.max(...dados.consultores.map(c => c.total > 0 ? Math.round((c.comLigacao / c.total) * 100) : 0)) + "%" : "0%"], ["Pior %:", dados.consultores.length > 0 ? Math.min(...dados.consultores.map(c => c.total > 0 ? Math.round((c.comLigacao / c.total) * 100) : 0)) + "%" : "0%"]];
  sheet.getRange(currentRow, 1, resumoFinal.length, 2).setValues(resumoFinal).setBorder(true, true, true, true, true, true);
  sheet.getRange(currentRow, 1, 1, 2).merge().setBackground("#000000").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  sheet.setColumnWidths(1, 12, 160);
  sheet.setFrozenRows(1);

  try {
    const ultimaLinha = sheet.getLastRow();
    const rangeA1 = sheet.getRange("A1");
    if (rangeA1.isPartOfMerge()) { sheet.getRange(2, 1, ultimaLinha - 1, 12).createFilter(); }
    else { sheet.getRange(1, 1, ultimaLinha, 12).createFilter(); }
  } catch (e) { console.log("Não foi possível criar filtro: " + e.message); }
}

// ============================================================================
// FUNÇÕES AUXILIARES
// ============================================================================

function normalizarConsultorResidual(nome) {
  if (!nome) return "TELEVENDAS";
  let n = nome.toUpperCase();
  if (n.includes("WEB SITE")) return "WEB SITE";
  for (const consultora in CONSULTORAS_RETENCAO) {
    if (n.includes(consultora.toUpperCase())) return "OUTROS";
  }
  if (OUTROS_LISTA.some(o => n.includes(o.toUpperCase()))) return "OUTROS";
  let nomeLimpo = n.replace(PREFIXO_AMOR, "").trim();
  let primeiroNome = nomeLimpo.split(" ")[0];
  for (const setor in CONSULTORES_MAP) {
    const encontrado = CONSULTORES_MAP[setor].find(c => nomeLimpo === c.toUpperCase() || primeiroNome === c.toUpperCase().split(" ")[0]);
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
  const w = tipo === "DOC" ? dados.vendasWeb  : dados.appWeb;
  const pG = g.total > 0 ? Math.round((tipo === "DOC" ? (g.aprovados || 0) : (g.sim || 0)) / g.total * 100) : 0;
  const pL = l.total > 0 ? Math.round((tipo === "DOC" ? (l.aprovados || 0) : (l.sim || 0)) / l.total * 100) : 0;
  const pW = w.total > 0 ? Math.round((tipo === "DOC" ? (w.aprovados || 0) : (w.sim || 0)) / w.total * 100) : 0;
  const labels = tipo === "DOC"
    ? ["TOTAL", "APROVADOS", "PENDÊNCIAS", "  NÃO ENVIADO", "  EXPIRADO", "  PENDENTE", "% APROVADOS"]
    : ["TOTAL", "COM APP (SIM)", "SEM APP (NÃO)", "CANCELADO", "OUTROS", "", "% COM APP"];

  function buildCard(titulo, src, p) {
    return [[titulo, ""], ...labels.map((lab, i) => {
      if (!lab) return ["", ""];
      if (i === 0) return [lab, src.total || 0];
      if (tipo === "DOC") {
        if (i === 1) return [lab, src.aprovados || 0];
        if (i === 2) return [lab, src.pendencias || 0];
        if (i === 3) return [lab, src.naoEnviado || 0];
        if (i === 4) return [lab, src.expirado || 0];
        if (i === 5) return [lab, src.pendente || 0];
        return [lab, p + "%"];
      } else {
        if (i === 1) return [lab, src.sim || 0];
        if (i === 2) return [lab, src.nao || 0];
        if (i === 3) return [lab, src.cancelado || 0];
        if (i === 4) return [lab, src.outros || 0];
        return [lab, p + "%"];
      }
    })];
  }

  const card1 = buildCard(tipo === "DOC" ? "VENDAS TOTAIS" : "APP - TOTAL", g, pG);
  const card2 = buildCard(tipo === "DOC" ? "VENDAS LOJA"  : "APP - LOJA",  l, pL);
  const card3 = buildCard(tipo === "DOC" ? "VENDAS WEB/TELEVENDAS" : "APP - WEB/TELEVENDAS", w, pW);

  const percentRow = row + 7;
  [[card1, 1, "#000000", pG], [card2, 4, "#78350f", pL], [card3, 7, "#1e3a8a", pW]].forEach(([card, col, cor, p]) => {
    sheet.getRange(row, col, card.length, 2).setValues(card).setBorder(true, true, true, true, true, true);
    sheet.getRange(row, col, 1, 2).merge().setBackground(cor).setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(percentRow, col + 1).setBackground(p >= 90 ? "#10b981" : p >= 80 ? "#f59e0b" : "#ef4444").setFontColor("#ffffff").setFontWeight("bold");
  });
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
    if (partesBarra.length >= 3) { ano = parseInt(partesBarra[2], 10); if (ano < 100) ano += 2000; }
    else { ano = new Date().getFullYear(); }
    if (!isNaN(mes) && !isNaN(ano)) return { mes: mes, ano: ano };
  }
  const partesMesAno = str.split("/");
  if (partesMesAno.length === 2) {
    const mes = parseInt(partesMesAno[0], 10);
    const ano = parseInt(partesMesAno[1], 10);
    if (!isNaN(mes) && !isNaN(ano)) return { mes: mes, ano: ano < 100 ? ano + 2000 : ano };
  }
  const numeros = str.match(/\d+/g);
  if (numeros && numeros.length >= 2) {
    const mes = parseInt(numeros[1], 10);
    const ano = numeros.length >= 3 ? parseInt(numeros[2], 10) : new Date().getFullYear();
    if (!isNaN(mes) && mes >= 1 && mes <= 12) return { mes: mes, ano: ano < 100 ? ano + 2000 : ano };
  }
  return { mes: 0, ano: 0 };
}

function createSuccessResponse(data) { return { status: "success", timestamp: new Date().toISOString(), data: data }; }
function createErrorResponse(msg) { return { status: "error", timestamp: new Date().toISOString(), error: msg }; }

function atualizarTodosDashboards() {
  const mes = new Date().getMonth() + 1;
  const ano = new Date().getFullYear();
  // Flush entre cada dashboard evita timeout e libera memória
  getDocumentacaoData(mes, ano);  SpreadsheetApp.flush();
  getAppData(mes, ano);           SpreadsheetApp.flush();
  getAdimplenciaData(mes, ano);   SpreadsheetApp.flush();
  getRefuturizaData(mes, ano);    SpreadsheetApp.flush();
  getRecorrenciaVendedorData();   SpreadsheetApp.flush();
  gerarDashboardRecorrencia();
  SpreadsheetApp.getUi().alert("✅ Todos os 6 dashboards foram atualizados!");
}

// ============================================================================
// FUNÇÕES ADICIONAIS
// ============================================================================

function showConfigDialog() {
  const html = `
    <div style="padding: 20px; font-family: Arial;">
      <h2>⚙️ Configurações V20.6.2</h2>
      <p><strong>🎨 ATUALIZAÇÃO V20.6.2 — DOCUMENTAÇÃO:</strong></p>
      <ul>
        <li>Banner do nome abrange as 4 colunas do card (fica claro quem é o dono)</li>
        <li>PROMOÇÃO e NORMAL ficam à direita, alinhados ao consultor</li>
        <li>Seção Promoção/Normal só aparece se o consultor tiver vendas promo</li>
        <li><strong>Espaço de 1 coluna entre cards (col 5 em branco = separador visual)</strong></li>
        <li><strong>2 cards por linha com mais respiro entre linhas e entre setores</strong></li>
      </ul>
      <p><strong>📊 Dashboards Disponíveis (6):</strong></p>
      <ul>
        <li>📁 Documentação (por mês)</li>
        <li>📱 App (por mês)</li>
        <li>💳 Adimplência (por mês)</li>
        <li>📈 Recorrência (por mês fixo)</li>
        <li>🤝 Recorrência Vendedor (TODOS OS DADOS)</li>
        <li>🔄 Refuturiza (por mês)</li>
      </ul>
      <button onclick="google.script.host.close()" style="background: #10b981; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">Fechar</button>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(500).setHeight(450), 'Configurações V20.6.1');
}

function deployAPI() {
  const webAppUrl = ScriptApp.getService().getUrl();
  const html = `
    <div style="padding: 20px; font-family: Arial;">
      <h2>🚀 API V20.6 Implantada!</h2>
      <p>Sua API está disponível em:</p>
      <div style="background: #f0f0f0; padding: 10px; border-radius: 5px; margin: 10px 0;"><code>${webAppUrl}</code></div>
      <p><strong>📋 Exemplos de uso (6 endpoints):</strong></p>
      <div style="background: #f8fafc; padding: 10px; border-radius: 5px; font-family: monospace;">
        <div>${webAppUrl}?endpoint=documentacao&mes=1&ano=2024</div>
        <div>${webAppUrl}?endpoint=app&mes=2&ano=2024</div>
        <div>${webAppUrl}?endpoint=adimplencia&mes=3&ano=2024</div>
        <div>${webAppUrl}?endpoint=recorrencia&mes=4&ano=2024</div>
        <div>${webAppUrl}?endpoint=recorrencia_vendedor</div>
        <div>${webAppUrl}?endpoint=refuturiza&mes=6&ano=2024</div>
      </div>
      <button onclick="google.script.host.close()" style="background: #10b981; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">Fechar</button>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(650).setHeight(400), 'API Implantada V20.6');
}

function getDashboardDataForHTML(tipo, mes, ano) {
  try {
    let resultado;
    switch(tipo) {
      case 'documentacao': resultado = getDocumentacaoData(mes, ano); break;
      case 'app': resultado = getAppData(mes, ano); break;
      case 'adimplencia': resultado = getAdimplenciaData(mes, ano); break;
      case 'recorrencia': resultado = processarDadosRecorrenciaPorMes(mes, ano, true); break;
      case 'recorrencia_vendedor': resultado = getRecorrenciaVendedorData(); break;
      case 'refuturiza': resultado = getRefuturizaData(mes, ano); break;
      default: return { status: "error", error: "Tipo de dashboard desconhecido" };
    }
    if (resultado.status === "success" && resultado.data) {
      resultado.data.tipo = tipo; resultado.data.mesNumero = mes; resultado.data.anoNumero = ano;
    }
    return resultado;
  } catch (error) { return { status: "error", error: error.message }; }
}

// ============================================================================
// TESTES
// ============================================================================

function testarTodosDashboardsMesesPassados() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ano = new Date().getFullYear();
    let resultados = [];
    for (let mes = 1; mes <= 3; mes++) { resultados.push(`Documentação ${MESES[mes-1]}: ${getDocumentacaoData(mes, ano).status === "success" ? "✅" : "❌"}`); }
    for (let mes = 1; mes <= 3; mes++) { resultados.push(`App ${MESES[mes-1]}: ${getAppData(mes, ano).status === "success" ? "✅" : "❌"}`); }
    for (let mes = 1; mes <= 3; mes++) { resultados.push(`Adimplência ${MESES[mes-1]}: ${getAdimplenciaData(mes, ano).status === "success" ? "✅" : "❌"}`); }
    for (let mes = 1; mes <= 3; mes++) { resultados.push(`Recorrência ${MESES[mes-1]}: ${processarDadosRecorrenciaPorMes(mes, ano, true).status === "success" ? "✅" : "❌"}`); }
    resultados.push(`Recorrência Vendedor (TODOS OS DADOS): ${getRecorrenciaVendedorData().status === "success" ? "✅" : "❌"}`);
    for (let mes = 1; mes <= 3; mes++) { resultados.push(`Refuturiza ${MESES[mes-1]}: ${getRefuturizaData(mes, ano).status === "success" ? "✅" : "❌"}`); }
    ui.alert("🧪 TESTE DE TODOS OS DASHBOARDS\n\n" + resultados.join("\n"));
  } catch (error) { SpreadsheetApp.getUi().alert("❌ Erro no teste: " + error.message); }
}

function testarRecorrenciaVendedor() {
  try {
    const ui = SpreadsheetApp.getUi();
    const resultado = getRecorrenciaVendedorData();
    if (resultado.status === "success") {
      const dados = resultado.data;
      let mensagem = `🧪 TESTE RECORRÊNCIA VENDEDOR (DADOS COMPLETOS)\n\n✅ Dashboard gerado!\n\n`;
      mensagem += `Total Vendas: ${dados.geral.totalVendasPromocao}\nOK: ${dados.geral.totalOk}\nAberto: ${dados.geral.totalEmAberto}\n`;
      mensagem += `Atraso: ${dados.geral.totalAtraso}\nOutros: ${dados.geral.totalOutros || 0}\n% OK: ${dados.geral.percentualVendasOk}%\n\n`;
      mensagem += `Consultores: ${dados.totalConsultores}\nRegistros: ${dados.totalRegistros}\n`;
      if (dados.consultores.length > 0) {
        mensagem += `\n🏆 Top 3:\n`;
        dados.consultores.slice(0, 3).forEach((c, i) => { mensagem += `${i+1}. ${c.nome}: ${c.totalVendasPromocao} vendas (${c.percentualVendasOk}% OK)\n`; });
      }
      ui.alert(mensagem);
    } else { ui.alert(`❌ Erro: ${resultado.error}`); }
  } catch (error) { SpreadsheetApp.getUi().alert("❌ Erro no teste: " + error.message); }
}

function debugRecorrenciaVendedor() {
  try {
    const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
    let sheet = ssDados.getSheetByName(NOME_ABA_RECORRENCIA_VENDAS);
    if (!sheet) {
      const sheets = ssDados.getSheets();
      sheet = sheets.find(s => { const nome = s.getName().toUpperCase(); return nome.includes("RECORRENCIA VENDAS") || nome.includes("RECORRENCIA_VENDAS") || nome.includes("RECORRENCIA-VENDAS"); });
    }
    if (!sheet) { SpreadsheetApp.getUi().alert("Sheet 'RECORRENCIA VENDAS' não encontrada"); return; }
    const data = sheet.getDataRange().getValues();
    const consultoresUnicos = new Set();
    const consultoresCompletos = {};
    for (let i = 1; i < Math.min(data.length, 100); i++) {
      const row = data[i];
      const consultorOriginal = String(row[2] || "").trim();
      if (consultorOriginal) {
        consultoresUnicos.add(consultorOriginal);
        const consultorNormalizado = normalizarNomeConsultorRecorrencia(consultorOriginal);
        const setor = determinarSetorConsultor(consultorNormalizado);
        if (!consultoresCompletos[consultorNormalizado]) {
          consultoresCompletos[consultorNormalizado] = { original: consultorOriginal, normalizado: consultorNormalizado, setor: setor, noMapeamento: consultorEstaNoMapeamento(consultorNormalizado) };
        }
      }
    }
    let mensagem = "🔍 DEBUG RECORRÊNCIA VENDEDOR\n\n";
    mensagem += `Total únicos: ${consultoresUnicos.size}\n\n📋 CONSULTORES:\n`;
    for (const [normalizado, info] of Object.entries(consultoresCompletos)) {
      mensagem += `${info.noMapeamento ? "✅" : "❌"} "${info.original}" → "${normalizado}" (${info.setor})\n`;
    }
    SpreadsheetApp.getUi().alert(mensagem);
  } catch (error) { SpreadsheetApp.getUi().alert("Erro no debug: " + error.message); }
}

// ============================================================================
// SYNC SUPABASE
// ============================================================================

const SUPABASE_URL  = 'https://vycjtmjvkwvxunxtkdyi.supabase.co';
const SUPABASE_KEY  = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ5Y2p0bWp2a3d2eHVueHRrZHlpIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3MTYwNjY5NiwiZXhwIjoyMDg3MTgyNjk2fQ.sT4DiipgLk4CfZqa-FP1bATcb6wreg0b_usNXkyBpbk';

function configurarTriggerEdicao() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if (t.getHandlerFunction() === 'onEdicaoPlanilha') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEdicaoPlanilha').forSpreadsheet(ID_PLANILHA_DADOS).onEdit().create();
  SpreadsheetApp.getUi().alert('✅ Trigger instalado!\n\nAgora toda edição na planilha de dados vai sincronizar automaticamente com o Supabase.\n\nDelay estimado: 2-5 segundos após salvar.');
}

function onEdicaoPlanilha(e) {
  if (!e || !e.range) return;
  const abaEditada = e.range.getSheet().getName().toLowerCase();
  const mes = new Date().getMonth() + 1;
  const ano = new Date().getFullYear();
  if (abaEditada.includes('documenta'))                        syncDocumentacao(mes, ano);
  else if (abaEditada.includes('app retencao') || abaEditada.includes('app retenção')) syncApp(mes, ano);
  else if (abaEditada === 'app')                               syncApp(mes, ano);
  else if (abaEditada.includes('adimpl'))                      syncAdimplencia(mes, ano);
  else if (abaEditada === 'recorrencia')                       syncRecorrencia(mes, ano);
  else if (abaEditada.includes('recorrencia vendas'))          syncRecorrenciaVendedor();
  else if (abaEditada.includes('refuturiza'))                  syncRefuturiza(mes, ano);
}

function syncDocumentacao(mes, ano) {
  try {
    const resultado = getDocumentacaoData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;
    supabaseUpsert('documentacao', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total: d.geral.total, geral_cancelados: d.geral.cancelados, geral_aprovados: d.geral.aprovados,
      geral_pendencias: d.geral.pendencias, geral_reprovados: d.geral.reprovados, geral_expirado: d.geral.expirado,
      geral_pendente: d.geral.pendente, geral_nao_enviado: d.geral.naoEnviado,
      loja_total: d.vendasLoja.total, loja_cancelados: d.vendasLoja.cancelados, loja_aprovados: d.vendasLoja.aprovados,
      loja_pendencias: d.vendasLoja.pendencias, loja_reprovados: d.vendasLoja.reprovados, loja_expirado: d.vendasLoja.expirado,
      loja_pendente: d.vendasLoja.pendente, loja_nao_enviado: d.vendasLoja.naoEnviado,
      web_total: d.vendasWeb.total, web_cancelados: d.vendasWeb.cancelados, web_aprovados: d.vendasWeb.aprovados,
      web_pendencias: d.vendasWeb.pendencias, web_reprovados: d.vendasWeb.reprovados, web_expirado: d.vendasWeb.expirado,
      web_pendente: d.vendasWeb.pendente, web_nao_enviado: d.vendasWeb.naoEnviado,
      tem_promocao: d.temColunasPromo, geral_promo: JSON.stringify(d.geral.promo), geral_normal: JSON.stringify(d.geral.normal),
      loja_promo: JSON.stringify(d.vendasLoja.promo), loja_normal: JSON.stringify(d.vendasLoja.normal),
      consultores: JSON.stringify(d.consultores), atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] documentacao ' + mes + '/' + ano + ' OK');
  } catch(err) { console.error('[sync] documentacao ERRO: ' + err.message); }
}

function syncApp(mes, ano) {
  try {
    const resultado = getAppData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;
    supabaseUpsert('app_dashboard', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total: d.geral.total, geral_sim: d.geral.sim, geral_nao: d.geral.nao, geral_cancelado: d.geral.cancelado, geral_outros: d.geral.outros || 0,
      loja_total: d.appLoja.total, loja_sim: d.appLoja.sim, loja_nao: d.appLoja.nao, loja_cancelado: d.appLoja.cancelado, loja_outros: d.appLoja.outros || 0,
      web_total: d.appWeb.total, web_sim: d.appWeb.sim, web_nao: d.appWeb.nao, web_cancelado: d.appWeb.cancelado, web_outros: d.appWeb.outros || 0,
      consultores: JSON.stringify(d.consultores), consultoras_retencao: JSON.stringify(d.consultorasRetencao || []),
      atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] app ' + mes + '/' + ano + ' OK');
  } catch(err) { console.error('[sync] app ERRO: ' + err.message); }
}

function syncAdimplencia(mes, ano) {
  try {
    const resultado = getAdimplenciaData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data; const g = d.geral;
    supabaseUpsert('adimplencia', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total_trocas: g.totalTrocas, geral_mens_ok: g.mensOk, geral_mens_aberto: g.mensAberto, geral_mens_atraso: g.mensAtraso,
      geral_aprovados: g.aprovados, geral_pendentes: g.pendentes, geral_total_bi: g.totalBi, geral_fora_bi: g.foraBi, geral_ok_bi: g.okBi,
      geral_percentual_aprovado: g.percentualAprovado, consultores: JSON.stringify(d.consultores), atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] adimplencia ' + mes + '/' + ano + ' OK');
  } catch(err) { console.error('[sync] adimplencia ERRO: ' + err.message); }
}

function syncRecorrencia(mes, ano) {
  try {
    const resultado = processarDadosRecorrenciaPorMes(mes, ano, true);
    if (resultado.status !== 'success') return;
    supabaseUpsert('recorrencia', { mes: mes, ano: ano, mes_nome: MESES[mes - 1], dados: JSON.stringify(resultado.data), atualizado_em: new Date().toISOString() }, 'mes,ano');
    console.log('[sync] recorrencia ' + mes + '/' + ano + ' OK');
  } catch(err) { console.error('[sync] recorrencia ERRO: ' + err.message); }
}

function syncRecorrenciaVendedor() {
  try {
    const resultado = getRecorrenciaVendedorData();
    if (resultado.status !== 'success') return;
    const d = resultado.data; const g = d.geral;
    supabaseUpdate('recorrencia_vendedor', {
      geral_total_vendas: g.totalVendasPromocao, geral_total_ok: g.totalOk, geral_total_em_aberto: g.totalEmAberto,
      geral_total_atraso: g.totalAtraso, geral_total_outros: g.totalOutros || 0, geral_percentual_ok: g.percentualVendasOk,
      consultores: JSON.stringify(d.consultores), dados_por_setor: JSON.stringify(d.dadosPorSetor),
      total_consultores: d.totalConsultores, total_registros: d.totalRegistros, atualizado_em: new Date().toISOString()
    }, 'id=eq.1');
    console.log('[sync] recorrencia_vendedor OK');
  } catch(err) { console.error('[sync] recorrencia_vendedor ERRO: ' + err.message); }
}

function syncRefuturiza(mes, ano) {
  try {
    const resultado = getRefuturizaData(mes, ano);
    if (resultado.status !== 'success') return;
    const d = resultado.data;
    supabaseUpsert('refuturiza', {
      mes: mes, ano: ano, mes_nome: d.mes,
      geral_total: d.geral.total, geral_com_ligacao: d.geral.comLigacao, geral_sem_ligacao: d.geral.semLigacao, geral_cancelado: d.geral.cancelado,
      consultores: JSON.stringify(d.consultores), atualizado_em: new Date().toISOString()
    }, 'mes,ano');
    console.log('[sync] refuturiza ' + mes + '/' + ano + ' OK');
  } catch(err) { console.error('[sync] refuturiza ERRO: ' + err.message); }
}

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
  ui.alert('✅ Sync completo!\n\nTodos os dados do mês ' + MESES[mes - 1] + '/' + ano + ' foram enviados ao Supabase.');
}

function supabaseUpsert(tabela, dados, conflictColumns) {
  const url = SUPABASE_URL + '/rest/v1/' + tabela + '?on_conflict=' + conflictColumns;
  const options = {
    method: 'post',
    headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json', 'Prefer': 'resolution=merge-duplicates,return=minimal' },
    payload: JSON.stringify(dados), muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200 && code !== 201 && code !== 204) throw new Error('Supabase ' + tabela + ' retornou ' + code + ': ' + response.getContentText());
}

function supabaseUpdate(tabela, dados, filtro) {
  const url = SUPABASE_URL + '/rest/v1/' + tabela + '?' + filtro;
  const options = {
    method: 'patch',
    headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json', 'Prefer': 'return=minimal' },
    payload: JSON.stringify(dados), muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200 && code !== 204) throw new Error('Supabase update ' + tabela + ' retornou ' + code + ': ' + response.getContentText());
}

function syncHistoricoCompleto() {
  const ano = 2026;
  const meses = [1, 2, 3, 4];
  meses.forEach(mes => {
    syncDocumentacao(mes, ano);
    syncApp(mes, ano);
    syncAdimplencia(mes, ano);
    syncRecorrencia(mes, ano);
    syncRefuturiza(mes, ano);
    Utilities.sleep(1000);
  });
  syncRecorrenciaVendedor();
  SpreadsheetApp.getUi().alert('✅ Histórico completo sincronizado!');
}

// ============================================================================
// CAMPANHA 14° SALÁRIO
// Planilha separada. Abas criadas automaticamente na primeira chamada.
// Setores são DINÂMICOS — lidos direto da planilha, sem lista fixa no código.
// Checkboxes nativos do Google Sheets nas colunas de mês.
// ============================================================================

// ----------------------------------------------------------------------------
// getCampanha14Data() — endpoint "campanha14"
// Lê somente a aba Destaques (tem tudo: nome, setor, unidade + checkboxes).
// Checkbox nativo do Sheets retorna true/false diretamente.
// ----------------------------------------------------------------------------
function getCampanha14Data() {
  try {
    const ss     = SpreadsheetApp.openById(ID_PLANILHA_CAMPANHA14);
    const shDest = getOrCreateCampanhaSheet(ss, "Destaques",
      ["ID","Nome","Setor","Unidade","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov"]);

    const rows = shDest.getDataRange().getValues().slice(1).filter(r => r[0] && r[1]);
    const mesAtual = new Date().getMonth(); // 0=Jan … 10=Nov

    const colaboradores = rows.map(r => {
      const id      = String(r[0]);
      const nome    = String(r[1]);
      const setor   = String(r[2] || "Sem setor");
      const unidade = String(r[3] || "");

      // Checkbox nativo retorna true/false diretamente
      const meses = MESES_CAMPANHA.map((_, i) => {
        const val = r[4 + i];
        return val === true || val === "TRUE" || val === "SIM" || val === "sim" || val === 1;
      });

      const total          = meses.filter(Boolean).length;
      const mesesRestantes = Math.max(0, 10 - mesAtual);
      const faltam         = Math.max(0, META_MESES_CAMPANHA - total);

      let status = "no_prazo";
      if (total >= META_MESES_CAMPANHA) status = "classificado";
      else if (faltam > mesesRestantes) status = "fora";

      return { id, nome, setor, unidade, meses, totalDestaques: total, status };
    });

    // Agrupa respeitando a ordem fixa dos setores
    const porSetor = {};
    SETORES_CAMPANHA.forEach(s => { porSetor[s] = []; });
    colaboradores.forEach(c => {
      if (porSetor[c.setor]) porSetor[c.setor].push(c);
      else porSetor[c.setor] = [c]; // setor não previsto cai no final
    });
    // Remove setores vazios
    Object.keys(porSetor).forEach(s => { if (!porSetor[s].length) delete porSetor[s]; });

    const resumo = {
      totalColabs:   colaboradores.length,
      classificados: colaboradores.filter(c => c.status === "classificado").length,
      emRisco:       colaboradores.filter(c => c.status === "fora").length,
      noPrazo:       colaboradores.filter(c => c.status === "no_prazo").length,
      metaMeses:     META_MESES_CAMPANHA,
      mesesCampanha: MESES_CAMPANHA
    };

    return { status: "success", data: { resumo, colaboradores, porSetor } };

  } catch(err) {
    return { status: "error", error: err.message };
  }
}

// ----------------------------------------------------------------------------
// getOrCreateCampanhaSheet() — cria aba se não existir
// ----------------------------------------------------------------------------
function getOrCreateCampanhaSheet(ss, nome, headers) {
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(headers);
    const hRange = sh.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight("bold")
          .setBackground("#00843d")
          .setFontColor("#ffffff")
          .setHorizontalAlignment("center");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ----------------------------------------------------------------------------
// inicializarPlanilhaCampanha()
// Rode UMA VEZ manualmente no Apps Script.
// Cria a aba Destaques com checkboxes nativos nas colunas de mês e
// pré-popula com os colaboradores já conhecidos.
// ----------------------------------------------------------------------------
function inicializarPlanilhaCampanha() {
  const ss     = SpreadsheetApp.openById(ID_PLANILHA_CAMPANHA14);
  const shDest = getOrCreateCampanhaSheet(ss, "Destaques",
    ["ID","Nome","Setor","Unidade","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov"]);

const colaboradoresIniciais = [
    // RECEPÇÃO
    [1,  "THIAGO DA SILVA CARDOSO",                   "Recepção",              "Ceilândia"],
    [2,  "MILENA VITORIA LEMOS DA SILVA",              "Recepção",              "Ceilândia"],
    [3,  "DANIELLE LIMA BRITO",                        "Recepção",              "Ceilândia"],
    // REFILIAÇÃO
    [4,  "INGRID CARVALHO RODRIGUES",                  "Refiliação",            "Ceilândia"],
    [5,  "JENIFFER THAYNNA LIMA DA ROCHA",             "Refiliação",            "Ceilândia"],
    [6,  "WANESSA EVELYN CARVALHO",                    "Refiliação",            "Ceilândia"],
    // HOMOLOGAÇÃO
    [7,  "KATIANE FERREIRA DOS SANTOS",                "Homologação",           "Ceilândia"],
    // VENDAS
    [8,  "FRANCISCO ROGEAN ALVES NASCIMENTO",          "Vendas",                "Ceilândia"],
    [9,  "THAYNARA ARAUJO DE OLIVEIRA",                "Vendas",                "Ceilândia"],
    [10, "JOISCIANE DE SOUSA SILVA",                   "Vendas",                "Ceilândia"],
    [11, "MARIA MADALENA SILVA FURTADO",               "Vendas",                "Ceilândia"],
    [12, "MARCUS LUIZ ARAUJO CUNHA",                   "Vendas",                "Ceilândia"],
    // PARCERIAS
    [13, "RAFAEL DOS SANTOS DE JESUS",                 "Parcerias",             "Ceilândia"],
    [14, "THARIK MENEZES SANTIAGO",                    "Parcerias",             "Ceilândia"],
    // ADIMPLÊNCIA
    [15, "ALYNNE GABRIELLE DO NASCIMENTO RIBEIRO",     "Adimplência",           "Ceilândia"],
    [16, "MARIANA NOGUEIRA DOS SANTOS",                "Adimplência",           "Ceilândia"],
    [17, "FLÁVIA ARAÚJO MARQUES",                      "Adimplência",           "Ceilândia"],
    [18, "RHYAN CARLOS FRANCO PEREIRA",                "Adimplência",           "Ceilândia"],
    [19, "KAIQUE DE SOUZA ANDRADE",                    "Adimplência",           "Ceilândia"],
    [20, "ALEXANDRA LISBOA AMORIM",                    "Adimplência",           "Ceilândia"],
    [21, "TAYNARA RODRIGUES DA CRUZ",                  "Adimplência",           "Ceilândia"],
    // LÍDERES
    [22, "FABIANA DA SILVA",                           "Líderes",               "Ceilândia"],
    [23, "CLEONICE CONCEICAO DE MARIA",                "Líderes",               "Ceilândia"],
    [24, "ESTEFFANNIE MENDES DA SILVA",                "Líderes",               "Ceilândia"],
    // RETENÇÃO
    [25, "ISAAC PEREIRA NUNES FERREIRA",               "Retenção",              "Ceilândia"],
    [26, "JACKSON RYLLER DOS SANTOS",                  "Retenção",              "Ceilândia"],
    // PREVISÃO DE DESAFILIAÇÃO
    [27, "ALICE VITORIA LEAL FERRAZ",                  "Previsão Desafiliação", "Ceilândia"],
    // COORDENADORES
    [28, "WILLIAN CARLOS RESENDE GONÇALVES",           "Coordenadores",         "Ceilândia"],
    [29, "JOSENILDO CAMILO DOS SANTOS FILHO",          "Coordenadores",         "Ceilândia"],
  ];

  if (shDest.getLastRow() <= 1) {
    const destRows = colaboradoresIniciais.map(c =>
      [c[0], c[1], c[2], c[3], ...Array(11).fill(false)]
    );
    shDest.getRange(2, 1, destRows.length, 15).setValues(destRows);

    // Aplica checkbox nativo do Sheets nas colunas 5-15 (Jan-Nov)
    const checkRange = shDest.getRange(2, 5, destRows.length, 11);
    checkRange.setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build()
    );

    checkRange.setHorizontalAlignment("center");
    shDest.setColumnWidth(1, 50);   // ID
    shDest.setColumnWidth(2, 240);  // Nome
    shDest.setColumnWidth(3, 120);  // Setor
    shDest.setColumnWidth(4, 110);  // Unidade
    shDest.setColumnWidths(5, 11, 50); // Jan-Nov
    shDest.setFrozenColumns(4);
  }

  SpreadsheetApp.getUi().alert(
    "✅ Planilha Campanha 14° inicializada!\n\n" +
    "Aba criada: Destaques\n\n" +
    "Como usar:\n" +
    "• Marque o checkbox do mês quando o colaborador for destaque\n" +
    "• Para novo setor: adicione uma linha com o setor desejado\n" +
    "  — o painel cria o card automaticamente\n\n" +
    "O painel lê e agrupa por setor dinamicamente."
  );
}

// ============================================================================
// FUNÇÕES AUXILIARES — REINTEGRADAS DO ORIGINAL
// ============================================================================

function corPorSetorRecorrenciaVendedor(setor) {
  const cores = {
    "VENDAS":    "#1e3a8a",
    "REFILIACAO":"#7c3aed",
    "RECEPCAO":  "#059669",
    "OUTROS":    "#4b5563"
  };
  return cores[setor] || "#4b5563";
}

// ============================================================================
// FUNÇÕES DE DEBUG E TESTE — REINTEGRADAS DO ORIGINAL
// ============================================================================

function testarCalculoOkBi() {
  const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
  const sheet = ssDados.getSheetByName("ADIMPLENCIA");
  const data = sheet.getDataRange().getValues();
  const headers = data[1];
  const idx = {
    consultor:    headers.indexOf("Consultor"),
    data:         headers.indexOf("Data"),
    kyc:          headers.indexOf("KYC"),
    mensalidadeOk:headers.indexOf("MENSALIDADE OK"),
    statusBi:     headers.indexOf("STATUS BI")
  };
  let totalOkBi = 0;
  let exemplos = [];
  for (let i = 2; i < data.length; i++) {
    const row        = data[i];
    const consultor  = String(row[idx.consultor]     || "").trim();
    const dataStr    = String(row[idx.data]           || "").trim();
    const kyc        = String(row[idx.kyc]            || "").trim();
    const mensalidade= String(row[idx.mensalidadeOk]  || "").trim();
    const statusBi   = String(row[idx.statusBi]       || "").trim();
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
  SpreadsheetApp.getUi().alert(`okBi para FLÁVIA em janeiro: ${totalOkBi}\n(Esperado: ~55)`);
}

function debugDetalhadoOkBi() {
  const ssDados = SpreadsheetApp.openById(ID_PLANILHA_DADOS);
  const sheet = ssDados.getSheetByName("ADIMPLENCIA");
  const data = sheet.getDataRange().getValues();
  const headers = data[1];
  const idx = {
    consultor:    headers.indexOf("Consultor"),
    data:         headers.indexOf("Data"),
    kyc:          headers.indexOf("KYC"),
    mensalidadeOk:headers.indexOf("MENSALIDADE OK"),
    statusBi:     headers.indexOf("STATUS BI")
  };
  let contadores = { totalFlavia:0, janeiroFlavia:0, statusBiOk:0, mensalidadeOk:0, ambosOk:0, exemplosProblema:[] };
  for (let i = 2; i < Math.min(data.length, 20); i++) {
    const row        = data[i];
    const consultor  = String(row[idx.consultor]     || "").trim();
    const dataStr    = String(row[idx.data]           || "").trim();
    const kyc        = String(row[idx.kyc]            || "").trim();
    const mensalidade= String(row[idx.mensalidadeOk]  || "").trim();
    const statusBi   = String(row[idx.statusBi]       || "").trim();
    if (consultor.toUpperCase().includes("FLÁVIA") || consultor.toUpperCase().includes("FLAVIA")) {
      contadores.totalFlavia++;
      const isJaneiro = dataStr.includes("/01") || dataStr.includes("01/") || dataStr.includes("01-") || dataStr.includes("-01");
      if (isJaneiro) {
        contadores.janeiroFlavia++;
        if (statusBi.toUpperCase()   === "OK") contadores.statusBiOk++;
        if (mensalidade.toUpperCase()=== "OK") contadores.mensalidadeOk++;
        if (statusBi.toUpperCase() === "OK" && mensalidade.toUpperCase() === "OK") {
          contadores.ambosOk++;
        } else {
          contadores.exemplosProblema.push({ linha: i+1, statusBi, mensalidade });
        }
      }
    }
  }
  let mensagem = `🔍 DEBUG DETALHADO OK_BI\n\n`;
  mensagem += `Total Flávia: ${contadores.totalFlavia}\n`;
  mensagem += `Flávia em janeiro: ${contadores.janeiroFlavia}\n`;
  mensagem += `Status BI = "OK": ${contadores.statusBiOk}\n`;
  mensagem += `Mensalidade OK = "OK": ${contadores.mensalidadeOk}\n`;
  mensagem += `AMBOS "OK" (okBi): ${contadores.ambosOk}\n\n`;
  if (contadores.exemplosProblema.length > 0) {
    mensagem += `📋 Exemplos não contados:\n`;
    contadores.exemplosProblema.slice(0,3).forEach(e => {
      mensagem += `• Linha ${e.linha}: BI="${e.statusBi}", Mens="${e.mensalidade}"\n`;
    });
  }
  SpreadsheetApp.getUi().alert(mensagem);
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
      mensagem += "\n🔍 VERIFICAÇÃO MANUAL:\n";
      ["VENDAS","REFILIACAO","RECEPCAO"].forEach(setor => {
        const n = dados.consultores.filter(c => c.setor === setor).length;
        mensagem += `${setor}: ${n} consultores\n`;
      });
      SpreadsheetApp.getUi().alert(mensagem);
      console.log("dadosPorSetor existe?", !!dados.dadosPorSetor);
    } else {
      SpreadsheetApp.getUi().alert(`❌ Erro: ${resultado.error}`);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("Erro no debug: " + error.message);
  }
}
function testarEndpointCampanha14() {
  const resultado = getCampanha14Data();
  console.log(JSON.stringify(resultado));
}
