// ============================================================================
// DASHBOARD V21.0 - app.js
// ============================================================================

const API_URL = 'https://script.google.com/macros/s/AKfycby5ffZWf5lrHveg3SZwqkX6U5e0d87NkNufTWR9vZzzTAq0r7kIheKF5CT1QgiNzXUQHA/exec';

const SUPABASE_URL  = 'https://vycjtmjvkwvxunxtkdyi.supabase.co';
const SUPABASE_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ5Y2p0bWp2a3d2eHVueHRrZHlpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE2MDY2OTYsImV4cCI6MjA4NzE4MjY5Nn0.5w4z1hX2a3b4c5d6e7f8g9h0i1j2k3l4m5n6o7p8q9r0';

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
// RENDERIZAÇÃO DOS DASHBOARDS (versão simplificada para teste)
// ============================================================================

function renderDashboard(){
    hideLoading();
    const data = window._dashboardData;
    
    // Versão simplificada para teste
    dashboardContent.innerHTML = `
        <h2>Dashboard ${currentDashboard}</h2>
        <pre>${JSON.stringify(data, null, 2)}</pre>
    `;
}

async function loadResumoDashboard(){
    dashboardContent.innerHTML = '<h2>Carregando Resumo...</h2>';
    // Implementação completa depois
}

function updateLastUpdateTime(){
    if(lastUpdateEl) lastUpdateEl.textContent = new Date().toLocaleString('pt-BR');
}

// ============================================================================
// UTILITÁRIOS
// ============================================================================

function destroyChart(id){ 
    if(chartInstances[id]){
        chartInstances[id].destroy();
        delete chartInstances[id];
    } 
}

function createChart(id,cfg){ 
    destroyChart(id); 
    const ctx=document.getElementById(id); 
    if(!ctx) return; 
    chartInstances[id]=new Chart(ctx,cfg); 
}

function showProgressLoading(label='Carregando...', pct=0){
    // Implementação simplificada
    hideLoading();
    const div = document.createElement('div');
    div.innerHTML = `<p>${label} ${pct}%</p>`;
    dashboardContent.appendChild(div);
}

function updateProgress(pct, label){
    const el = document.querySelector('#progressLoader');
    if(el) el.innerHTML = `<p>${label} ${pct}%</p>`;
}

function hideLoading(){
    if(loadingEl) loadingEl.style.display = 'none';
    const pl = document.getElementById('progressLoader');
    if(pl) pl.remove();
}

function showToast(msg){
    alert(msg); // Simplificado para teste
}

function showError(msg){ 
    hideLoading(); 
    dashboardContent.innerHTML=`<div style="color:red; padding:20px;">❌ ${msg}</div>`; 
}

// ============================================================================
// EXPORT
// ============================================================================

async function exportPage(){
    alert('Função de exportação desabilitada para teste');
}

function getMonthName(m){ 
    return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][m-1]; 
}

// ============================================================================
// PAINEL ADMIN (simplificado)
// ============================================================================

function toggleAdminPanel() {
    alert('Painel Admin será implementado depois');
}
