// ============================================================================
// main.js — ponto de entrada: inicialização, eventos da sidebar/topbar e roteador
// (<script type="module" src="js/main.js"> no index.html)
// ============================================================================
import { SEM_FILTRO, state, hooks } from './config.js';
import { cacheKey, getCached, setCache, invalidateCache, fetchData } from './api.js';
import { iniciarAutenticacao } from './auth.js';
import {
    showSkeleton, hideSkeleton, hideLoading, showToast, showError,
    updateLastUpdateTime, updateFabVisibility, exportPage, initSnapButtons
} from './ui.js';
import { renderResumoDashboard, renderDocumentacaoDashboard, renderAppDashboard } from './views-vendas.js';
import {
    renderAdimplenciaDashboard, renderRecorrenciaDashboard, renderRecorrenciaVendedorDashboard,
    renderRefuturizaDashboard, renderVendasTimDashboard, renderCampanha14Dashboard
} from './views-extras.js';
import { renderEquipeDashboard } from './equipe.js';

const TITLES = {
    resumo:'Resumo Geral', documentacao:'Vendas / Documentação',
    app:'App', adimplencia:'Adimplência', recorrencia:'Recorrência',
    recorrencia_vendedor:'Recorrência Vendedor', refuturiza:'Refuturiza',
    campanha14:'Campanha 14° Salário',
    vendas_tim:'📱 Vendas TIM',
    equipe:'Equipe / Usuários'
};

// ── Init ───────────────────────────────────────────────────────────────────
function init(){
    Chart.defaults.font.family="'Plus Jakarta Sans', sans-serif";
    Chart.defaults.font.size=12; Chart.defaults.color='#5a7a65';

    const dom = state.dom;
    dom.dashboardBtns    = document.querySelectorAll('.sidebar-btn');
    dom.monthSelect      = document.getElementById('monthSelect');
    dom.yearSelect       = document.getElementById('yearSelect');
    dom.dashboardContent = document.getElementById('dashboardContent');
    dom.refreshBtn       = document.getElementById('refreshBtn');
    dom.downloadBtn      = document.getElementById('downloadBtn');
    dom.loadingEl        = document.getElementById('loading');
    dom.lastUpdateEl     = document.getElementById('lastUpdate');
    dom.periodSelector   = document.getElementById('periodSelector');

    const anoAtual = new Date().getFullYear();
    for(let y = anoAtual; y >= anoAtual - 3; y--){
        const opt = document.createElement('option');
        opt.value = y; opt.textContent = y;
        dom.yearSelect.appendChild(opt);
    }
    dom.monthSelect.value = state.currentMonth;
    dom.yearSelect.value  = state.currentYear;

    const topbarTitle = document.getElementById('topbarTitle');
    const logoPadrao = document.getElementById('logoPadrao');
    const logoTim = document.getElementById('logoTim');

    const sidebar        = document.getElementById('sidebar');
    const sidebarOverlay = document.getElementById('sidebarOverlay');
    const sidebarToggle  = document.getElementById('sidebarToggle');
    const collapseBtn    = document.getElementById('collapseBtn');
    const collapseIcon   = document.getElementById('collapseIcon');

    function closeSidebarMobile() {
        if (sidebar)        sidebar.classList.remove('open');
        if (sidebarOverlay) sidebarOverlay.classList.remove('open');
    }

    dom.dashboardBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            dom.dashboardBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.currentDashboard = btn.dataset.dashboard;
            if(topbarTitle) topbarTitle.textContent = TITLES[state.currentDashboard] || state.currentDashboard;

            if(dom.periodSelector){
                dom.periodSelector.classList.toggle('no-filter', SEM_FILTRO.includes(state.currentDashboard));
            }

            if (state.currentDashboard === 'vendas_tim') {
                if (logoPadrao) logoPadrao.style.display = 'block';
                if (logoTim) logoTim.style.display = 'block';
            } else {
                if (logoPadrao) logoPadrao.style.display = 'block';
                if (logoTim) logoTim.style.display = 'none';
            }

            closeSidebarMobile();
            updateFabVisibility();
            loadDashboard();
        });
    });

    if (state.currentDashboard === 'vendas_tim') {
        if (logoPadrao) logoPadrao.style.display = 'block';
        if (logoTim) logoTim.style.display = 'block';
    }

    dom.monthSelect.addEventListener('change', () => { state.currentMonth = parseInt(dom.monthSelect.value); loadDashboard(); });
    dom.yearSelect.addEventListener('change',  () => { state.currentYear  = parseInt(dom.yearSelect.value);  loadDashboard(); });
    dom.refreshBtn.addEventListener('click',   () => { invalidateCache(cacheKey(state.currentDashboard, state.currentMonth, state.currentYear)); loadDashboard(); });
    dom.downloadBtn.addEventListener('click',  exportPage);

    const COLLAPSE_KEY = 'sidebar_collapsed';

    function setSidebarCollapsed(collapsed) {
        if (collapsed) {
            sidebar.classList.add('collapsed');
            if (collapseIcon) { collapseIcon.classList.remove('fa-chevron-left'); collapseIcon.classList.add('fa-chevron-right'); }
        } else {
            sidebar.classList.remove('collapsed');
            if (collapseIcon) { collapseIcon.classList.remove('fa-chevron-right'); collapseIcon.classList.add('fa-chevron-left'); }
        }
        try { localStorage.setItem(COLLAPSE_KEY, collapsed ? '1' : '0'); } catch(e){}
    }

    try { if (localStorage.getItem(COLLAPSE_KEY) === '1') setSidebarCollapsed(true); } catch(e){}

    if (collapseBtn) {
        collapseBtn.addEventListener('click', () => {
            setSidebarCollapsed(!sidebar.classList.contains('collapsed'));
        });
    }

    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', () => {
            sidebar.classList.toggle('open');
            sidebarOverlay.classList.toggle('open');
        });
    }

    if (sidebarOverlay) sidebarOverlay.addEventListener('click', closeSidebarMobile);

    // Câmera de cada bloco (um listener só, por delegação)
    initSnapButtons();

    // ★ Nada é carregado antes do login: o fluxo continua em iniciarAutenticacao()
    iniciarAutenticacao();
}

// Registra as pontes usadas por módulos que não podem importar main.js (evita ciclos).
// Vem ANTES do init(): auth.js e ui.js as usam assim que o login termina.
hooks.loadDashboard   = loadDashboard;
hooks.renderDashboard = renderDashboard;

// Scripts type="module" são adiados até o HTML ser analisado, mas o guard cobre qualquer caso.
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

// ── loadDashboard ──────────────────────────────────────────────────────────
// `seq` (state.loadSeq) identifica ESTA carga. Se o usuário trocar de aba, de período ou de
// unidade enquanto a resposta não chega, loadSeq muda e a resposta velha é descartada
// (antes ela podia desenhar dados de outra aba/unidade por cima da tela atual).
export async function loadDashboard(){
    const seq = ++state.loadSeq;

    if(state.currentDashboard === 'resumo'){ await loadResumoDashboard(seq); return; }

    if(state.currentDashboard === 'equipe'){
        await renderEquipeDashboard();      // tem a própria busca (Edge Function), sem cache
        updateLastUpdateTime();
        return;
    }

    const k      = cacheKey(state.currentDashboard, state.currentMonth, state.currentYear);
    const cached = getCached(k);

    if(cached){
        state.data = cached;
        renderDashboard();
        updateLastUpdateTime();
        fetchData(state.currentDashboard, state.currentMonth, state.currentYear).then(r => {
            if(r.status==='success' && JSON.stringify(r.data) !== JSON.stringify(cached)){
                setCache(k, r.data);
                if(seq !== state.loadSeq) return;
                state.data = r.data;
                renderDashboard();
                updateLastUpdateTime();
                showToast('Dados atualizados ✓');
            }
        }).catch(_=>{});
        return;
    }

    showSkeleton();
    try{
        const result = await fetchData(state.currentDashboard, state.currentMonth, state.currentYear);
        if(seq !== state.loadSeq) return;
        hideSkeleton();
        if(result.status === 'success'){
            setCache(k, result.data);
            state.data = result.data;
            if(result.fonte === 'appscript') showToast('⚠️ Supabase sem dados — carregado do Apps Script');
            renderDashboard();
        } else {
            showError(result.error || 'Erro desconhecido', true);
        }
    }catch(err){
        if(seq !== state.loadSeq) return;
        hideSkeleton();
        const online = navigator.onLine;
        if(!online){
            showError('Sem conexão com a internet. Verifique sua rede e tente novamente.', false);
        } else {
            showError('Não foi possível conectar ao servidor. Tente novamente em instantes.', true);
        }
    }
    updateLastUpdateTime();
}

// ── Roteador ───────────────────────────────────────────────────────────────
export function renderDashboard(){
    hideLoading();
    const data = state.data;
    switch(state.currentDashboard){
        case 'documentacao':         renderDocumentacaoDashboard(data);        break;
        case 'app':                  renderAppDashboard(data);                 break;
        case 'adimplencia':          renderAdimplenciaDashboard(data);         break;
        case 'recorrencia':          renderRecorrenciaDashboard(data);         break;
        case 'recorrencia_vendedor': renderRecorrenciaVendedorDashboard(data); break;
        case 'refuturiza':           renderRefuturizaDashboard(data);          break;
        case 'campanha14':           renderCampanha14Dashboard(data);          break;
        case 'vendas_tim':           renderVendasTimDashboard(data);           break;
        case 'equipe':               renderEquipeDashboard();                  break;
        default: showError('Dashboard não encontrado: ' + state.currentDashboard);
    }
}

// ── loadResumoDashboard ────────────────────────────────────────────────────
async function loadResumoDashboard(seq){
    const eps     = ['documentacao','app','adimplencia'];
    const results = {};

    const allCached = eps.every(ep => !!getCached(cacheKey(ep, state.currentMonth, state.currentYear)));

    if(allCached){
        eps.forEach(ep => { results[ep] = getCached(cacheKey(ep, state.currentMonth, state.currentYear)); });
        renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
        updateLastUpdateTime();
        Promise.all(eps.map(async ep => {
            const k = cacheKey(ep, state.currentMonth, state.currentYear);
            const r = await fetchData(ep, state.currentMonth, state.currentYear);
            if(r.status==='success'){ setCache(k, r.data); results[ep]=r.data; }
        })).then(() => {
            if(seq !== state.loadSeq) return;
            if(results['documentacao']) renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
        }).catch(_=>{});
        return;
    }

    showSkeleton();

    await Promise.all(eps.map(async ep => {
        const k = cacheKey(ep, state.currentMonth, state.currentYear);
        const cached = getCached(k);
        if(cached){ results[ep] = cached; return; }
        try{
            const r = await fetchData(ep, state.currentMonth, state.currentYear);
            if(r.status==='success'){ setCache(k, r.data); results[ep]=r.data; }
        }catch(_){}
    }));

    if(seq !== state.loadSeq) return;
    hideSkeleton();
    if(!results['documentacao']){
        showError('Não foi possível carregar os dados de Vendas. Verifique a conexão e tente novamente.', true);
        return;
    }
    renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
    updateLastUpdateTime();
}
