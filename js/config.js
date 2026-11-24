// ============================================================================
// config.js — constantes, estado de autenticação e estado de navegação
// Sem dependências: é o módulo-base que todos os outros importam.
// ============================================================================

export const API_URL = 'https://script.google.com/macros/s/AKfycbxY4lTmtAnQHGo_XuWDgbUOwsa2V6PWHVvWgtu1--2tWjux6u6kOzom0eS7L6ir-60Qsg/exec';

export const SUPABASE_URL  = 'https://vycjtmjvkwvxunxtkdyi.supabase.co';
export const SUPABASE_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ5Y2p0bWp2a3d2eHVueHRrZHlpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE2MDY2OTYsImV4cCI6MjA4NzE4MjY5Nn0.wOoAZpA1i-320E8Rc-Ry6nk0KYsedFXb3aS4gkmbjHU';

// Cores da marca (UI)
export const C = {
    green:'#00a651', greenFade:'rgba(0,166,81,0.12)',
    lime:'#7ed321',  teal:'#0d9488',
    warn:'#f59e0b',  warnFade:'rgba(245,158,11,0.12)',
    danger:'#ef4444',gray:'#6b8072', border:'#d1e8d9',
};

// Paleta equilibrada para gráficos (menos saturada que a da marca)
export const CH = {
    green:  '#10b981',
    warn:   '#f59e0b',
    danger: '#ef4444',
    teal:   '#0d9488',
    lime:   '#84cc16',
    gray:   '#94a3b8',
    grid:   '#eef2f6',
    text:   '#64748b'
};

// ── Autenticação / Unidade (multi-franquia) ────────────────────────────────
// Apps Script (API_URL) só serve a unidade padrão; as demais usam apenas Supabase.
export const UNIDADE_PADRAO_ID   = 'cdt-ceilandia';
export const COR_PRIMARIA_PADRAO = '#00a651';

// unidade  → unidade ATIVA agora (inclui .papel: 'admin' | 'leitor')
// unidades → todas as unidades às quais o usuário está vinculado
export const AUTH = {
    client: null, session: null,
    unidade: null, unidades: [],
    entrando: false, periodModalOk: false
};

export const RANKING_SETORES = ['VENDAS','RECEPCAO','REFILIACAO'];
// Telas que ignoram o filtro de mês/ano ('equipe' é administrativa, sem período)
export const SEM_FILTRO      = ['recorrencia_vendedor','campanha14','equipe'];

// ── Estado de navegação ────────────────────────────────────────────────────
// Módulos ES não permitem reatribuir `let` importado, por isso o estado mutável
// vive neste objeto: state.currentDashboard = 'app' funciona em qualquer módulo.
export const state = {
    currentDashboard: 'resumo',
    currentMonth:     new Date().getMonth() + 1,
    currentYear:      new Date().getFullYear(),
    data:             null,   // dados do dashboard atual (antes: window._dashboardData)
    metaAprovacao:    90,     // meta da unidade ativa (antes: window.META_APROVACAO)
    loadSeq:          0,      // incrementa a cada carga; descarta respostas "velhas"
    dom: {                    // preenchido por main.js no DOMContentLoaded
        dashboardBtns: null, monthSelect: null, yearSelect: null,
        dashboardContent: null, refreshBtn: null, downloadBtn: null,
        loadingEl: null, lastUpdateEl: null, periodSelector: null
    }
};

// Pontes para funções que ficam em main.js (evita import circular):
//   hooks.loadDashboard   → usado por auth.js e ui.js (botão "Tentar novamente")
//   hooks.renderDashboard → usado por api.js (silentRefresh)
export const hooks = {
    loadDashboard:   () => {},
    renderDashboard: () => {}
};
