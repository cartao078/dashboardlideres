// ============================================================================
// auth.js — login (Supabase Auth), unidade ativa e suporte a MÚLTIPLAS unidades
// ============================================================================
import { AUTH, SUPABASE_URL, SUPABASE_ANON, COR_PRIMARIA_PADRAO, state, hooks } from './config.js';
import { limparCaches } from './api.js';
import {
    contentEl, destroyAllCharts, hideLoading, initPeriodModal, updateFabVisibility
} from './ui.js';
import { escapeHtml } from './utils.js';

const LS_UNIDADE_ATIVA = 'cdt_unidade_ativa';   // última unidade escolhida (por navegador)

function authEl(id){ return document.getElementById(id); }

function authMostrarErro(msg){
    const el = authEl('authError');
    if (!el) return;
    el.textContent = msg || '';
    el.classList.toggle('show', !!msg);
}

function authModoChecando(on){
    const card = authEl('authCard');
    if (card) card.classList.toggle('is-checking', !!on);
}

function mostrarLogin(msg){
    document.body.classList.add('auth-pending');
    authModoChecando(false);
    authMostrarErro(msg || '');
    const btn = authEl('authSubmit');
    if (btn) { btn.disabled = false; btn.querySelector('span').textContent = 'Entrar'; }
    const email = authEl('authEmail');
    if (email && !email.value) email.focus();
}

function esconderLogin(){
    document.body.classList.remove('auth-pending');
    authMostrarErro('');
    const pass = authEl('authPassword');
    if (pass) pass.value = '';
}

function normalizarModulos(m){
    if (Array.isArray(m)) return m.map(String);
    if (typeof m === 'string') { try { const p = JSON.parse(m); return Array.isArray(p) ? p.map(String) : []; } catch(_) { return []; } }
    return [];
}

// Ajusta cor (#rrggbb) misturando com branco (t>0) ou preto (t<0)
function ajustarCor(hex, t){
    const n = parseInt(hex.slice(1), 16);
    const c = [(n>>16)&255, (n>>8)&255, n&255].map(v => {
        const alvo = t >= 0 ? 255 : 0;
        return Math.round(v + (alvo - v) * Math.abs(t));
    });
    return '#' + c.map(v => v.toString(16).padStart(2,'0')).join('');
}

function aplicarCorPrimaria(cor){
    const ok = typeof cor === 'string' && /^#[0-9a-fA-F]{6}$/.test(cor);
    const base = ok ? cor : COR_PRIMARIA_PADRAO;
    const root = document.documentElement.style;
    if (base.toLowerCase() === COR_PRIMARIA_PADRAO) {
        // cor padrão: remove overrides e deixa o styles.css valer (visual atual intacto)
        ['--primary','--primary-dark','--primary-light','--primary-mid','--success'].forEach(v => root.removeProperty(v));
        return;
    }
    root.setProperty('--primary', base);
    root.setProperty('--primary-dark', ajustarCor(base, -0.25));
    root.setProperty('--primary-mid', ajustarCor(base, 0.15));
    root.setProperty('--primary-light', ajustarCor(base, 0.88));
}

// Mostra/esconde o botão "Equipe" (e o rótulo da seção) conforme o papel na unidade ativa.
// Isto é só UX: a Edge Function gerenciar-equipe valida o papel no servidor.
function aplicarMenuAdmin(isAdmin){
    const sec = authEl('sidebarAdminSection');
    if (sec) sec.hidden = !isAdmin;
    const btn = document.querySelector('.sidebar-btn[data-dashboard="equipe"]');
    if (btn) { btn.hidden = !isAdmin; btn.style.display = isAdmin ? '' : 'none'; }
}

export function aplicarUnidade(u){
    // Identidade
    document.title = `Dashboard — ${u.nome_exibicao}`;
    const nomeSide = authEl('sidebarUnidadeNome');
    if (nomeSide) nomeSide.textContent = String(u.nome_exibicao || '').toUpperCase();
    const foot = authEl('footerUnidade');
    if (foot) foot.textContent = String(u.nome_exibicao || '').toUpperCase();

    aplicarCorPrimaria(u.cor_primaria);

    // Sempre define o logo: ao trocar de unidade, o logo da anterior não pode "sobrar".
    const logo = authEl('logoPadrao');
    if (logo) logo.src = (u.logo_url && /^https:\/\//i.test(u.logo_url)) ? u.logo_url : 'logo.png';

    // Meta de aprovação da unidade (disponível para os renders)
    state.metaAprovacao = Number.isFinite(Number(u.meta_aprovacao)) ? Number(u.meta_aprovacao) : 90;

    // Módulos: só os botões de modulos_ativos ficam visíveis ("equipe" depende do papel, não dos módulos)
    const ativos = normalizarModulos(u.modulos_ativos);
    document.querySelectorAll('.sidebar-btn').forEach(btn => {
        if (btn.dataset.dashboard === 'equipe') return;
        btn.hidden = !ativos.includes(btn.dataset.dashboard);
        btn.style.display = btn.hidden ? 'none' : '';
    });
    aplicarMenuAdmin(u.papel === 'admin');
}

// ── Vínculos do usuário (multi-unidade) ────────────────────────────────────
async function buscarVinculos(userId){
    let r = await AUTH.client.from('usuarios_unidades').select('unidade_id,papel').eq('user_id', userId);
    if (r.error) {
        // Compatibilidade: se a coluna `papel` ainda não existe, todos entram como "leitor".
        r = await AUTH.client.from('usuarios_unidades').select('unidade_id').eq('user_id', userId);
    }
    return r.error ? [] : (r.data || []);
}

// Retorna TODAS as unidades ativas do usuário (cada uma com .papel), na ordem de id.
async function carregarUnidadesDoUsuario(userId){
    const vinculos = await buscarVinculos(userId);
    if (vinculos.length === 0) return [];

    const papelPorUnidade = {};
    vinculos.forEach(v => { papelPorUnidade[v.unidade_id] = (v.papel === 'admin') ? 'admin' : 'leitor'; });

    const { data: unidades, error } = await AUTH.client
        .from('unidades')
        .select('id,nome_exibicao,logo_url,cor_primaria,modulos_ativos,meta_aprovacao,ativo')
        .in('id', Object.keys(papelPorUnidade)).eq('ativo', true).order('id');
    if (error || !unidades || unidades.length === 0) return [];

    return unidades.map(u => ({ ...u, papel: papelPorUnidade[u.id] || 'leitor' }));
}

function lembrarUnidade(id){ try { localStorage.setItem(LS_UNIDADE_ATIVA, id); } catch(_) {} }
function unidadeLembrada(){ try { return localStorage.getItem(LS_UNIDADE_ATIVA); } catch(_) { return null; } }

// ── Seletor de unidades (só aparece com 2+ unidades) ───────────────────────
function montarSeletorUnidades(){
    const wrap = authEl('unidadeSwitch');
    const sel  = authEl('unidadeSelect');
    if (!wrap || !sel) return;

    const multi = AUTH.unidades.length > 1;
    wrap.hidden = !multi;
    if (!multi) { sel.innerHTML = ''; return; }

    const porNome = [...AUTH.unidades].sort((a, b) => String(a.nome_exibicao).localeCompare(String(b.nome_exibicao), 'pt-BR'));
    sel.innerHTML = porNome.map(u => `<option value="${escapeHtml(u.id)}">${escapeHtml(u.nome_exibicao)}</option>`).join('');
    sel.value = AUTH.unidade ? AUTH.unidade.id : porNome[0].id;
}

// Aplica a unidade e abre o dashboard certo (usado no login e na troca de unidade)
function abrirUnidade(u){
    aplicarUnidade(u);

    const visiveis   = Array.from(document.querySelectorAll('.sidebar-btn')).filter(b => !b.hidden);
    const dashboards = visiveis.filter(b => b.dataset.dashboard !== 'equipe');

    if (dashboards.length === 0) {
        state.loadSeq++;                 // invalida qualquer carga em andamento
        destroyAllCharts();
        hideLoading();
        contentEl().innerHTML = '<div class="error-message"><h3>Nenhum módulo ativo</h3><p>Sua unidade ainda não tem dashboards habilitados.</p></div>';
        return;
    }

    const atual = visiveis.find(b => b.dataset.dashboard === state.currentDashboard);
    if (atual) {
        document.querySelectorAll('.sidebar-btn').forEach(b => b.classList.toggle('active', b === atual));
        updateFabVisibility();
        hooks.loadDashboard();
    } else {
        dashboards[0].click();   // o handler do botão atualiza título/filtro e chama loadDashboard()
    }
}

// Troca em tempo real, sem recarregar a página.
export function trocarUnidade(id){
    const u = AUTH.unidades.find(x => x.id === id);
    if (!u || (AUTH.unidade && AUTH.unidade.id === u.id)) return;

    AUTH.unidade = u;           // as chaves de cache incluem a unidade: nada de dados misturados
    lembrarUnidade(u.id);
    state.loadSeq++;            // respostas em voo da unidade anterior serão descartadas
    state.data = null;
    destroyAllCharts();
    contentEl().innerHTML = '';
    abrirUnidade(u);
}

export async function entrarNoDashboard(session){
    if (AUTH.entrando) return;
    AUTH.entrando = true;
    try {
        AUTH.session = session;
        authModoChecando(true);

        const unidades = await carregarUnidadesDoUsuario(session.user.id);
        if (unidades.length === 0) {
            await AUTH.client.auth.signOut();
            mostrarLogin('Seu usuário não está vinculado a uma unidade ativa. Fale com o administrador.');
            return;
        }

        AUTH.unidades = unidades;
        const lembrada = unidadeLembrada();
        AUTH.unidade = unidades.find(u => u.id === lembrada) || unidades[0];
        lembrarUnidade(AUTH.unidade.id);

        montarSeletorUnidades();
        esconderLogin();

        if (!AUTH.periodModalOk) { initPeriodModal(); AUTH.periodModalOk = true; }

        abrirUnidade(AUTH.unidade);
    } catch (err) {
        console.error('[auth] erro ao entrar:', err);
        mostrarLogin('Não foi possível carregar sua unidade. Tente novamente.');
    } finally {
        AUTH.entrando = false;
    }
}

function limparSessaoLocal(){
    // Gráficos, dados em memória e caches (inclui localStorage) da sessão anterior
    destroyAllCharts();
    state.data = null;
    state.loadSeq++;
    limparCaches();
    try { localStorage.removeItem(LS_UNIDADE_ATIVA); } catch(_) {}

    AUTH.unidade = null;
    AUTH.unidades = [];
    AUTH.session = null;
    montarSeletorUnidades();

    // Volta identidade/visual ao neutro e reexibe os botões
    aplicarCorPrimaria(COR_PRIMARIA_PADRAO);
    document.title = 'Dashboard — Entrar';
    const nomeSide = authEl('sidebarUnidadeNome'); if (nomeSide) nomeSide.textContent = '';
    const foot = authEl('footerUnidade'); if (foot) foot.textContent = '';
    const logo = authEl('logoPadrao'); if (logo) logo.src = 'logo.png';
    document.querySelectorAll('.sidebar-btn').forEach(b => { b.hidden = false; b.style.display = ''; });
    aplicarMenuAdmin(false);
    if (contentEl()) contentEl().innerHTML = '';
}

export async function sairDoDashboard(){
    try { await AUTH.client.auth.signOut(); } catch(_) {}
    limparSessaoLocal();
    mostrarLogin('');
}

export async function iniciarAutenticacao(){
    authModoChecando(true);

    if (!window.supabase || !window.supabase.createClient) {
        mostrarLogin('Não foi possível carregar o módulo de login. Verifique sua conexão e recarregue a página.');
        return;
    }

    AUTH.client = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON, {
        auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false }
    });

    // Mantém o token atualizado; sessão encerrada em outra aba → volta ao login.
    // (callback síncrono de propósito: não chamar o Supabase aqui dentro)
    AUTH.client.auth.onAuthStateChange((event, session) => {
        if (event === 'SIGNED_OUT') {
            if (AUTH.unidade) { limparSessaoLocal(); mostrarLogin('Sessão encerrada.'); }
            return;
        }
        if (session) AUTH.session = session;
    });

    const form = authEl('authForm');
    form.addEventListener('submit', async (ev) => {
        ev.preventDefault();
        const email = authEl('authEmail').value.trim();
        const senha = authEl('authPassword').value;
        if (!email || !senha) { authMostrarErro('Informe e-mail e senha.'); return; }

        const btn = authEl('authSubmit');
        btn.disabled = true;
        btn.querySelector('span').textContent = 'Entrando…';
        authMostrarErro('');

        const { data, error } = await AUTH.client.auth.signInWithPassword({ email, password: senha });
        if (error || !data || !data.session) {
            const msg = (error && /invalid login credentials/i.test(error.message))
                ? 'E-mail ou senha incorretos.'
                : 'Não foi possível entrar. Verifique sua conexão e tente novamente.';
            mostrarLogin(msg);
            return;
        }
        await entrarNoDashboard(data.session);
    });

    const logoutBtn = authEl('logoutBtn');
    if (logoutBtn) logoutBtn.addEventListener('click', sairDoDashboard);

    // Seletor de unidades da sidebar
    const unidadeSelect = authEl('unidadeSelect');
    if (unidadeSelect) unidadeSelect.addEventListener('change', () => trocarUnidade(unidadeSelect.value));

    // Sessão já existente (recarregou a página)?
    try {
        const { data } = await AUTH.client.auth.getSession();
        if (data && data.session) { await entrarNoDashboard(data.session); return; }
    } catch (err) {
        console.error('[auth] getSession falhou:', err);
    }
    mostrarLogin('');
}
