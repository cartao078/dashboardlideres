// ============================================================================
// api.js — cache, consultas ao Supabase (com fallback Apps Script) e auto-refresh
// ============================================================================
import { AUTH, API_URL, SUPABASE_URL, SUPABASE_ANON, UNIDADE_PADRAO_ID, SEM_FILTRO, state, hooks } from './config.js';
import { updateLastUpdateTime } from './ui.js';
import { renderResumoDashboard } from './views-vendas.js';

// ── Cache ──────────────────────────────────────────────────────────────────
const CACHE_MEM  = new Map();
const CACHE_TTL  = 30 * 60 * 1000;
const LS_PREFIX  = 'cdt_dash_';

// A chave inclui a unidade: cache (e localStorage) nunca é compartilhado entre franquias/usuários.
export function cacheKey(ep,m,y){
    const u = AUTH.unidade ? AUTH.unidade.id : '-';
    return SEM_FILTRO.includes(ep) ? `${u}|${ep}` : `${u}|${ep}:${m}:${y}`;
}

export function getCached(k){
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

export function setCache(k,d){
    const entry = {data:d, ts:Date.now()};
    CACHE_MEM.set(k, entry);
    try{ localStorage.setItem(LS_PREFIX+k, JSON.stringify(entry)); }catch(_){}
}

export function invalidateCache(k){
    CACHE_MEM.delete(k);
    try{ localStorage.removeItem(LS_PREFIX+k); }catch(_){}
}

// Limpa TODOS os caches (memória + localStorage) — usado no logout
export function limparCaches(){
    CACHE_MEM.clear();
    try {
        Object.keys(localStorage).filter(k => k.startsWith(LS_PREFIX)).forEach(k => localStorage.removeItem(k));
    } catch(_) {}
}

// ============================================================================
// CAMADA DE DADOS — Supabase primeiro, fallback Apps Script
// ============================================================================

export async function getAccessToken() {
    try {
        if (AUTH.client) {
            const { data } = await AUTH.client.auth.getSession();   // renova o token se preciso
            if (data && data.session) { AUTH.session = data.session; return data.session.access_token; }
        }
    } catch (_) {}
    return AUTH.session ? AUTH.session.access_token : null;
}

export async function fetchSupabase(endpoint, mes, ano) {
    if (endpoint === 'campanha14') return null;
    const unidade = AUTH.unidade;
    if (!unidade) return null;   // sem login não há consulta

    try {
        let url;
        const token = await getAccessToken();
        if (!token) return null;
        const headers = {
            'apikey': SUPABASE_ANON,
            'Authorization': `Bearer ${token}`
        };
        // ★ Multi-franquia: toda consulta é filtrada pela unidade do usuário
        const uf = `&unidade_id=eq.${encodeURIComponent(unidade.id)}`;
        if (endpoint === 'recorrencia_vendedor') {
            url = `${SUPABASE_URL}/rest/v1/recorrencia_vendedor?select=*&limit=1${uf}`;
        } else if (endpoint === 'app') {
            url = `${SUPABASE_URL}/rest/v1/app_dashboard?mes=eq.${mes}&ano=eq.${ano}&select=*${uf}`;
        } else if (endpoint === 'vendas_tim') {
            url = `${SUPABASE_URL}/rest/v1/vendas_tim?mes=eq.${mes}&ano=eq.${ano}&select=*${uf}`;
        } else if (endpoint === 'recorrencia') {
            // ★ NOVO: Busca a nova tabela de qualidade
            url = `${SUPABASE_URL}/rest/v1/qualidade_retencao?mes=eq.${mes}&ano=eq.${ano}&select=*${uf}`;
        } else {
            url = `${SUPABASE_URL}/rest/v1/${endpoint}?mes=eq.${mes}&ano=eq.${ano}&select=*${uf}`;
        }
        const resp = await fetch(url, { headers });
        if (!resp.ok) return null;
        const rows = await resp.json();
        if (!rows || rows.length === 0) return null;
        return rowToData(endpoint, rows[0]);
    } catch(_) {
        return null;
    }
}

export function rowToData(endpoint, row) {
    if (!row) return null;

    const parseJ = v => {
        if (typeof v === 'string') { try { return JSON.parse(v); } catch(_) { return {}; } }
        return v || {};
    };

    if (endpoint === 'documentacao') {
        const totVazio = {total:0,cancelados:0,aprovados:0,reprovados:0,expirado:0,pendente:0,naoEnviado:0,pendencias:0,promo:{total:0,cancelados:0,aprovados:0},normal:{total:0,cancelados:0,aprovados:0}};
        return {
            mes: row.mes_nome, ano: row.ano,
            temColunasPromo: row.tem_promocao || false,
            geral: {
                total: row.geral_total, cancelados: row.geral_cancelados||0,
                aprovados: row.geral_aprovados, pendencias: row.geral_pendencias,
                reprovados: row.geral_reprovados, expirado: row.geral_expirado,
                pendente: row.geral_pendente, naoEnviado: row.geral_nao_enviado,
                promo:  parseJ(row.geral_promo)  || {total:0,cancelados:0,aprovados:0},
                normal: parseJ(row.geral_normal) || {total:0,cancelados:0,aprovados:0}
            },
            vendasLoja: {
                total: row.loja_total, cancelados: row.loja_cancelados||0,
                aprovados: row.loja_aprovados, pendencias: row.loja_pendencias,
                reprovados: row.loja_reprovados, expirado: row.loja_expirado,
                pendente: row.loja_pendente, naoEnviado: row.loja_nao_enviado,
                promo:  parseJ(row.loja_promo)  || {total:0,cancelados:0,aprovados:0},
                normal: parseJ(row.loja_normal) || {total:0,cancelados:0,aprovados:0}
            },
            vendasWeb: {
                total: row.web_total, cancelados: row.web_cancelados||0,
                aprovados: row.web_aprovados, pendencias: row.web_pendencias,
                reprovados: row.web_reprovados, expirado: row.web_expirado,
                pendente: row.web_pendente, naoEnviado: row.web_nao_enviado,
                promo:  {total:0,cancelados:0,aprovados:0},
                normal: {total:0,cancelados:0,aprovados:0}
            },
            temRecepcao: row.tem_recepcao || false,
            recepcao: {
                indicacao: parseJ(row.recepcao_indicacao) || totVazio,
                direta:    parseJ(row.recepcao_direta)    || totVazio,
                total:     parseJ(row.recepcao_total)     || totVazio
            },
            consultores: typeof row.consultores === 'string' ? JSON.parse(row.consultores) : (row.consultores || [])
        };
    }

    if (endpoint === 'app') {
        return {
            mes: row.mes_nome, ano: row.ano,
            geral:   { total: row.geral_total, sim: row.geral_sim, nao: row.geral_nao, cancelado: row.geral_cancelado, outros: row.geral_outros },
            appLoja: { total: row.loja_total, sim: row.loja_sim, nao: row.loja_nao, cancelado: row.loja_cancelado, outros: row.loja_outros },
            appWeb:  { total: row.web_total, sim: row.web_sim, nao: row.web_nao, cancelado: row.web_cancelado, outros: row.web_outros },
            equipeVendas:      parseJ(row.equipe_vendas)      || {total:0,sim:0,nao:0,cancelado:0,outros:0},
            equipeConciliacao: parseJ(row.equipe_conciliacao) || {total:0,sim:0,nao:0,cancelado:0,outros:0},
            consultores: typeof row.consultores === 'string' ? JSON.parse(row.consultores) : (row.consultores || []),
            consultorasRetencao: typeof row.consultoras_retencao === 'string' ? JSON.parse(row.consultoras_retencao) : (row.consultoras_retencao || [])
        };
    }

    if (endpoint === 'adimplencia') {
        return {
            mes: row.mes_nome, ano: row.ano,
            geral: { totalTrocas: row.geral_total_trocas, mensOk: row.geral_mens_ok, mensAberto: row.geral_mens_aberto, mensAtraso: row.geral_mens_atraso, aprovados: row.geral_aprovados, pendentes: row.geral_pendentes, totalBi: row.geral_total_bi, foraBi: row.geral_fora_bi, okBi: row.geral_ok_bi, percentualAprovado: row.geral_percentual_aprovado },
            consultores: typeof row.consultores === 'string' ? JSON.parse(row.consultores) : (row.consultores || [])
        };
    }

    if (endpoint === 'recorrencia') {
        // ★ NOVO: Lê o JSON completo da tabela qualidade_retencao
        return typeof row.dados === 'string' ? JSON.parse(row.dados) : (row.dados || null);
    }

    if (endpoint === 'recorrencia_vendedor') {
        return {
            mes: 'TODOS OS MESES', ano: 'TODOS OS ANOS',
            geral: { totalVendasPromocao: row.geral_total_vendas, totalOk: row.geral_total_ok, totalEmAberto: row.geral_total_em_aberto, totalAtraso: row.geral_total_atraso, totalOutros: row.geral_total_outros, percentualVendasOk: row.geral_percentual_ok },
            consultores: typeof row.consultores === 'string' ? JSON.parse(row.consultores) : (row.consultores || []),
            dadosPorSetor: typeof row.dados_por_setor === 'string' ? JSON.parse(row.dados_por_setor) : row.dados_por_setor,
            totalConsultores: row.total_consultores,
            totalRegistros: row.total_registros
        };
    }

    if (endpoint === 'refuturiza') {
        let consultores = [];
        try {
            if (typeof row.consultores === 'string') consultores = JSON.parse(row.consultores);
            else if (Array.isArray(row.consultores)) consultores = row.consultores;
            else if (row.consultores && typeof row.consultores === 'object') consultores = Object.values(row.consultores);
        } catch(e) { consultores = []; }

        const consultoresNormalizados = consultores.map(c => ({
            nome: c.nome || 'Desconhecido',
            total: Number(c.total) || Number(c.totalVendas) || 0,
            cancelado: Number(c.cancelado) || Number(c.cancelados) || 0,
            comLigacao: Number(c.comLigacao) || Number(c.com_ligacao) || 0,
            semLigacao: Number(c.semLigacao) || Number(c.sem_ligacao) || 0
        }));

        const consultoresUnicos = {};
        consultoresNormalizados.forEach(c => {
            const key = c.nome.toUpperCase();
            if (!consultoresUnicos[key] || c.total > consultoresUnicos[key].total) consultoresUnicos[key] = c;
        });
        const consultoresFinal = Object.values(consultoresUnicos);

        const totalBruto = Number(row.geral_total) || 0;
        const cancelado  = Number(row.geral_cancelado) || 0;
        const comLigacao = Number(row.geral_com_ligacao) || 0;
        const semLigacao = Number(row.geral_sem_ligacao) || 0;
        const baseLiquida = totalBruto - cancelado;
        const percentualCorrigido = baseLiquida > 0 ? Math.round((comLigacao / baseLiquida) * 100) : 0;

        if (consultoresFinal.length === 0 && totalBruto > 0) {
            consultoresFinal.push({ nome:'Total Geral', total:totalBruto, cancelado, comLigacao, semLigacao });
        }

        return {
            mes: row.mes_nome || 'Mês Desconhecido',
            ano: row.ano || 'Ano Desconhecido',
            geral: { total: totalBruto, comLigacao, semLigacao, cancelado, baseLiquida, percentualComLigacao: percentualCorrigido },
            consultores: consultoresFinal
        };
    }

    if (endpoint === 'vendas_tim') {
        let vendedores = [];
        try {
            if (typeof row.vendedores === 'string') vendedores = JSON.parse(row.vendedores);
            else if (Array.isArray(row.vendedores)) vendedores = row.vendedores;
            else if (row.vendedores && typeof row.vendedores === 'object') vendedores = Object.values(row.vendedores);
        } catch(e) { vendedores = []; }

        return {
            mes: row.mes_nome || 'Mês Desconhecido',
            ano: row.ano || 'Ano Desconhecido',
            geral: {
                total: row.geral_total || 0,
                ativacao: row.geral_ativacao || 0,
                portabilidade: row.geral_portabilidade || 0,
                pctAtivacao: row.geral_pct_ativacao || 0,
                pctPortabilidade: row.geral_pct_portabilidade || 0
            },
            vendedores: vendedores,
            totalVendedores: row.total_vendedores || vendedores.length || 0
        };
    }

    return null;
}

export async function fetchData(endpoint, mes, ano) {
    const unidade = AUTH.unidade;   // congela a unidade desta consulta
    const dadosSupabase = await fetchSupabase(endpoint, mes, ano);
    if (dadosSupabase) {
        console.log(`✅ Dados do Supabase para ${endpoint}:`, dadosSupabase);
        return { status: 'success', data: dadosSupabase, fonte: 'supabase' };
    }

    // Apps Script (API_URL) pertence à unidade padrão: outras franquias nunca caem nele
    // (evita exibir dados de outra unidade).
    if (!unidade || unidade.id !== UNIDADE_PADRAO_ID) {
        return { status: 'error', error: 'Ainda não há dados sincronizados para este período (ou houve falha de conexão). Tente novamente em instantes.' };
    }

    const url = buildUrl(endpoint, mes, ano);
    try {
        const resp = await fetch(url);
        if (!resp.ok) {
            console.warn(`[Apps Script] HTTP ${resp.status} para endpoint "${endpoint}"`);
            return { status:'error', error:`Servidor retornou erro ${resp.status}. Tente novamente em instantes.` };
        }
        const json = await resp.json();
        if (json.status === 'success') {
            json.fonte = 'appscript';
            console.log(`✅ Dados do Apps Script para ${endpoint}:`, json.data);
        }
        return json;
    } catch (err) {
        const isOffline = !navigator.onLine || err instanceof TypeError;
        console.error(`[Apps Script] Falha no endpoint "${endpoint}":`, err);
        return {
            status: 'error',
            error: isOffline ? 'Sem conexão com a internet.' : 'Não foi possível conectar ao servidor de dados.'
        };
    }
}

function buildUrl(ep,m,y){ let u=`${API_URL}?endpoint=${ep}`; if(!SEM_FILTRO.includes(ep))u+=`&mes=${m}&ano=${y}`; return u; }

// ── Auto-refresh silencioso a cada 30min ──────────────────────────────────
setInterval(() => {
    if (!AUTH.unidade) return;   // deslogado: nada a atualizar
    if (state.currentDashboard === 'equipe') return;   // tela administrativa não usa cache de dados
    const k = cacheKey(state.currentDashboard, state.currentMonth, state.currentYear);
    invalidateCache(k);
    silentRefresh();
}, CACHE_TTL);

export async function silentRefresh(){
    // Se o usuário trocar de tela/unidade enquanto buscamos, loadSeq muda e o resultado é descartado.
    const seq = state.loadSeq;
    try{
        if(state.currentDashboard === 'resumo'){
            const eps = ['documentacao','app','adimplencia'];
            const results = {};
            await Promise.all(eps.map(async ep => {
                const k = cacheKey(ep, state.currentMonth, state.currentYear);
                const r = await fetchData(ep, state.currentMonth, state.currentYear);
                if(r.status==='success'){ setCache(k,r.data); results[ep]=r.data; }
            }));
            if(seq !== state.loadSeq) return;
            if(results['documentacao']) renderResumoDashboard(results['documentacao'], results['app'], results['adimplencia']);
        } else {
            const k = cacheKey(state.currentDashboard, state.currentMonth, state.currentYear);
            const r = await fetchData(state.currentDashboard, state.currentMonth, state.currentYear);
            if(r.status==='success'){
                setCache(k,r.data);
                if(seq !== state.loadSeq) return;
                state.data = r.data;
                hooks.renderDashboard();
            }
        }
        updateLastUpdateTime();
    }catch(_){}
}
