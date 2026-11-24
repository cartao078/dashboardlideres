// ============================================================================
// ui.js — helpers visuais (cards, métricas, gráficos), loading/erros, modal de
// período e EXPORTAÇÃO de imagens (página inteira e bloco a bloco).
// ============================================================================
import { state, hooks, CH, SEM_FILTRO } from './config.js';
import { getPercentClass, getMonthName } from './utils.js';

// ── Acesso ao container principal ─────────────────────────────────────────
export function contentEl(){ return state.dom.dashboardContent; }

// ── Gráficos (Chart.js é carregado via <script> clássico → global `Chart`) ──
export const chartInstances = {};
export function destroyChart(id){ if(chartInstances[id]){chartInstances[id].destroy();delete chartInstances[id];} }
export function createChart(id,cfg){ destroyChart(id); const ctx=document.getElementById(id); if(!ctx)return; chartInstances[id]=new Chart(ctx,cfg); }

export function destroyAllCharts(){
    Object.keys(chartInstances).forEach(id => destroyChart(id));
}

// Tooltips escuros modernos em TODOS os gráficos (defaults globais do Chart.js)
(function aplicarTemaCharts(){
    if (typeof Chart === 'undefined') return;
    const t = Chart.defaults.plugins.tooltip;
    t.backgroundColor = '#0f172a';
    t.padding         = 10;
    t.cornerRadius    = 8;
    t.titleFont       = { weight: '700', size: 12 };
    t.bodyFont        = { size: 12 };
    t.boxPadding      = 4;
    t.caretSize       = 5;
    t.usePointStyle   = true;
    Chart.defaults.plugins.legend.labels.usePointStyle = true;
})();

// ── Cards / métricas ──────────────────────────────────────────────────────
export function kpiCard(label, value, icon, cor, sub){
    return `<div class="kpi-card"><div class="kpi-icon" style="background:${cor}20;color:${cor}"><i class="${icon}"></i></div><div class="kpi-body"><div class="kpi-label">${label}</div><div class="kpi-value" style="color:${cor}">${value}</div>${sub?`<div class="kpi-sub">${sub}</div>`:''}</div></div>`;
}

export function metricItem(l,v,c){ const s=c?`style="color:${c};"`:''; return `<div class="metric-item"><div class="metric-label">${l}</div><div class="metric-value" ${s}>${v??0}</div></div>`; }
export function metricPercent(l,v){ const cls=getPercentClass(typeof v==='number'?v:parseInt(v)); return `<div class="metric-item"><div class="metric-label">${l}</div><div class="metric-percent ${cls}">${v}%</div></div>`; }
export function metricItemWhite(l,v){ return `<div class="metric-item"><div class="metric-label" style="color:rgba(255,255,255,0.75)">${l}</div><div class="metric-value" style="color:white">${v}</div></div>`; }

// ── Opções de gráficos ────────────────────────────────────────────────────
export function scalesXY(){
    return {
        x:{ grid:{display:false}, border:{display:false}, ticks:{color:CH.text} },
        y:{ grid:{color:CH.grid}, border:{display:false}, ticks:{color:CH.text, padding:6}, beginAtZero:true }
    };
}
export function legendTop(){ return {position:'top', align:'end', labels:{boxWidth:10, padding:14, usePointStyle:true}}; }

export function barDS(label, data, color){
    return {
        label, data, backgroundColor: color,
        borderRadius: { topLeft: 6, topRight: 6, bottomLeft: 0, bottomRight: 0 },
        borderSkipped: false,
        maxBarThickness: 38
    };
}

export function tooltipDark(){
    return { backgroundColor:'#0f172a', padding:10, cornerRadius:8, titleFont:{weight:'700',size:12}, bodyFont:{size:12}, boxPadding:4, usePointStyle:true };
}

export function doughnutOpts(){
    return { responsive:true, maintainAspectRatio:false, cutout:'72%',
        plugins:{ legend:{ position:'bottom', labels:{ boxWidth:10, padding:14, usePointStyle:true } }, tooltip: tooltipDark() } };
}

// ── Card de vendedor "pro" ────────────────────────────────────────────────
// Meta de aprovação da unidade (definida em aplicarUnidade → state.metaAprovacao)
export function metaAprovacao(){
    const m = Number(state.metaAprovacao);
    return Number.isFinite(m) && m > 0 ? m : 90;
}

// Tom da barra: verde ≥ meta · âmbar até 10 pts abaixo · vermelho abaixo disso
export function progressTone(p){
    const m = metaAprovacao();
    return p >= m ? 'ok' : p >= m - 10 ? 'warn' : 'bad';
}

/**
 * Card de vendedor "pro":
 *  - 3 indicadores no topo (2 informados + % atingido, colorido pelo tom)
 *  - barra de progresso com marcador da meta
 *  - detalhes secundários em chips compactos
 *
 * opts: { nome, badge(html), cardClass, borderColor, hero:[{label,value,color}],
 *         pct, pctLabel, details:[{label,value,color}], extra(html) }
 */
export function consultantCardPro(o){
    const pct  = Math.max(0, Math.min(100, Math.round(Number(o.pct) || 0)));
    const meta = metaAprovacao();
    const tone = progressTone(pct);

    const hero = (o.hero || []).map(h => `
        <div class="cpro-stat">
            <div class="cpro-stat-label">${h.label}</div>
            <div class="cpro-stat-value"${h.color ? ` style="color:${h.color}"` : ''}>${h.value ?? 0}</div>
        </div>`).join('');

    const details = (o.details || []).map(d => `
        <span class="cpro-chip">${d.label}<b${d.color ? ` style="color:${d.color}"` : ''}>${d.value ?? 0}</b></span>`).join('');

    return `
    <div class="consultant-card cpro ${o.cardClass || ''}"${o.borderColor ? ` style="border-left-color:${o.borderColor}"` : ''}>
        <div class="consultant-header">
            <div class="consultant-name">${o.nome || 'Sem nome'}</div>
            ${o.badge || ''}
        </div>
        <div class="cpro-hero">
            ${hero}
            <div class="cpro-stat">
                <div class="cpro-stat-label">${o.pctLabel || '% Atingido'}</div>
                <div class="cpro-stat-value tone-${tone}">${pct}%</div>
            </div>
        </div>
        <div class="cpro-progress" title="Meta da unidade: ${meta}%">
            <div class="cpro-progress-head"><span>Meta ${meta}%</span><b class="cpro-progress-pct">${pct}%</b></div>
            <div class="cpro-track">
                <div class="cpro-fill tone-${tone}" style="width:${pct}%"></div>
                <div class="cpro-meta-mark" style="left:${Math.min(100, meta)}%"></div>
            </div>
        </div>
        ${details ? `<div class="cpro-details">${details}</div>` : ''}
        ${o.extra || ''}
    </div>`;
}

// Detalhes padrão dos cards de documentação (Cancelados, Base líquida, ...)
export function docDetails(c){
    return [
        { label: 'Cancelados',   value: c.cancelados || 0, color: '#b91c1c' },
        { label: 'Base líquida', value: (c.total || 0) - (c.cancelados || 0) },
        { label: 'Pendências',   value: c.pendencias || 0, color: 'var(--danger)' },
        { label: 'Não enviado',  value: c.naoEnviado || 0, color: '#b45309' },
        { label: 'Expirado',     value: c.expirado   || 0 }
    ];
}

// ── Loading / toast / erro ────────────────────────────────────────────────
export function showSkeleton(){
    if(state.dom.loadingEl) state.dom.loadingEl.style.display='none';
    let sk = document.getElementById('skeletonLoader');
    if(sk) return;
    sk = document.createElement('div');
    sk.id = 'skeletonLoader';
    sk.style.cssText = 'position:absolute;inset:0;background:rgba(240,247,242,0.85);z-index:10;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px;backdrop-filter:blur(2px);border-radius:inherit';
    sk.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;gap:12px">
            <div style="width:44px;height:44px;border:3px solid #d1e8d9;border-top-color:#00a651;border-radius:50%;animation:spin .8s linear infinite"></div>
            <span style="font-size:13px;font-weight:600;color:#5a7a65" id="skeletonMsg">Carregando dados...</span>
        </div>
        <div style="display:flex;flex-direction:column;gap:10px;width:min(420px,90%)">
            <div class="sk-line" style="height:20px;border-radius:8px;width:60%"></div>
            <div class="sk-line" style="height:100px;border-radius:12px"></div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px">
                <div class="sk-line" style="height:80px;border-radius:12px"></div>
                <div class="sk-line" style="height:80px;border-radius:12px"></div>
                <div class="sk-line" style="height:80px;border-radius:12px"></div>
            </div>
            <div class="sk-line" style="height:60px;border-radius:12px"></div>
        </div>`;
    contentEl().style.position = 'relative';
    contentEl().appendChild(sk);
}

export function hideSkeleton(){
    const sk = document.getElementById('skeletonLoader');
    if(sk) sk.remove();
    if(state.dom.loadingEl) state.dom.loadingEl.style.display='none';
}

export function hideLoading(){ hideSkeleton(); }

export function showToast(msg){
    let t=document.getElementById('toast');
    if(!t){ t=document.createElement('div'); t.id='toast'; document.body.appendChild(t); }
    t.textContent=msg; t.className='toast toast-show';
    setTimeout(()=>{ t.className='toast'; },2500);
}

export function showError(msg, podeRetentar=true){
    hideSkeleton();
    const online = navigator.onLine;
    const iconeCor = online ? '#f59e0b' : '#ef4444';
    const icone    = online ? 'fa-exclamation-triangle' : 'fa-wifi';
    const dica     = online ? 'Os dados podem estar indisponíveis no momento.' : 'Verifique sua conexão com a internet.';

    contentEl().innerHTML = `
    <div class="error-message">
        <div style="width:64px;height:64px;border-radius:50%;background:${iconeCor}1a;display:flex;align-items:center;justify-content:center;margin:0 auto 16px">
            <i class="fas ${icone}" style="font-size:1.8rem;color:${iconeCor}"></i>
        </div>
        <h3 style="margin-bottom:8px;font-size:1.1rem">Não foi possível carregar</h3>
        <p style="color:var(--text-muted);margin-bottom:4px">${msg}</p>
        <p style="color:var(--text-muted);font-size:0.82rem;margin-bottom:20px">${dica}</p>
        <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
            ${podeRetentar ? `<button class="btn" id="errRetryBtn" style="background:#ef4444;color:white;padding:10px 20px"><i class="fas fa-redo"></i> Tentar novamente</button>` : ''}
            <button class="btn btn-topbar-refresh" id="errHomeBtn"><i class="fas fa-home"></i> Voltar ao início</button>
        </div>
    </div>`;

    // Módulos ES não expõem funções ao onclick="" inline — os eventos são ligados aqui.
    const retry = document.getElementById('errRetryBtn');
    if (retry) retry.addEventListener('click', () => hooks.loadDashboard());
    const home = document.getElementById('errHomeBtn');
    if (home) home.addEventListener('click', () => { state.currentDashboard = 'resumo'; hooks.loadDashboard(); });
}

export function updateLastUpdateTime(){
    if(state.dom.lastUpdateEl) state.dom.lastUpdateEl.textContent = new Date().toLocaleString('pt-BR');
}

// ── Modal de período para mobile ───────────────────────────────────────────
export function initPeriodModal(){
    const fab = document.createElement('button');
    fab.id = 'periodFab';
    fab.innerHTML = '<i class="fas fa-calendar-alt"></i>';
    fab.title = 'Selecionar período';
    document.body.appendChild(fab);

    const modal = document.createElement('div');
    modal.id = 'periodModal';
    modal.innerHTML = `
        <div class="period-modal-backdrop" id="periodModalBackdrop"></div>
        <div class="period-modal-sheet">
            <div class="period-modal-handle"></div>
            <div class="period-modal-title"><i class="fas fa-calendar-alt" style="color:var(--primary)"></i> Selecionar Período</div>
            <div class="period-modal-body">
                <div class="select-group" style="width:100%"><label>Mês</label>
                    <select id="monthSelectMobile">
                        <option value="1">Janeiro</option><option value="2">Fevereiro</option>
                        <option value="3">Março</option><option value="4">Abril</option>
                        <option value="5">Maio</option><option value="6">Junho</option>
                        <option value="7">Julho</option><option value="8">Agosto</option>
                        <option value="9">Setembro</option><option value="10">Outubro</option>
                        <option value="11">Novembro</option><option value="12">Dezembro</option>
                    </select>
                </div>
                <div class="select-group" style="width:100%"><label>Ano</label><select id="yearSelectMobile"></select></div>
                <button class="btn btn-topbar-refresh" id="applyPeriodBtn" style="width:100%;justify-content:center;padding:12px"><i class="fas fa-check"></i> Aplicar</button>
            </div>
        </div>`;
    document.body.appendChild(modal);

    const anoAtual = new Date().getFullYear();
    const yearMob = document.getElementById('yearSelectMobile');
    for(let y = anoAtual; y >= anoAtual - 3; y--){
        const opt = document.createElement('option');
        opt.value = y; opt.textContent = y;
        yearMob.appendChild(opt);
    }

    function syncMobile(){
        const mm = document.getElementById('monthSelectMobile');
        const ym = document.getElementById('yearSelectMobile');
        if(mm) mm.value = state.currentMonth;
        if(ym) ym.value  = state.currentYear;
    }
    function openModal(){ syncMobile(); modal.classList.add('open'); document.body.style.overflow = 'hidden'; }
    function closeModal(){ modal.classList.remove('open'); document.body.style.overflow = ''; }

    fab.addEventListener('click', openModal);
    document.getElementById('periodModalBackdrop').addEventListener('click', closeModal);
    document.getElementById('applyPeriodBtn').addEventListener('click', () => {
        const mm = document.getElementById('monthSelectMobile');
        const ym = document.getElementById('yearSelectMobile');
        state.currentMonth = parseInt(mm.value);
        state.currentYear  = parseInt(ym.value);
        if(state.dom.monthSelect) state.dom.monthSelect.value = state.currentMonth;
        if(state.dom.yearSelect)  state.dom.yearSelect.value  = state.currentYear;
        closeModal();
        hooks.loadDashboard();
    });

    updateFabVisibility();
}

// Esconde o botão flutuante de período nas telas sem filtro (recorrência vendedor, campanha 14°, equipe)
export function updateFabVisibility(){
    const fab = document.getElementById('periodFab');
    if (fab) fab.classList.toggle('fab-hidden', SEM_FILTRO.includes(state.currentDashboard));
}
// ── Exportação de imagens (WhatsApp-friendly) ──────────────────────────────
// Largura fixa de captura: o print fica sempre com o mesmo enquadramento e a
// mesma legibilidade, independente do tamanho da janela de quem exporta.
const EXPORT_WIDTH = 1100;
const EXPORT_SCALE = 2;
// Navegadores limitam a altura de um <canvas> (~16k px em vários deles). Acima disso, cai para escala 1.
const EXPORT_MAX_LIVE_HEIGHT = 7500;

// Botão discreto de câmera: vai no cabeçalho de cada .sector-card e .ranking-full
export function snapBtn(){
    return '<button type="button" class="btn-snap-setor" title="Baixar imagem deste bloco" aria-label="Baixar imagem deste bloco"><i class="fas fa-camera"></i></button>';
}

// Ajusta o clone que o html2canvas renderiza (o DOM ao vivo não é tocado)
function prepararClone(doc){
    // Sem a margem da sidebar o conteúdo ocupa exatamente EXPORT_WIDTH
    const mw = doc.querySelector('.main-wrapper');
    if (mw){ mw.style.marginLeft = '0'; mw.style.width = EXPORT_WIDTH + 'px'; }
    const dc = doc.querySelector('.dashboard-content');
    if (dc) dc.style.overflow = 'visible';

    // Controles que não devem aparecer na imagem
    doc.querySelectorAll('.btn-snap-setor, #skeletonLoader').forEach(n => n.remove());

    // Gráficos: o bitmap do <canvas> tem proporção fixa. Se o box do clone tiver outra proporção,
    // o html2canvas esticaria a imagem. Trocamos por <img> com a proporção original, centralizada.
    doc.querySelectorAll('canvas').forEach(cv => {
        const vivo = cv.id ? document.getElementById(cv.id) : null;
        if (!vivo || !vivo.width || !vivo.height) return;
        let url;
        try { url = vivo.toDataURL('image/png'); } catch (_) { return; }
        const pai = cv.parentElement;
        const pw = pai ? pai.clientWidth  : 0;
        const ph = pai ? pai.clientHeight : 0;
        const ratio = vivo.width / vivo.height;
        let w = pw || vivo.width, h = w / ratio;
        if (ph && h > ph){ h = ph; w = h * ratio; }
        const img = doc.createElement('img');
        img.src = url;
        img.style.cssText = `display:block;margin:0 auto;width:${Math.round(w)}px;height:${Math.round(h)}px`;
        cv.replaceWith(img);
    });
}

async function capturar(el){
    const alturaViva = el.scrollHeight || 0;
    const scale = alturaViva > EXPORT_MAX_LIVE_HEIGHT ? 1 : EXPORT_SCALE;
    return html2canvas(el, {
        scale,
        useCORS: true,
        backgroundColor: '#ffffff',
        scrollX: 0,
        scrollY: -window.scrollY,
        windowWidth: EXPORT_WIDTH,
        windowHeight: alturaViva || undefined,
        onclone: prepararClone
    });
}

function baixarCanvas(canvas, nomeArquivo){
    return new Promise((resolve, reject) => {
        canvas.toBlob(blob => {
            if (!blob) { reject(new Error('Falha ao gerar o PNG')); return; }
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.download = nomeArquivo;
            link.href = url;
            document.body.appendChild(link);
            link.click();
            link.remove();
            setTimeout(() => URL.revokeObjectURL(url), 5000);
            resolve();
        }, 'image/png');
    });
}

function slug(txt){
    return String(txt || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
        .replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '').toUpperCase();
}

function sufixoPeriodo(){
    return SEM_FILTRO.includes(state.currentDashboard) ? '' : `_${getMonthName(state.currentMonth)}-${state.currentYear}`;
}

// Botão geral (topbar): a página inteira
export async function exportPage(){
    const btn = document.getElementById('downloadBtn');
    const orig = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando...';
    btn.disabled = true;
    try{
        const container = document.querySelector('.main-wrapper');
        const canvas = await capturar(container);
        const nome = state.currentDashboard.toUpperCase().replace('_','-');
        await baixarCanvas(canvas, `Dashboard_${nome}${sufixoPeriodo()}.png`);
    }catch(err){ console.error('[export] página:', err); showToast('❌ Erro ao gerar imagem.'); }
    finally{ btn.innerHTML = orig; btn.disabled = false; }
}

// Câmera de cada bloco (.sector-card / .ranking-full)
function nomeDoBloco(bloco){
    if (bloco.classList.contains('ranking-full')) return 'RANKING';
    const titulo = bloco.querySelector('.sector-title');
    if (!titulo) return 'BLOCO';
    const copia = titulo.cloneNode(true);
    copia.querySelectorAll('.sector-count').forEach(n => n.remove());
    return slug(copia.textContent) || 'BLOCO';
}

export async function exportBloco(bloco, btn){
    const icone = btn ? btn.querySelector('i') : null;
    if (btn) btn.disabled = true;
    if (icone) icone.className = 'fas fa-spinner fa-spin';
    try{
        const canvas = await capturar(bloco);
        const dash = state.currentDashboard.toUpperCase().replace('_','-');
        await baixarCanvas(canvas, `Dashboard_${dash}_${nomeDoBloco(bloco)}${sufixoPeriodo()}.png`);
    }catch(err){ console.error('[export] bloco:', err); showToast('❌ Erro ao gerar imagem.'); }
    finally{
        if (icone) icone.className = 'fas fa-camera';
        if (btn) btn.disabled = false;
    }
}

// Um único listener (delegação) cobre todos os blocos, mesmo os re-renderizados
export function initSnapButtons(){
    contentEl().addEventListener('click', ev => {
        const btn = ev.target.closest('.btn-snap-setor');
        if (!btn) return;
        const bloco = btn.closest('.sector-card, .ranking-full');
        if (bloco) exportBloco(bloco, btn);
    });
}
