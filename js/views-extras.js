// ============================================================================
// views-extras.js — Adimplência, Recorrência, Rec. Vendedor, Refuturiza,
// Vendas TIM e Campanha 14°
// ============================================================================
import { C } from './config.js';
import {
    contentEl, showError, metricItem, metricPercent, metricItemWhite, consultantCardPro, snapBtn
} from './ui.js';
import { calcPercent, getPercentClass, getSectorClass } from './utils.js';

export function renderAdimplenciaDashboard(d){
    const {geral,consultores,mes,ano}=d;
    contentEl().innerHTML=`
    <h2 class="dash-title"><i class="fas fa-credit-card" style="color:var(--success)"></i> Dashboard Adimplência — ${mes} ${ano}</h2>
    <div class="card card-adim" style="max-width:600px;margin:0 auto 28px;"><div class="card-header"><div class="card-title">Adimplência — Total da Loja</div><div class="card-icon"><i class="fas fa-chart-line"></i></div></div><div class="metric-grid">${metricItem('Total Trocas',geral.totalTrocas)}${metricItem('Mens. OK',geral.mensOk,'var(--success)')}${metricItem('Mens. Aberto',geral.mensAberto,'var(--warning)')}${metricItem('Mens. Atraso',geral.mensAtraso,'var(--danger)')}${metricItem('Aprovados',geral.aprovados,'var(--success)')}${metricItem('Pendentes',geral.pendentes,'var(--danger)')}${metricItem('Total BI',geral.totalBi)}${metricItem('Fora BI',geral.foraBi,'var(--danger)')}${metricPercent('% Aprovados',geral.percentualAprovado)}</div></div>
    <h3 class="section-title"><i class="fas fa-user-tie" style="color:var(--primary)"></i> Desempenho por Consultor</h3>
    <div class="table-wrapper"><table class="data-table"><thead><tr><th>Consultor</th><th>Total Trocas</th><th>Mens. OK</th><th>Mens. Aberto</th><th>Mens. Atraso</th><th>Aprovados</th><th>Pendentes</th><th>Total BI</th><th>Fora BI</th><th>% Aprovados</th></tr></thead><tbody>${consultores.map(c=>`<tr><td><strong>${c.nome}</strong></td><td>${c.totalTrocas}</td><td style="color:var(--success)">${c.mensOk}</td><td style="color:var(--warning)">${c.mensAberto}</td><td style="color:var(--danger)">${c.mensAtraso}</td><td style="color:var(--success)">${c.aprovados}</td><td style="color:var(--danger)">${c.pendentes}</td><td>${c.totalBi}</td><td style="color:var(--danger)">${c.foraBi}</td><td><span class="metric-percent ${getPercentClass(c.percentualAprovado)}">${c.percentualAprovado}%</span></td></tr>`).join('')}</tbody></table></div>`;
}

export function renderRecorrenciaDashboard(d){
    if (!d || !d.geral || !d.consultores) {
        showError('Dados de Qualidade da Retenção não encontrados.');
        return;
    }

    const { geral, consultores, periodo } = d;
    const nomes = Object.keys(consultores).sort();

    function cardConsultor(nome) {
        const c = consultores[nome];
        return consultantCardPro({
            nome: c.nome,
            borderColor: 'var(--teal)',
            badge: '<div class="consultant-sector sector-retencao">RETENÇÃO</div>',
            hero: [
                { label: 'Retenção',  value: c.totalRetencao },
                { label: 'Com termo', value: c.totalComTermo, color: 'var(--success)' }
            ],
            pct: c.percentualOK || 0,
            pctLabel: '% OK',
            details: [
                { label: 'Sem termo', value: c.totalSemTermo, color: 'var(--danger)' }
            ]
        });
    }

    contentEl().innerHTML = `
    <h2 class="dash-title"><i class="fas fa-clipboard-check" style="color:var(--teal)"></i> Qualidade da Retenção — ${periodo.atual}</h2>

    <div style="background:rgba(13,148,136,0.07);padding:14px 18px;border-radius:12px;margin-bottom:24px;border:1px solid rgba(13,148,136,0.2);border-left:4px solid var(--teal);">
        <p style="margin:0;color:#0f766e;font-weight:600;font-size:0.9rem;">
            <i class="fas fa-info-circle"></i> Período: <strong>${periodo.atual}</strong>
        </p>
    </div>

    <h3 class="section-title" style="color:var(--teal)">
        <i class="fas fa-globe" style="color:var(--teal)"></i> Resumo Geral
    </h3>

    <div class="card" style="max-width: 600px; margin: 0 auto 36px; border-top: 4px solid var(--teal);">
        <div class="card-header">
            <div class="card-title" style="color: var(--teal)">Qualidade Global</div>
            <div class="card-icon" style="background: rgba(13,148,136,0.12); color: var(--teal)"><i class="fas fa-check-double"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItem('Total Retenção', geral.totalRetencao)}
            ${metricItem('Com Termo (SIM)', geral.totalComTermo, 'var(--success)')}
            ${metricItem('Sem Termo', geral.totalSemTermo, 'var(--danger)')}
            ${metricPercent('% OK Geral', geral.percentualOK)}
        </div>
    </div>

    ${nomes.length > 0 ? `
    <h3 class="section-title" style="color:var(--primary)">
        <i class="fas fa-users" style="color:var(--primary)"></i> Detalhamento por Consultor
    </h3>
    <div class="consultant-grid">
        ${nomes.map(nome => cardConsultor(nome)).join('')}
    </div>
    ` : `
    <div style="text-align:center;padding:40px;color:var(--gray)">
        <i class="fas fa-info-circle" style="font-size:2.5rem;margin-bottom:16px;display:block;color:var(--teal);"></i>
        <h3>Nenhum consultor com dados para ${periodo.atual}</h3>
    </div>`}
    `;
}

export function renderRecorrenciaVendedorDashboard(d){
    const {geral,consultores,totalConsultores,totalRegistros}=d;
    if(!geral||!consultores){showError('Dados de Recorrência Vendedor não encontrados');return;}
    const pG=calcPercent(geral.totalOk+geral.totalEmAberto,geral.totalVendasPromocao);
    const setores=[{nome:'VENDAS',cor:C.green,icone:'fas fa-shopping-cart'},{nome:'REFILIACAO',cor:C.warn,icone:'fas fa-user-plus'},{nome:'RECEPCAO',cor:C.teal,icone:'fas fa-headset'}];

    const vendCard = c => consultantCardPro({
        nome: c.nome,
        badge: `<div class="consultant-sector ${getSectorClass(c.setor)}">${c.setor}</div>`,
        hero: [
            { label: 'Vendas', value: c.totalVendasPromocao },
            { label: 'OK',     value: c.totalOk || 0, color: 'var(--success)' }
        ],
        pct: calcPercent((c.totalOk||0)+(c.totalEmAberto||0), c.totalVendasPromocao),
        pctLabel: '% OK+Aberto',
        details: [
            { label: 'Em aberto', value: c.totalEmAberto || 0, color: '#b45309' },
            { label: 'Atraso',    value: c.totalAtraso || 0,   color: 'var(--danger)' },
            { label: 'Outros',    value: c.totalOutros || 0 }
        ]
    });

    contentEl().innerHTML=`
    <h2 class="dash-title"><i class="fas fa-handshake" style="color:var(--teal)"></i> Recorrência Vendedor — Dados Completos</h2>
    <div style="background:rgba(13,148,136,0.07);padding:14px 18px;border-radius:12px;margin-bottom:24px;border:1px solid rgba(13,148,136,0.2);border-left:4px solid var(--teal)"><p style="margin:0;color:#065f46;font-weight:600;font-size:0.9rem"><i class="fas fa-database"></i> <strong>${totalRegistros||0}</strong> registros | <strong>${totalConsultores}</strong> consultores</p></div>
    <div class="main-cards"><div class="card" style="border-top:3px solid var(--teal)"><div class="card-header"><div class="card-title">Total Geral</div><div class="card-icon" style="background:#e0f2f1;color:var(--teal)"><i class="fas fa-chart-bar"></i></div></div><div class="metric-grid">${metricItem('Total Vendas',geral.totalVendasPromocao)}${metricItem('OK',geral.totalOk,'var(--success)')}${metricItem('Em Aberto',geral.totalEmAberto,'var(--warning)')}${metricItem('Atraso',geral.totalAtraso,'var(--danger)')}${metricItem('Outros',geral.totalOutros||0,'var(--gray)')}${metricPercent('% OK+Aberto',pG)}</div></div></div>
    ${setores.map(si=>{const lista=consultores.filter(c=>c.setor===si.nome);if(!lista.length)return '';const tot=lista.reduce((s,c)=>s+c.totalVendasPromocao,0),ok=lista.reduce((s,c)=>s+(c.totalOk||0)+(c.totalEmAberto||0),0);const pct=calcPercent(ok,tot);return `<div class="sector-card"><div class="sector-header"><div class="sector-title"><i class="${si.icone}" style="color:${si.cor}"></i><span style="color:${si.cor}">${si.nome}</span><span class="sector-count">${lista.length} consultor${lista.length!==1?'es':''}</span></div><div class="sector-header-right"><div class="metric-percent ${getPercentClass(pct)}">${pct}% OK+Aberto</div>${snapBtn()}</div></div><div class="consultant-grid">${lista.sort((a,b)=>b.totalVendasPromocao-a.totalVendasPromocao).map(vendCard).join('')}</div></div>`;}).join('')}
    <div class="card" style="margin-top:36px;background:linear-gradient(135deg,var(--primary),var(--primary-dark));color:white;border:none"><div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2)"><div class="card-title" style="color:white">📊 Resumo Final</div><div class="card-icon" style="background:rgba(255,255,255,0.2);color:white"><i class="fas fa-clipboard-list"></i></div></div><div class="metric-grid">${metricItemWhite('Total Consultores',totalConsultores||consultores.length)}${metricItemWhite('Total Vendas',geral.totalVendasPromocao)}${metricItemWhite('Média por Consultor',consultores.length>0?Math.round(geral.totalVendasPromocao/consultores.length):0)}${metricItemWhite('Melhor %',consultores.length>0?Math.max(...consultores.map(c=>calcPercent((c.totalOk||0)+(c.totalEmAberto||0),c.totalVendasPromocao)))+'%':'0%')}</div></div>`;
}

export function renderRefuturizaDashboard(d) {
    if (!d || !d.geral) { showError('Dados do Refuturiza não encontrados'); return; }

    let consultoresList = Array.isArray(d.consultores) ? d.consultores : [];

    const totalBruto = Number(d.geral.total) || 0;
    const cancelados = Number(d.geral.cancelado) || 0;
    const baseLiquida = totalBruto - cancelados;
    const pG = baseLiquida > 0 ? Math.round((Number(d.geral.comLigacao) / baseLiquida) * 100) : 0;

    const consultoresComVendas = consultoresList.filter(c => (Number(c.total) || 0) > 0);
    let consultoresParaExibir = consultoresComVendas;
    if (consultoresComVendas.length === 0 && totalBruto > 0) {
        consultoresParaExibir = [{ nome:'📊 Total Geral', total: totalBruto, cancelado: cancelados, comLigacao: Number(d.geral.comLigacao)||0, semLigacao: Number(d.geral.semLigacao)||0 }];
    }

    const consultoresOrdenados = [...consultoresParaExibir].sort((a, b) => {
        const baseA = (Number(a.total)||0) - (Number(a.cancelado)||0);
        const baseB = (Number(b.total)||0) - (Number(b.cancelado)||0);
        return baseB - baseA;
    });

    contentEl().innerHTML = `
    <h2 class="dash-title"><i class="fas fa-book" style="color:var(--accent-dark)"></i> Dashboard Refuturiza — ${d.mes || 'Mês'} ${d.ano || ''}</h2>

    <div class="card card-refut" style="max-width:500px;margin:0 auto 28px">
        <div class="card-header">
            <div class="card-title">Refuturiza — Total da Loja</div>
            <div class="card-icon"><i class="fas fa-book-open"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItem('TOTAL BRUTO', totalBruto)}
            ${metricItem('CANCELADOS', cancelados, '#991b1b')}
            ${metricItem('BASE LÍQUIDA', baseLiquida, 'var(--primary)')}
            ${metricItem('COM LIGAÇÃO', d.geral.comLigacao || 0, 'var(--success)')}
            ${metricItem('SEM LIGAÇÃO', d.geral.semLigacao || 0, 'var(--danger)')}
            ${metricPercent('% COM LIGAÇÃO (s/ cancelados)', pG)}
        </div>
    </div>

    ${consultoresOrdenados.length > 0 ? `
    <h3 class="section-title"><i class="fas fa-users" style="color:var(--accent-dark)"></i> Desempenho por Consultor (${consultoresOrdenados.length})</h3>
    <div class="consultant-grid">
        ${consultoresOrdenados.map(c => {
            const cTotal = Number(c.total) || 0;
            const cCancel = Number(c.cancelado) || 0;
            const cBase = cTotal - cCancel;
            const p = cBase > 0 ? Math.round(((Number(c.comLigacao) || 0) / cBase) * 100) : 0;
            return consultantCardPro({
                nome: c.nome || 'Desconhecido',
                badge: '<div class="consultant-sector" style="background:var(--primary-light);color:var(--accent-dark)">REFUTURIZA</div>',
                hero: [
                    { label: 'Base líquida', value: cBase },
                    { label: 'Com ligação',  value: c.comLigacao || 0, color: 'var(--success)' }
                ],
                pct: p,
                pctLabel: '% Com ligação',
                details: [
                    { label: 'Total bruto',  value: cTotal },
                    { label: 'Cancelados',   value: cCancel, color: '#b91c1c' },
                    { label: 'Sem ligação',  value: c.semLigacao || 0, color: 'var(--danger)' }
                ]
            });
        }).join('')}
    </div>

    <div class="card" style="margin-top:36px;background:linear-gradient(135deg,var(--primary),var(--primary-dark));color:white;border:none">
        <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2)">
            <div class="card-title" style="color:white">📊 Resumo Final</div>
            <div class="card-icon" style="background:rgba(255,255,255,0.2);color:white"><i class="fas fa-graduation-cap"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItemWhite('Total Consultores c/ vendas', consultoresOrdenados.length)}
            ${metricItemWhite('Total Bruto', totalBruto)}
            ${metricItemWhite('Total Cancelados', cancelados)}
            ${metricItemWhite('Base Líquida', baseLiquida)}
            ${metricItemWhite('Média por Consultor', consultoresOrdenados.length > 0 ? Math.round(baseLiquida / consultoresOrdenados.length) : 0)}
            ${metricItemWhite('Taxa de Contato (s/ cancelados)', pG + '%')}
        </div>
    </div>
    ` : `
    <div style="text-align:center;padding:40px;color:var(--gray)">
        <i class="fas fa-info-circle" style="font-size:2.5rem;margin-bottom:16px;display:block;"></i>
        <h3>Nenhum consultor com vendas para ${d.mes || 'Mês'} ${d.ano || ''}</h3>
        <p style="font-size:0.9rem;margin-top:8px;">Total bruto: ${totalBruto} | Cancelados: ${cancelados}</p>
    </div>`}`;
}

export function renderVendasTimDashboard(d) {
    if (!d || !d.geral) { showError('Dados do Vendas TIM não encontrados'); return; }

    const { geral, vendedores, mes, ano, totalVendedores } = d;
    const TIM_BLUE = '#0054a6';
    const TIM_BLUE_DARK = '#003d7a';

    const pAtivacao = geral.pctAtivacao || 0;
    const pPortabilidade = geral.pctPortabilidade || 0;
    const vendedoresOrdenados = (vendedores || []).sort((a, b) => b.total - a.total);

    contentEl().innerHTML = `
    <h2 class="dash-title" style="border-bottom-color:${TIM_BLUE};">
        <i class="fas fa-sim-card" style="color:${TIM_BLUE};"></i>
        Dashboard Vendas TIM — ${mes || 'Mês'} ${ano || ''}
    </h2>
    <div class="card card-tim" style="max-width:480px;margin:0 auto 28px;border-top:4px solid ${TIM_BLUE};">
        <div class="card-header">
            <div class="card-title" style="color:${TIM_BLUE};"><i class="fas fa-chart-bar"></i> Total Geral</div>
            <div class="card-icon"><i class="fas fa-sim-card"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItem('TOTAL VENDAS', geral.total, TIM_BLUE)}
            ${metricItem('ATIVAÇÃO', geral.ativacao, TIM_BLUE)}
            ${metricItem('PORTABILIDADE', geral.portabilidade, TIM_BLUE)}
            ${metricPercentTim('% ATIVAÇÃO', pAtivacao)}
            ${metricPercentTim('% PORTABILIDADE', pPortabilidade)}
        </div>
    </div>
    ${vendedoresOrdenados.length > 0 ? `
    <h3 class="section-title"><i class="fas fa-users" style="color:${TIM_BLUE};"></i> Vendedores (${vendedoresOrdenados.length})</h3>
    <div class="consultant-grid">
        ${vendedoresOrdenados.map(c => `
            <div class="consultant-card" style="border-left-color:${TIM_BLUE};">
                <div class="consultant-header">
                    <div class="consultant-name" style="color:${TIM_BLUE};">${c.nome}</div>
                    <div class="consultant-sector sector-tim">TIM</div>
                </div>
                <div class="metric-grid">
                    ${metricItem('TOTAL', c.total, TIM_BLUE)}
                    ${metricItem('ATIVAÇÃO', c.ativacao, TIM_BLUE)}
                    ${metricItem('PORTABILIDADE', c.portabilidade, TIM_BLUE)}
                    ${metricPercentTim('% ATIVAÇÃO', c.pctAtivacao || 0)}
                    ${metricPercentTim('% PORTABILIDADE', c.pctPortabilidade || 0)}
                </div>
            </div>`).join('')}
    </div>
    <div class="card" style="margin-top:36px;background:linear-gradient(135deg,${TIM_BLUE},${TIM_BLUE_DARK});color:white;border:none;">
        <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.2);">
            <div class="card-title" style="color:white;">📊 Resumo Final</div>
            <div class="card-icon" style="background:rgba(255,255,255,0.2);color:white;"><i class="fas fa-clipboard-list"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItemWhite('Total Vendedores', vendedoresOrdenados.length)}
            ${metricItemWhite('Total Vendas', geral.total)}
            ${metricItemWhite('Média por Vendedor', vendedoresOrdenados.length > 0 ? Math.round(geral.total / vendedoresOrdenados.length) : 0)}
            ${metricItemWhite('Melhor Vendedor', vendedoresOrdenados.length > 0 ? vendedoresOrdenados[0].nome + ' (' + vendedoresOrdenados[0].total + ')' : 'N/A')}
        </div>
    </div>
    ` : `<div style="text-align:center;padding:40px;color:var(--gray)"><i class="fas fa-info-circle" style="font-size:2.5rem;margin-bottom:16px;display:block;color:${TIM_BLUE};"></i><h3>Nenhum vendedor com vendas para ${mes || 'Mês'} ${ano || ''}</h3></div>`}`;
}

function metricPercentTim(label, v) {
    const cls = typeof v === 'number' ? getPercentClass(v) : 'percent-low';
    const cor = v >= 70 ? '#0054a6' : v >= 50 ? '#f59e0b' : '#ef4444';
    return `<div class="metric-item"><div class="metric-label">${label}</div><div class="metric-percent ${cls}" style="background:rgba(0,84,166,0.1);color:${cor}">${v}%</div></div>`;
}

export function renderCampanha14Dashboard(d) {
    if (!d || !d.resumo || !d.colaboradores) { showError('Dados da Campanha 14° não encontrados'); return; }
    const { resumo, porSetor } = d;
    const MESES    = resumo.mesesCampanha || ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov"];
    const META     = resumo.metaMeses || 9;
    const mesAtual = new Date().getMonth();

    function badgeStatus(c) {
        const faltam = META - c.totalDestaques;
        if (c.status === 'classificado')
            return `<span style="background:#dcfce7;color:#166534;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600"><i class="fas fa-check" style="font-size:9px"></i> Classificado</span>`;
        if (c.status === 'fora')
            return `<span style="background:#fee2e2;color:#991b1b;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600"><i class="fas fa-times" style="font-size:9px"></i> Fora da meta</span>`;
        return `<span style="background:#fef9c3;color:#854d0e;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600">${faltam} ${faltam===1?'mês':'meses'} p/ meta</span>`;
    }

    function mesesHtml(c) {
        return MESES.map((m, i) => {
            const dest = c.meses[i];
            const atual = i === mesAtual;
            const fut = i > mesAtual;
            let bg, color, border, icon;
            if (dest)     { bg='#dcfce7'; color='#166534'; border='#16a34a'; icon='<i class="fas fa-check" style="font-size:9px"></i>'; }
            else if (fut) { bg='#f8fafc'; color='#94a3b8'; border='#e2e8f0'; icon=m; }
            else          { bg='#fee2e2'; color='#991b1b'; border='#fca5a5'; icon=m; }
            const ring = atual ? ';outline:2px solid var(--teal);outline-offset:2px' : '';
            return `<div title="${m}" style="width:38px;height:38px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;background:${bg};color:${color};border:1px solid ${border}${ring}">${icon}</div>`;
        }).join('');
    }

    function progressBar(c) {
        const pct = Math.min(100, Math.round((c.totalDestaques / META) * 100));
        const cor = c.status==='classificado' ? '#00a651' : c.status==='fora' ? '#ef4444' : '#f59e0b';
        return `<div style="display:flex;align-items:center;gap:10px;margin-top:10px"><div style="flex:1;height:6px;background:#e5e7eb;border-radius:3px"><div style="height:6px;border-radius:3px;background:${cor};width:${pct}%;transition:.4s"></div></div><span style="font-size:11px;color:var(--gray);white-space:nowrap;min-width:80px;text-align:right">${c.totalDestaques}/${META} meses • ${pct}%</span></div>`;
    }

    function colabCard(c) {
        return `<div style="background:var(--light);border:1px solid var(--border);border-radius:12px;padding:14px 16px;margin-bottom:8px"><div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px;margin-bottom:10px;flex-wrap:wrap"><span style="font-weight:600;font-size:13px;color:var(--text);flex:1;line-height:1.3">${c.nome}</span>${badgeStatus(c)}</div><div style="display:flex;gap:5px;flex-wrap:wrap">${mesesHtml(c)}</div>${progressBar(c)}</div>`;
    }

    let setorIdx = 0;
    const setorCards = Object.entries(porSetor).map(([setor, lista]) => {
        const classif = lista.filter(c => c.status==='classificado').length;
        const comDest = lista.filter(c => c.totalDestaques > 0);
        const semDest = lista.filter(c => c.totalDestaques === 0);
        const sid = 'setor_' + (setorIdx++);

        const btnOcultos = semDest.length > 0 ? `<button onclick="var el=document.getElementById('${sid}');var btn=this;if(el.style.display==='none'){el.style.display='block';btn.innerHTML='<i class=\\'fas fa-eye-slash\\'></i> Ocultar sem destaque (${semDest.length})';}else{el.style.display='none';btn.innerHTML='<i class=\\'fas fa-eye\\'></i> Ver sem destaque (${semDest.length})';}" style="background:none;border:1px dashed var(--border);border-radius:8px;padding:6px 12px;font-size:12px;color:var(--gray);cursor:pointer;margin-top:4px;width:100%"><i class='fas fa-eye'></i> Ver sem destaque (${semDest.length})</button>` : '';

        return `
            <div class="card" style="margin-bottom:16px">
                <div class="card-header" style="padding-bottom:12px;margin-bottom:12px">
                    <div class="card-title" style="font-size:15px"><i class="fas fa-users" style="color:var(--primary)"></i> ${setor}</div>
                    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
                        ${semDest.length > 0 ? `<span style="font-size:11px;color:var(--gray)">${semDest.length} sem destaque</span>` : ''}
                        <span style="font-size:12px;color:var(--success);font-weight:600">${classif}/${lista.length} classificados</span>
                    </div>
                </div>
                ${comDest.map(colabCard).join('')}
                ${btnOcultos}
                <div id="${sid}" style="display:none;margin-top:4px">${semDest.map(colabCard).join('')}</div>
            </div>`;
    }).join('');

    contentEl().innerHTML = `
        <h2 class="dash-title"><i class="fas fa-trophy" style="color:#f59e0b"></i> Campanha 14° Salário — Jan a Nov 2026</h2>
        <div style="background:#fefce8;padding:10px 16px;border-radius:10px;margin-bottom:16px;border-left:4px solid #f59e0b">
            <p style="margin:0;color:#854d0e;font-size:12px"><i class="fas fa-info-circle"></i> <strong>Meta:</strong> destaque em pelo menos <strong>${META} de 11 meses</strong> para garantir o 14° salário.</p>
        </div>
        <div class="main-cards" style="margin-bottom:20px">
            ${metricItem('Colaboradores', resumo.totalColabs)}
            ${metricItem('Classificados', resumo.classificados, 'var(--success)')}
            ${metricItem('Em risco', resumo.emRisco, 'var(--danger)')}
            ${metricItem('No prazo', resumo.noPrazo, 'var(--warning)')}
        </div>
        ${setorCards}`;
}
