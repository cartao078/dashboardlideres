// ============================================================================
// views-vendas.js — Resumo, Vendas/Documentação e App
// ============================================================================
import { C, CH, RANKING_SETORES } from './config.js';
import {
    contentEl, kpiCard, metricItem, metricPercent, consultantCardPro, docDetails,
    createChart, barDS, legendTop, tooltipDark, scalesXY, doughnutOpts, snapBtn
} from './ui.js';
import { calcPercent, groupBySector, sortSectors, getPercentClass, getSectorClass, getSectorIcon } from './utils.js';

export function renderResumoDashboard(vendas, app, adim){
    const { geral, consultores, mes, ano } = vendas;
    const pAprov = calcPercentDoc(geral.aprovados, geral.total, geral.cancelados||0);
    const pApp   = app  ? calcPercent(app.geral.sim, app.geral.total) : null;
    const pAdim  = adim ? adim.geral.percentualAprovado : null;

    const bySector = groupBySector(consultores);
    const order    = ['VENDAS','RECEPCAO','REFILIACAO','WEB SITE','TELEVENDAS','OUTROS'];
    const sectors  = sortSectors(Object.keys(bySector), order);

    const consultoresRanking = consultores
        .filter(c => RANKING_SETORES.includes((c.setor||'').toUpperCase()))
        .sort((a,b) => b.total - a.total);

    contentEl().innerHTML = `
    <h2 class="dash-title"><i class="fas fa-chart-line" style="color:var(--primary)"></i> Resumo Geral — ${mes} ${ano}</h2>
    <div class="kpi-strip">
        ${kpiCard('Total de Vendas', geral.total, 'fas fa-shopping-bag', C.green, '')}
        ${kpiCard('% Aprovação', pAprov+'%', 'fas fa-check-circle', pAprov>=90?C.green:pAprov>=80?C.warn:C.danger, '')}
        ${app  ? kpiCard('% Com App',      pApp+'%',  'fas fa-mobile-alt',  pApp>=90?C.green:pApp>=80?C.warn:C.danger,  '') : ''}
        ${adim ? kpiCard('% Adimplência',  pAdim+'%', 'fas fa-credit-card', pAdim>=90?C.green:pAdim>=80?C.warn:C.danger,'') : ''}
    </div>
    <div class="charts-row">
        <div class="chart-card"><div class="chart-card-title"><i class="fas fa-chart-bar"></i> Vendas por Setor</div><div class="chart-card-subtitle">Aprovados e pendências de cada setor no mês</div><div class="chart-wrap"><canvas id="chRes1"></canvas></div></div>
        <div class="chart-card"><div class="chart-card-title"><i class="fas fa-chart-pie"></i> Status das Vendas</div><div class="chart-card-subtitle">Distribuição geral: aprovados, pendências, não enviado e expirado</div><div class="chart-wrap"><canvas id="chRes2"></canvas></div></div>
        ${app ? `<div class="chart-card"><div class="chart-card-title"><i class="fas fa-mobile-alt"></i> Adesão ao App por Setor</div><div class="chart-card-subtitle">Clientes com e sem app cadastrado, separado por setor</div><div class="chart-wrap"><canvas id="chRes3"></canvas></div></div>` : ''}
    </div>
    <div class="charts-row ${adim?'charts-row-2':'charts-row-1'}">
        ${adim ? `<div class="chart-card"><div class="chart-card-title"><i class="fas fa-credit-card"></i> Situação das Mensalidades</div><div class="chart-card-subtitle">Mensalidades em dia, em aberto e em atraso</div><div class="chart-wrap"><canvas id="chRes4"></canvas></div></div>` : ''}
        <div class="chart-card"><div class="chart-card-title"><i class="fas fa-store"></i> Loja vs Web/Tele</div><div class="chart-card-subtitle">Comparativo de total, aprovados e pendências entre os canais de venda</div><div class="chart-wrap"><canvas id="chRes5"></canvas></div></div>
    </div>
    ${buildRankingCompleto(consultoresRanking, mes, ano)}`;

    const barOpts = () => ({
        responsive:true, maintainAspectRatio:false,
        interaction:{ mode:'index', intersect:false },
        plugins:{ legend:legendTop(), tooltip:tooltipDark() },
        scales:scalesXY()
    });

    createChart('chRes1',{type:'bar',data:{labels:sectors,datasets:[
        barDS('Aprovados',  sectors.map(s=>bySector[s].reduce((a,c)=>a+c.aprovados,0)),  CH.green),
        barDS('Pendências', sectors.map(s=>bySector[s].reduce((a,c)=>a+c.pendencias,0)), CH.warn)
    ]},options:barOpts()});

    createChart('chRes2',{type:'doughnut',data:{labels:['Aprovados','Pendências','Não Enviado','Expirado'],datasets:[{
        data:[geral.aprovados,geral.pendencias,geral.naoEnviado,geral.expirado||0],
        backgroundColor:[CH.green,CH.warn,CH.danger,CH.gray],borderWidth:2,borderColor:'#ffffff',hoverOffset:6
    }]},options:doughnutOpts()});

    if(app){
        const appBySector = groupBySector(app.consultores.filter(c=>!c.origem||c.origem!=='retencao'));
        createChart('chRes3',{type:'bar',data:{labels:sectors,datasets:[
            barDS('Com App', sectors.map(s=>(appBySector[s]||[]).reduce((a,c)=>a+(c.sim||0),0)), CH.teal),
            barDS('Sem App', sectors.map(s=>(appBySector[s]||[]).reduce((a,c)=>a+(c.nao||0),0)), CH.danger)
        ]},options:barOpts()});
    }
    if(adim){
        createChart('chRes4',{type:'doughnut',data:{labels:['Mens. OK','Mens. Aberto','Mens. Atraso'],datasets:[{
            data:[adim.geral.mensOk,adim.geral.mensAberto,adim.geral.mensAtraso],
            backgroundColor:[CH.green,CH.warn,CH.danger],borderWidth:2,borderColor:'#ffffff',hoverOffset:6
        }]},options:doughnutOpts()});
    }
    createChart('chRes5',{type:'bar',data:{labels:['Total','Aprovados','Pendências'],datasets:[
        barDS('Loja',     [vendas.vendasLoja.total,vendas.vendasLoja.aprovados,vendas.vendasLoja.pendencias], CH.green),
        barDS('Web/Tele', [vendas.vendasWeb.total,vendas.vendasWeb.aprovados,vendas.vendasWeb.pendencias],    CH.lime)
    ]},options:barOpts()});
}


function buildRankingCompleto(consultores, mes, ano){
    if(!consultores.length) return '';
    const medal = i => i===0?'🥇':i===1?'🥈':i===2?'🥉':(i+1);
    const posClass = i => i===0?'rank-gold':i===1?'rank-silver':i===2?'rank-bronze':'rank-other';
    const maxTotal = consultores[0].total || 1;
    const porSetor = {};
    RANKING_SETORES.forEach(s => { porSetor[s] = consultores.filter(c=>(c.setor||'').toUpperCase()===s); });
    const top10 = consultores.slice(0, 10);
    return `
    <div class="ranking-full">
        <div class="ranking-full-header"><i class="fas fa-trophy"></i> 🏆 Ranking de Vendedores — ${mes} ${ano}<span class="ranking-badge">Vendas · Recepção · Refiliação</span>${snapBtn()}</div>
        <div class="ranking-section-title">Top 10 Geral</div>
        <div class="rank-grid">${top10.map((c,i) => { const pct=calcPercent(c.aprovados,c.total); const barW=Math.round((c.total/maxTotal)*100); return `<div class="rank-card ${i<3?'rank-card-destaque':''}"><div class="rank-card-pos ${posClass(i)}">${medal(i)}</div><div class="rank-card-info"><div class="rank-card-name">${c.nome}</div><div class="rank-card-sector ${getSectorClass(c.setor)}">${c.setor||'—'}</div></div><div class="rank-card-metrics"><div class="rank-card-total">${c.total} <span>vendas</span></div><div class="rank-card-pct ${getPercentClass(pct)}">${pct}%</div></div><div class="rank-bar-wrap" style="width:120px"><div class="rank-bar-bg"><div class="rank-bar-fill" style="width:${barW}%"></div></div></div></div>`; }).join('')}</div>
        <div class="ranking-setores">${RANKING_SETORES.map(setor => { const lista=porSetor[setor]; if(!lista||!lista.length)return ''; return `<div class="ranking-setor-col"><div class="ranking-setor-header ${getSectorClass(setor)}"><i class="${getSectorIcon(setor)}"></i> ${setor}</div>${lista.slice(0,5).map((c,i) => { const pct=calcPercent(c.aprovados,c.total); return `<div class="rank-setor-row"><div class="rank-pos ${posClass(i)}" style="width:24px;height:24px;font-size:0.7rem">${medal(i)}</div><div class="rank-name" style="font-size:0.83rem">${c.nome}</div><div class="rank-pct ${getPercentClass(pct)}" style="font-size:0.78rem">${pct}%</div><div class="rank-total" style="font-size:0.75rem">${c.total}vd</div></div>`; }).join('')}</div>`; }).join('')}</div>
    </div>`;
}

export function renderDocumentacaoDashboard(d) {
    const { geral, vendasLoja, vendasWeb, consultores, mes, ano, temColunasPromo, recepcao, temRecepcao } = d;

    if (!consultores || consultores.length === 0) {
        contentEl().innerHTML = `
            <h2 class="dash-title"><i class="fas fa-folder" style="color:var(--primary)"></i> Dashboard de Vendas — ${mes} ${ano}</h2>
            <div class="error-message">
                <i class="fas fa-info-circle" style="font-size:2rem;color:var(--warning)"></i>
                <h3>Nenhum dado encontrado</h3>
                <p>Não há registros de vendas para o período selecionado.</p>
            </div>
        `;
        return;
    }

    const pG = calcPercentDoc(geral.aprovados, geral.total, geral.cancelados||0);
    const pL = calcPercentDoc(vendasLoja.aprovados, vendasLoja.total, vendasLoja.cancelados||0);
    const pW = calcPercentDoc(vendasWeb.aprovados, vendasWeb.total, vendasWeb.cancelados||0);

    const bySector = groupBySector(consultores);
    if (!bySector['OUTROS']) bySector['OUTROS'] = [];
    const sectors = sortSectors(Object.keys(bySector), ['VENDAS','RECEPCAO','REFILIACAO','WEB SITE','TELEVENDAS','OUTROS']);

    contentEl().innerHTML = `
    <h2 class="dash-title"><i class="fas fa-folder" style="color:var(--primary)"></i> Dashboard de Vendas — ${mes} ${ano}</h2>

    <div class="main-cards">
        ${cardDoc('Total de Vendas', 'fas fa-chart-bar', geral, pG)}
        ${cardDoc('Vendas Loja',     'fas fa-store',     vendasLoja, pL)}
        ${cardDoc('Vendas Web/Tele', 'fas fa-globe',     vendasWeb,  pW)}
    </div>

    ${temRecepcao ? renderRecepcaoBlocoHTML(recepcao) : ''}

    ${promoSection(d, temColunasPromo)}

    <h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--primary)"></i> Desempenho por Setor</h3>

    ${sectors.map(sector => {
        const list = bySector[sector] || [];

        if (sector === 'TELEVENDAS' && list.length > 0) {
            const cons = {
                total: 0, cancelados: 0, aprovados: 0, pendencias: 0,
                naoEnviado: 0, expirado: 0, pendente: 0, reprovados: 0,
                promo:  { total:0, cancelados:0, aprovados:0 },
                normal: { total:0, cancelados:0, aprovados:0 }
            };
            list.forEach(c => {
                cons.total        += c.total        || 0;
                cons.cancelados   += c.cancelados   || 0;
                cons.aprovados    += c.aprovados    || 0;
                cons.pendencias   += c.pendencias   || 0;
                cons.naoEnviado   += c.naoEnviado   || 0;
                cons.expirado     += c.expirado     || 0;
                cons.pendente     += c.pendente     || 0;
                cons.reprovados   += c.reprovados   || 0;
                if (c.promo) {
                    cons.promo.total      += c.promo.total      || 0;
                    cons.promo.cancelados += c.promo.cancelados || 0;
                    cons.promo.aprovados  += c.promo.aprovados  || 0;
                }
                if (c.normal) {
                    cons.normal.total      += c.normal.total      || 0;
                    cons.normal.cancelados += c.normal.cancelados || 0;
                    cons.normal.aprovados  += c.normal.aprovados  || 0;
                }
            });

            const pct = calcPercentDoc(cons.aprovados, cons.total, cons.cancelados);

            return `
            <div class="sector-card">
                <div class="sector-header">
                    <div class="sector-title">
                        <i class="${getSectorIcon(sector)}"></i> ${sector}
                        <span class="sector-count">${list.length} origem${list.length!==1?'s':''} consolidada${list.length!==1?'s':''}</span>
                    </div>
                    <div class="sector-header-right"><div class="metric-percent ${getPercentClass(pct)}">${pct}% aprovados</div>${snapBtn()}</div>
                </div>
                <div class="consultant-grid">
                    ${consultantCardPro({
                        nome: '📺 TELEVENDAS — Consolidado',
                        badge: `<div class="consultant-sector ${getSectorClass(sector)}">${sector}</div>`,
                        hero: [
                            { label: 'Total',     value: cons.total || 0 },
                            { label: 'Aprovados', value: cons.aprovados || 0, color: 'var(--success)' }
                        ],
                        pct,
                        pctLabel: '% Aprovados',
                        details: docDetails(cons),
                        extra: temColunasPromo ? promoMiniCards(cons) : ''
                    })}
                </div>
            </div>`;
        }

        if (list.length === 0) {
            return `
            <div class="sector-card" style="border-top-color:var(--gray)">
                <div class="sector-header">
                    <div class="sector-title">
                        <i class="${getSectorIcon(sector)}"></i> ${sector}
                        <span class="sector-count">0 consultores</span>
                    </div>
                    <div class="sector-header-right"><div class="metric-percent percent-low">0% aprovados</div>${snapBtn()}</div>
                </div>
                <div style="text-align:center;padding:20px;color:var(--text-muted);font-size:0.85rem">
                    <i class="fas fa-info-circle"></i> Nenhum dado neste setor
                </div>
            </div>`;
        }

        const tot   = list.reduce((s,c) => s + (c.total || 0), 0);
        const canc  = list.reduce((s,c) => s + (c.cancelados || 0), 0);
        const aprov = list.reduce((s,c) => s + (c.aprovados || 0), 0);
        const pct   = calcPercentDoc(aprov, tot, canc);

        return `
        <div class="sector-card">
            <div class="sector-header">
                <div class="sector-title">
                    <i class="${getSectorIcon(sector)}"></i> ${sector}
                    <span class="sector-count">${list.length} consultor${list.length!==1?'es':''}</span>
                </div>
                <div class="sector-header-right"><div class="metric-percent ${getPercentClass(pct)}">${pct}% aprovados</div>${snapBtn()}</div>
            </div>
            <div class="consultant-grid">
                ${list.sort((a,b) => (b.total || 0) - (a.total || 0)).map(c => {
                    const p = calcPercentDoc(c.aprovados, c.total, c.cancelados||0);
                    const isRecepcao = (c.setor||'').toUpperCase() === 'RECEPCAO';

                    let badgeRecepcao = '';
                    if (isRecepcao && c.recepcao) {
                        const indT = c.recepcao.indicacao?.total || 0;
                        const dirT = c.recepcao.direta?.total || 0;
                        badgeRecepcao = `
                            <div class="recepcao-badge-mini">
                                <span class="rec-badge rec-badge-ind"><i class="fas fa-crown"></i> ${indT}</span>
                                <span class="rec-badge rec-badge-dir"><i class="fas fa-bullseye"></i> ${dirT}</span>
                            </div>`;
                    }

                    return consultantCardPro({
                        nome: c.nome,
                        cardClass: isRecepcao ? 'consultant-card-recepcao' : '',
                        badge: isRecepcao ? badgeRecepcao : `<div class="consultant-sector ${getSectorClass(sector)}">${sector}</div>`,
                        hero: [
                            { label: 'Total',     value: c.total || 0 },
                            { label: 'Aprovados', value: c.aprovados || 0, color: 'var(--success)' }
                        ],
                        pct: p,
                        pctLabel: '% Aprovados',
                        details: docDetails(c),
                        extra: temColunasPromo ? promoMiniCards(c) : ''
                    });
                }).join('')}
            </div>
        </div>`;
    }).join('')}`;

    // Doughnut da Recepção: o canvas acabou de ser injetado no DOM acima
    if (temRecepcao) renderRecepcaoDoughnut(recepcao);
}

function renderRecepcaoBlocoHTML(rec) {
    if (!rec) return '';
    const ind = rec.indicacao || {};
    const dir = rec.direta    || {};
    const tot = rec.total     || {};

    const totalInd = tot.total || 0;
    const totalDir = dir.total || 0;
    const pInd = calcPercentDoc(ind.aprovados || 0, ind.total || 0, ind.cancelados || 0);
    const pDir = calcPercentDoc(dir.aprovados || 0, dir.total || 0, dir.cancelados || 0);
    const pTot = calcPercentDoc(tot.aprovados || 0, tot.total || 0, tot.cancelados || 0);

    const pctInd = totalInd > 0 ? Math.round((ind.total / totalInd) * 100) : 0;
    const pctDir = totalInd > 0 ? 100 - pctInd : 0;

    return `
    <h3 class="section-title" style="border-bottom-color:#9333ea">
        <i class="fas fa-users" style="color:#9333ea"></i> Recepção — Indicação vs Venda Direta
    </h3>

    <div class="recepcao-grid">
        <div class="recepcao-card recepcao-card-ind">
            <div class="recepcao-card-header">
                <i class="fas fa-crown"></i>
                <span>INDICAÇÃO</span>
            </div>
            <div class="recepcao-card-body">
                <div class="recepcao-numero">${ind.total || 0}</div>
                <div class="recepcao-numero-label">vendas</div>
                <div class="recepcao-kpis">
                    <div class="recepcao-kpi">
                        <span class="recepcao-kpi-icon"><i class="fas fa-check-circle"></i></span>
                        <span class="recepcao-kpi-label">Aprovados</span>
                        <span class="recepcao-kpi-value">${ind.aprovados || 0}</span>
                    </div>
                    <div class="recepcao-kpi">
                        <span class="recepcao-kpi-icon"><i class="fas fa-times-circle"></i></span>
                        <span class="recepcao-kpi-label">Cancelados</span>
                        <span class="recepcao-kpi-value">${ind.cancelados || 0}</span>
                    </div>
                    <div class="recepcao-kpi">
                        <span class="recepcao-kpi-icon"><i class="fas fa-exclamation-triangle"></i></span>
                        <span class="recepcao-kpi-label">Pendências</span>
                        <span class="recepcao-kpi-value">${ind.pendencias || 0}</span>
                    </div>
                </div>
                <div class="recepcao-barra-wrap">
                    <div class="recepcao-barra-label">
                        <span>Aprovação</span>
                        <span class="recepcao-barra-pct">${pInd}%</span>
                    </div>
                    <div class="recepcao-barra-bg">
                        <div class="recepcao-barra-fill recepcao-barra-ind" style="width:${pInd}%"></div>
                    </div>
                </div>
            </div>
        </div>

        <div class="recepcao-card recepcao-card-dir">
            <div class="recepcao-card-header">
                <i class="fas fa-bullseye"></i>
                <span>VENDA DIRETA</span>
            </div>
            <div class="recepcao-card-body">
                <div class="recepcao-numero">${dir.total || 0}</div>
                <div class="recepcao-numero-label">vendas</div>
                <div class="recepcao-kpis">
                    <div class="recepcao-kpi">
                        <span class="recepcao-kpi-icon"><i class="fas fa-check-circle"></i></span>
                        <span class="recepcao-kpi-label">Aprovados</span>
                        <span class="recepcao-kpi-value">${dir.aprovados || 0}</span>
                    </div>
                    <div class="recepcao-kpi">
                        <span class="recepcao-kpi-icon"><i class="fas fa-times-circle"></i></span>
                        <span class="recepcao-kpi-label">Cancelados</span>
                        <span class="recepcao-kpi-value">${dir.cancelados || 0}</span>
                    </div>
                    <div class="recepcao-kpi">
                        <span class="recepcao-kpi-icon"><i class="fas fa-exclamation-triangle"></i></span>
                        <span class="recepcao-kpi-label">Pendências</span>
                        <span class="recepcao-kpi-value">${dir.pendencias || 0}</span>
                    </div>
                </div>
                <div class="recepcao-barra-wrap">
                    <div class="recepcao-barra-label">
                        <span>Aprovação</span>
                        <span class="recepcao-barra-pct">${pDir}%</span>
                    </div>
                    <div class="recepcao-barra-bg">
                        <div class="recepcao-barra-fill recepcao-barra-dir" style="width:${pDir}%"></div>
                    </div>
                </div>
            </div>
        </div>

        <div class="recepcao-card recepcao-card-tot">
            <div class="recepcao-card-header">
                <i class="fas fa-chart-pie"></i>
                <span>TOTAL RECEPÇÃO</span>
            </div>
            <div class="recepcao-card-body">
                <div class="recepcao-numero">${tot.total || 0}</div>
                <div class="recepcao-numero-label">vendas</div>
                <div class="recepcao-doughnut-wrap">
                    <canvas id="recDoughnut"></canvas>
                </div>
                <div class="recepcao-doughnut-legend">
                    <div class="rec-leg-item">
                        <span class="rec-leg-dot" style="background:#9333ea"></span>
                        <span class="rec-leg-label">Indicação</span>
                        <span class="rec-leg-val">${ind.total || 0} (${pctInd}%)</span>
                    </div>
                    <div class="rec-leg-item">
                        <span class="rec-leg-dot" style="background:#1e40af"></span>
                        <span class="rec-leg-label">Venda Direta</span>
                        <span class="rec-leg-val">${dir.total || 0} (${pctDir}%)</span>
                    </div>
                </div>
                <div class="recepcao-barra-wrap">
                    <div class="recepcao-barra-label">
                        <span>Aprovação total</span>
                        <span class="recepcao-barra-pct">${pTot}%</span>
                    </div>
                    <div class="recepcao-barra-bg">
                        <div class="recepcao-barra-fill recepcao-barra-tot" style="width:${pTot}%"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>`;
}

function renderRecepcaoDoughnut(rec) {
    if (!rec || !document.getElementById('recDoughnut')) return;
    const ind = rec.indicacao?.total || 0;
    const dir = rec.direta?.total || 0;
    if (ind + dir === 0) return;

    createChart('recDoughnut', {
        type: 'doughnut',
        data: {
            labels: ['Indicação', 'Venda Direta'],
            datasets: [{
                data: [ind, dir],
                backgroundColor: ['#9333ea', '#1e40af'],
                borderWidth: 0,
                hoverOffset: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '70%',
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: ctx => `${ctx.label}: ${ctx.raw}` } }
            }
        }
    });
}

export function calcPercentDoc(aprovados, total, cancelados) {
    const base = (total||0) - (cancelados||0);
    return base > 0 ? Math.round((aprovados / base) * 100) : 0;
}

function promoMiniCards(c) {
    if (!c.promo && !c.normal) return '';
    const p  = c.promo  || {};
    const n  = c.normal || {};
    const pP = calcPercentDoc(p.aprovados||0, p.total||0, p.cancelados||0);
    const pN = calcPercentDoc(n.aprovados||0, n.total||0, n.cancelados||0);
    if (!p.total && !n.total) return '';
    return `
    <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
        <div style="background:var(--primary-light);border-radius:8px;padding:8px;border-left:3px solid var(--primary)">
            <div style="font-size:10px;font-weight:700;color:var(--primary-dark);margin-bottom:4px">▶ PROMOÇÃO</div>
            <div style="font-size:11px;color:var(--text)">Total: <b>${p.total||0}</b></div>
            <div style="font-size:11px;color:var(--text)">Aprovados: <b style="color:var(--success)">${p.aprovados||0}</b></div>
            <div style="margin-top:4px"><span class="metric-percent ${getPercentClass(pP)}">${pP}%</span></div>
        </div>
        <div style="background:#f0f9e8;border-radius:8px;padding:8px;border-left:3px solid var(--accent)">
            <div style="font-size:10px;font-weight:700;color:var(--accent-dark);margin-bottom:4px">▶ NORMAL</div>
            <div style="font-size:11px;color:var(--text)">Total: <b>${n.total||0}</b></div>
            <div style="font-size:11px;color:var(--text)">Aprovados: <b style="color:var(--success)">${n.aprovados||0}</b></div>
            <div style="margin-top:4px"><span class="metric-percent ${getPercentClass(pN)}">${pN}%</span></div>
        </div>
    </div>`;
}

function promoSection(d, temPromo) {
    if (!temPromo) return '';
    const p  = d.geral.promo  || {};
    const n  = d.geral.normal || {};
    const lp = d.vendasLoja.promo  || {};
    const ln = d.vendasLoja.normal || {};
    const pP  = calcPercentDoc(p.aprovados||0,  p.total||0,  p.cancelados||0);
    const pN  = calcPercentDoc(n.aprovados||0,  n.total||0,  n.cancelados||0);
    const pLP = calcPercentDoc(lp.aprovados||0, lp.total||0, lp.cancelados||0);
    const pLN = calcPercentDoc(ln.aprovados||0, ln.total||0, ln.cancelados||0);

    function miniCard(label, obj, pct, cor) {
        return `
        <div class="card" style="border-left:4px solid ${cor}">
            <div class="card-header" style="padding-bottom:8px">
                <div class="card-title" style="font-size:13px;color:${cor}">${label}</div>
            </div>
            <div class="metric-grid">
                ${metricItem('Total', obj.total||0)}
                ${metricItem('Cancelados', obj.cancelados||0, '#991b1b')}
                ${metricItem('Base líquida', (obj.total||0)-(obj.cancelados||0))}
                ${metricItem('Aprovados', obj.aprovados||0, 'var(--success)')}
                ${metricItem('Pendências', (obj.reprovados||0)+(obj.expirado||0)+(obj.pendente||0)+(obj.naoEnviado||0), 'var(--danger)')}
                ${metricPercent('% Aprovados', pct)}
            </div>
        </div>`;
    }

    return `
    <h3 class="section-title" style="color:var(--primary-dark)">
        <i class="fas fa-tags" style="color:var(--primary-dark)"></i> Comparativo: Promoção vs Normal
    </h3>
    <div class="main-cards" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr))">
        ${miniCard('Promoção — Geral',  p,  pP,  'var(--primary)')}
        ${miniCard('Normal — Geral',    n,  pN,  'var(--accent-dark)')}
        ${miniCard('Promoção — Loja',   lp, pLP, 'var(--primary-dark)')}
        ${miniCard('Normal — Loja',     ln, pLN, 'var(--accent-dark)')}
    </div>`;
}

function cardDoc(t, icon, d, pct) {
    const canc = d.cancelados || 0;
    const base = d.total - canc;
    return `
    <div class="card card-doc">
        <div class="card-header">
            <div class="card-title">${t}</div>
            <div class="card-icon"><i class="${icon}"></i></div>
        </div>
        <div class="metric-grid">
            ${metricItem('Total Vendas',  d.total)}
            ${metricItem('Cancelados',    canc,          '#991b1b')}
            ${metricItem('Base Líquida',  base)}
            ${metricItem('Aprovados',     d.aprovados,   'var(--success)')}
            ${metricItem('Pendências',    d.pendencias,  'var(--danger)')}
            ${metricItem('Não Enviado',   d.naoEnviado,  'var(--warning)')}
            ${metricItem('Expirado',      d.expirado,    'var(--gray)')}
            ${metricPercent('% Aprovados', pct)}
        </div>
    </div>`;
}

export function renderAppDashboard(d){
    const {geral, appLoja, appWeb, consultores, consultorasRetencao, mes, ano, equipeVendas, equipeConciliacao} = d;

    const pG = calcPercent(geral.sim, geral.total);
    const pL = calcPercent(appLoja.sim, appLoja.total);
    const pW = calcPercent(appWeb.sim, appWeb.total);

    const regular = consultores.filter(c => !c.origem || c.origem !== 'retencao');
    const bySector = groupBySector(regular);
    const sectors = sortSectors(Object.keys(bySector), ['VENDAS','RECEPCAO','REFILIACAO','WEB SITE','TELEVENDAS','OUTROS']);

    const appCard = (c, badge, borderColor) => consultantCardPro({
        nome: c.nome,
        badge,
        borderColor,
        hero: [
            { label: 'Total',   value: c.total || 0 },
            { label: 'Com App', value: c.sim || 0, color: 'var(--success)' }
        ],
        pct: calcPercent(c.sim || 0, c.total),
        pctLabel: '% Com App',
        details: [
            { label: 'Sem app',    value: c.nao || 0, color: 'var(--danger)' },
            { label: 'Cancelados', value: c.cancelado || 0 }
        ]
    });

    contentEl().innerHTML = `
    <h2 class="dash-title"><i class="fas fa-mobile-alt" style="color:var(--teal)"></i> Dashboard App — ${mes} ${ano}</h2>

    <div class="main-cards">
        ${cardApp('App — Total Geral', 'fas fa-chart-pie', geral, pG)}
        ${cardApp('App — Loja',        'fas fa-store',     appLoja, pL)}
        ${cardApp('App — Web/Tele',    'fas fa-globe',     appWeb,  pW)}
    </div>

    ${renderEquipesHTML(equipeVendas, equipeConciliacao)}

    ${consultorasRetencao && consultorasRetencao.length > 0 ? `
        <div class="retention-section">
            <div class="retention-header">
                <i class="fas fa-crown"></i>
                <h3>Consultoras de Retenção</h3>
            </div>
            <div class="consultant-grid">
                ${consultorasRetencao.map(c => appCard(
                    { ...c, nome: `${c.nome} (RETENÇÃO)` },
                    '<div class="consultant-sector sector-retencao">RETENÇÃO</div>',
                    '#f59e0b'
                )).join('')}
            </div>
        </div>` : ''}

    <h3 class="section-title"><i class="fas fa-layer-group" style="color:var(--teal)"></i> Desempenho por Setor</h3>

    ${sectors.map(sector => {
        const list = bySector[sector];
        const tot = list.reduce((s,c) => s + c.total, 0);
        const sim = list.reduce((s,c) => s + (c.sim||0), 0);
        const pct = calcPercent(sim, tot);
        return `
        <div class="sector-card">
            <div class="sector-header">
                <div class="sector-title">
                    <i class="${getSectorIcon(sector)}"></i> ${sector}
                    <span class="sector-count">${list.length} consultor${list.length!==1?'es':''}</span>
                </div>
                <div class="sector-header-right"><div class="metric-percent ${getPercentClass(pct)}">${pct}% com app</div>${snapBtn()}</div>
            </div>
            <div class="consultant-grid">
                ${list.sort((a,b) => b.total - a.total).map(c =>
                    appCard(c, `<div class="consultant-sector ${getSectorClass(sector)}">${sector}</div>`)
                ).join('')}
            </div>
        </div>`;
    }).join('')}`;
}

function renderEquipesHTML(eqVendas, eqConc) {
    if (!eqVendas && !eqConc) return '';
    const v = eqVendas || {total:0,sim:0,nao:0,cancelado:0,outros:0};
    const c = eqConc   || {total:0,sim:0,nao:0,cancelado:0,outros:0};
    const pV = calcPercent(v.sim, v.total);
    const pC = calcPercent(c.sim, c.total);

    return `
    <h3 class="section-title" style="border-bottom-color:#1e40af">
        <i class="fas fa-users" style="color:#1e40af"></i> Comparativo de Equipes
    </h3>

    <div class="equipe-grid">
        <div class="equipe-card equipe-card-vendas">
            <div class="equipe-card-header">
                <i class="fas fa-trophy"></i>
                <span>EQUIPE VENDAS</span>
            </div>
            <div class="equipe-card-body">
                <div class="equipe-numero">${v.total || 0}</div>
                <div class="equipe-numero-label">clientes</div>
                <div class="equipe-kpis">
                    <div class="equipe-kpi">
                        <span class="equipe-kpi-icon"><i class="fas fa-check-circle"></i></span>
                        <span class="equipe-kpi-label">Com App</span>
                        <span class="equipe-kpi-value">${v.sim || 0}</span>
                    </div>
                    <div class="equipe-kpi">
                        <span class="equipe-kpi-icon"><i class="fas fa-times-circle"></i></span>
                        <span class="equipe-kpi-label">Sem App</span>
                        <span class="equipe-kpi-value">${v.nao || 0}</span>
                    </div>
                    <div class="equipe-kpi">
                        <span class="equipe-kpi-icon"><i class="fas fa-ban"></i></span>
                        <span class="equipe-kpi-label">Cancelados</span>
                        <span class="equipe-kpi-value">${v.cancelado || 0}</span>
                    </div>
                </div>
                <div class="recepcao-barra-wrap">
                    <div class="recepcao-barra-label">
                        <span>Adesão ao App</span>
                        <span class="recepcao-barra-pct">${pV}%</span>
                    </div>
                    <div class="recepcao-barra-bg">
                        <div class="recepcao-barra-fill recepcao-barra-ind" style="width:${pV}%"></div>
                    </div>
                </div>
            </div>
        </div>

        <div class="equipe-card equipe-card-conc">
            <div class="equipe-card-header">
                <i class="fas fa-handshake"></i>
                <span>EQUIPE CONCILIAÇÃO</span>
            </div>
            <div class="equipe-card-body">
                <div class="equipe-numero">${c.total || 0}</div>
                <div class="equipe-numero-label">clientes</div>
                <div class="equipe-kpis">
                    <div class="equipe-kpi">
                        <span class="equipe-kpi-icon"><i class="fas fa-check-circle"></i></span>
                        <span class="equipe-kpi-label">Com App</span>
                        <span class="equipe-kpi-value">${c.sim || 0}</span>
                    </div>
                    <div class="equipe-kpi">
                        <span class="equipe-kpi-icon"><i class="fas fa-times-circle"></i></span>
                        <span class="equipe-kpi-label">Sem App</span>
                        <span class="equipe-kpi-value">${c.nao || 0}</span>
                    </div>
                    <div class="equipe-kpi">
                        <span class="equipe-kpi-icon"><i class="fas fa-ban"></i></span>
                        <span class="equipe-kpi-label">Cancelados</span>
                        <span class="equipe-kpi-value">${c.cancelado || 0}</span>
                    </div>
                </div>
                <div class="recepcao-barra-wrap">
                    <div class="recepcao-barra-label">
                        <span>Adesão ao App</span>
                        <span class="recepcao-barra-pct">${pC}%</span>
                    </div>
                    <div class="recepcao-barra-bg">
                        <div class="recepcao-barra-fill recepcao-barra-dir" style="width:${pC}%"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>`;
}

function cardApp(t,icon,d,pct){ return `<div class="card card-app"><div class="card-header"><div class="card-title">${t}</div><div class="card-icon"><i class="${icon}"></i></div></div><div class="metric-grid">${metricItem('Total Clientes',d.total)}${metricItem('Com App (SIM)',d.sim,'var(--success)')}${metricItem('Sem App (NÃO)',d.nao,'var(--danger)')}${metricItem('Cancelados',d.cancelado||0,'var(--gray)')}${metricPercent('% Com App',pct)}</div></div>`; }
