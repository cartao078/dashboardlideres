// ============================================================================
// utils.js — helpers puros (sem DOM, sem estado): cálculos, agrupamentos, rótulos
// ============================================================================

export function calcPercent(n,d){ if(!d||d===0)return 0; return Math.round((n/d)*100); }

export function groupBySector(list) {
    const m = {};
    if (!list || list.length === 0) return m;
    list.forEach(c => {
        const s = (c.setor || 'OUTROS').toUpperCase();
        if (!m[s]) m[s] = [];
        m[s].push(c);
    });
    if (!m['OUTROS']) m['OUTROS'] = [];
    return m;
}

export function sortSectors(keys, order) {
    const uniqueKeys = [...new Set(keys)];
    const orderWithOutros = order.includes('OUTROS') ? order : [...order, 'OUTROS'];
    return uniqueKeys.sort((a, b) => {
        const ia = orderWithOutros.indexOf(a);
        const ib = orderWithOutros.indexOf(b);
        if (ia === -1 && ib === -1) return a.localeCompare(b);
        if (ia === -1) return 1;
        if (ib === -1) return -1;
        return ia - ib;
    });
}

export function getPercentClass(p){ p=typeof p==='number'?p:parseInt(p)||0; if(p>=90)return 'percent-high'; if(p>=80)return 'percent-medium'; return 'percent-low'; }
export function getSectorClass(s){ switch((s||'').toUpperCase()){case 'VENDAS':return 'sector-vendas';case 'RECEPCAO':return 'sector-recepcao';case 'REFILIACAO':return 'sector-refiliacao';case 'WEB SITE':case 'WEB':return 'sector-web';case 'TELEVENDAS':return 'sector-televendas';case 'RETENÇÃO':case 'RETENCAO':return 'sector-retencao';default:return 'sector-outros';} }
export function getSectorIcon(s){ switch((s||'').toUpperCase()){case 'VENDAS':return 'fas fa-shopping-cart';case 'RECEPCAO':return 'fas fa-headset';case 'REFILIACAO':return 'fas fa-user-plus';case 'WEB SITE':case 'WEB':return 'fas fa-globe';case 'TELEVENDAS':return 'fas fa-phone-alt';case 'RETENÇÃO':case 'RETENCAO':return 'fas fa-crown';default:return 'fas fa-users';} }

export function getMonthName(m){ return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][m-1]; }

// Escapa texto vindo de fora (e-mails, nomes) antes de entrar em innerHTML
export function escapeHtml(v){
    return String(v ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}
