// ============================================================================
// equipe.js — aba "Equipe / Usuários" (só para papel 'admin' na unidade ativa)
// Toda a gestão passa pela Edge Function `gerenciar-equipe`, que revalida o
// papel no servidor (o botão escondido na sidebar não é uma barreira de segurança).
// ============================================================================
import { AUTH, state } from './config.js';
import { contentEl, showSkeleton, hideSkeleton, showError, showToast } from './ui.js';
import { escapeHtml } from './utils.js';

const PAPEL_ROTULO = { admin: 'Administrador', leitor: 'Leitor' };
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

// Chama a Edge Function. O supabase-js anexa o JWT do usuário logado automaticamente.
async function chamarEquipe(corpo){
    const { data, error } = await AUTH.client.functions.invoke('gerenciar-equipe', { body: corpo });
    if (error) {
        let msg = 'Não foi possível falar com o servidor. Tente novamente.';
        try {
            const j = await error.context.json();   // resposta de erro da função: { error: "..." }
            if (j && j.error) msg = j.error;
        } catch (_) {}
        throw new Error(msg);
    }
    if (data && data.error) throw new Error(data.error);
    return data || {};
}

function formatarAcesso(iso){
    if (!iso) return '<span class="team-muted">Nunca acessou</span>';
    const d = new Date(iso);
    return Number.isNaN(d.getTime()) ? '—' : d.toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
}

function linhaUsuario(u, meuId){
    const ehEu = u.user_id === meuId;
    const papel = u.papel === 'admin' ? 'admin' : 'leitor';
    return `
    <tr>
        <td>
            <strong>${escapeHtml(u.email || '(sem e-mail)')}</strong>
            ${ehEu ? '<span class="team-you">você</span>' : ''}
        </td>
        <td><span class="team-badge team-badge-${papel}">${PAPEL_ROTULO[papel]}</span></td>
        <td>${formatarAcesso(u.ultimo_acesso)}</td>
        <td class="team-actions">
            ${ehEu
                ? '<span class="team-muted" title="Você não pode remover o próprio acesso">—</span>'
                : `<button type="button" class="team-remove" data-user-id="${escapeHtml(u.user_id)}" data-email="${escapeHtml(u.email || '')}" title="Remover da unidade"><i class="fas fa-user-minus"></i> Remover</button>`}
        </td>
    </tr>`;
}

function desenhar(unidade, usuarios){
    const meuId = AUTH.session && AUTH.session.user ? AUTH.session.user.id : null;

    contentEl().innerHTML = `
    <h2 class="dash-title"><i class="fas fa-users-cog" style="color:var(--primary)"></i> Equipe — ${escapeHtml(unidade.nome_exibicao)}</h2>

    <div class="team-layout">
        <section class="card team-card">
            <div class="card-header">
                <div class="card-title">Usuários da unidade (${usuarios.length})</div>
                <div class="card-icon"><i class="fas fa-users"></i></div>
            </div>
            ${usuarios.length === 0
                ? '<p class="team-muted" style="padding:8px 0">Nenhum usuário vinculado.</p>'
                : `<div class="table-wrapper"><table class="data-table">
                    <thead><tr><th>E-mail</th><th>Papel</th><th>Último acesso</th><th></th></tr></thead>
                    <tbody>${usuarios.map(u => linhaUsuario(u, meuId)).join('')}</tbody>
                </table></div>`}
        </section>

        <section class="card team-card">
            <div class="card-header">
                <div class="card-title">Novo usuário</div>
                <div class="card-icon"><i class="fas fa-user-plus"></i></div>
            </div>
            <form id="teamForm" novalidate>
                <div class="team-field">
                    <label for="teamEmail">E-mail</label>
                    <input type="email" id="teamEmail" autocomplete="off" placeholder="lider@empresa.com" required>
                </div>
                <div class="team-field">
                    <label for="teamPassword">Senha provisória (mín. 8 caracteres)</label>
                    <input type="password" id="teamPassword" autocomplete="new-password" minlength="8" placeholder="••••••••" required>
                </div>
                <div class="team-field">
                    <label for="teamPapel">Papel</label>
                    <select id="teamPapel">
                        <option value="leitor">Leitor — só visualiza os dashboards</option>
                        <option value="admin">Administrador — também gerencia a equipe</option>
                    </select>
                </div>
                <div class="team-msg" id="teamMsg" role="alert"></div>
                <button type="submit" class="btn btn-topbar-export team-submit" id="teamSubmit">
                    <i class="fas fa-plus"></i><span>Criar usuário</span>
                </button>
            </form>
        </section>
    </div>`;

    ligarEventos(unidade);
}

function mostrarMsg(texto, tipo){
    const el = document.getElementById('teamMsg');
    if (!el) return;
    el.textContent = texto || '';
    el.className = 'team-msg' + (texto ? ` show ${tipo || 'error'}` : '');
}

function ligarEventos(unidade){
    const form = document.getElementById('teamForm');
    if (form) form.addEventListener('submit', async ev => {
        ev.preventDefault();
        const email = document.getElementById('teamEmail').value.trim().toLowerCase();
        const password = document.getElementById('teamPassword').value;
        const papel = document.getElementById('teamPapel').value;

        if (!EMAIL_RE.test(email)) { mostrarMsg('Informe um e-mail válido.'); return; }
        if (password.length < 8)   { mostrarMsg('A senha deve ter pelo menos 8 caracteres.'); return; }
        if (!['leitor', 'admin'].includes(papel)) { mostrarMsg('Escolha um papel válido.'); return; }

        const btn = document.getElementById('teamSubmit');
        btn.disabled = true;
        btn.querySelector('span').textContent = 'Criando…';
        mostrarMsg('');
        try {
            await chamarEquipe({ acao: 'criar', unidade_id: unidade.id, email, password, papel });
            showToast('Usuário criado ✓');
            await renderEquipeDashboard();          // recarrega a lista (e limpa o formulário)
        } catch (e) {
            mostrarMsg(e.message);
            btn.disabled = false;
            btn.querySelector('span').textContent = 'Criar usuário';
        }
    });

    document.querySelectorAll('.team-remove').forEach(btn => btn.addEventListener('click', async () => {
        const email = btn.dataset.email || 'este usuário';
        if (!window.confirm(`Remover ${email} da unidade ${unidade.nome_exibicao}?\n\nEle perderá o acesso a este dashboard.`)) return;
        btn.disabled = true;
        try {
            await chamarEquipe({ acao: 'remover', unidade_id: unidade.id, user_id: btn.dataset.userId });
            showToast('Usuário removido ✓');
            await renderEquipeDashboard();
        } catch (e) {
            btn.disabled = false;
            showToast('❌ ' + e.message);
        }
    }));
}

export async function renderEquipeDashboard(){
    const unidade = AUTH.unidade;
    if (!unidade || unidade.papel !== 'admin') {
        showError('Acesso restrito a administradores da unidade.', false);
        return;
    }

    const seq = state.loadSeq;       // se trocar de tela/unidade durante a busca, descarta o resultado
    showSkeleton();
    try {
        const resp = await chamarEquipe({ acao: 'listar', unidade_id: unidade.id });
        if (seq !== state.loadSeq) return;
        hideSkeleton();
        desenhar(unidade, Array.isArray(resp.usuarios) ? resp.usuarios : []);
    } catch (e) {
        if (seq !== state.loadSeq) return;
        showError(e.message, true);
    }
}
