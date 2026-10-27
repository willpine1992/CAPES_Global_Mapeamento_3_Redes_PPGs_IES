// Abas, tooltip, explorador de PPGs e exportação PNG. DADOS é injetado pelo build_dashboard.py.
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
let vistaAtual = 'geral';

function mostrar(v) {
  vistaAtual = v;
  $$('.vista').forEach(s => s.classList.toggle('on', s.dataset.vista === v));
  $$('.aba').forEach(b => b.classList.toggle('on', b.dataset.vista === v));
  if (location.hash.slice(1) !== v) history.replaceState(null, '', '#' + v);
  window.scrollTo({ top: 0 });
}
$$('.aba').forEach(b => b.addEventListener('click', () => mostrar(b.dataset.vista)));
$$('[data-ir]').forEach(b => b.addEventListener('click', () => mostrar(b.dataset.ir)));

// tooltip: data-tip="Título|linha|linha"
const tip = $('#tip');
document.addEventListener('mouseover', ev => {
  const el = ev.target.closest('[data-tip]');
  if (!el) { tip.classList.remove('on'); return; }
  const [t, ...ls] = el.dataset.tip.split('|');
  const esc = s => s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  tip.innerHTML = `<b>${esc(t)}</b>${ls.length ? '<ul>' + ls.map(l => `<li>${esc(l)}</li>`).join('') + '</ul>' : ''}`;
  tip.classList.add('on');
});
document.addEventListener('mousemove', ev => {
  if (!tip.classList.contains('on')) return;
  const r = tip.getBoundingClientRect();
  let x = ev.clientX + 14, y = ev.clientY + 14;
  if (x + r.width > innerWidth - 8) x = ev.clientX - r.width - 14;
  if (y + r.height > innerHeight - 8) y = Math.max(8, ev.clientY - r.height - 14);
  tip.style.left = x + 'px'; tip.style.top = y + 'px';
});

// explorador
const COLS = ['rede', 'area', 'ies', 'ppg', 'cod', 'nota', 'pag'];
const LINHAS = DADOS.v.map(a => Object.fromEntries(COLS.map((k, i) => [k, a[i]])));
const F = { rede: $('#f-rede'), area: $('#f-area'), ies: $('#f-ies'), nota: $('#f-nota'), q: $('#f-q') };
let ordem = { k: null, dir: 1 };
const opts = (sel, vals, rot = v => v) => {
  const atual = sel.value;
  sel.innerHTML = '<option value="">Todas</option>' + vals.map(v => `<option value="${v}">${rot(v)}</option>`).join('');
  sel.value = vals.includes(atual) ? atual : '';
};
function opcoes() {
  const r = F.rede.value;
  const redes = r ? [r] : Object.keys(DADOS.redes);
  opts(F.area, redes.flatMap(x => DADOS.areas[x]));
  opts(F.ies, [...new Set(redes.flatMap(x => DADOS.ies[x]))].sort((a, b) => a === 'UEA' ? -1 : b === 'UEA' ? 1 : a.localeCompare(b)));
}
opts(F.rede, Object.keys(DADOS.redes), v => DADOS.redes[v]);
opcoes();
const norm = s => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
function filtrar() {
  const q = norm(F.q.value.trim()), nm = +F.nota.value;
  let ls = LINHAS.filter(l => (!F.rede.value || l.rede === F.rede.value) && (!F.area.value || l.area === F.area.value)
    && (!F.ies.value || l.ies === F.ies.value) && l.nota >= nm && (!q || norm(l.ppg).includes(q) || l.cod.toLowerCase().includes(q)));
  if (ordem.k) ls = [...ls].sort((a, b) => (a[ordem.k] > b[ordem.k] ? 1 : a[ordem.k] < b[ordem.k] ? -1 : 0) * ordem.dir);
  return ls;
}
function render() {
  const ls = filtrar();
  const distintos = new Set(ls.map(l => l.rede + l.cod)).size;
  $('#f-resumo').textContent = `${ls.length} vinculações · ${distintos} PPGs distintos (por rede) · ${new Set(ls.map(l => l.ies)).size} universidades`;
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  $('#f-corpo').innerHTML = ls.map(l => `<tr class="${l.ies === 'UEA' ? 'uea' : ''}">
    <td><span class="rd"><i class="dot" style="background:${DADOS.cor[l.rede]}"></i>${esc(DADOS.redes[l.rede])}</span></td>
    <td>${esc(l.area)}</td><td class="ies-c"><i class="lg lg-${l.ies.replace(/[^A-Za-z0-9]/g, '_')} lg-x" role="img" aria-label="${esc(l.ies)}" title="${esc(l.ies)}"></i></td><td>${esc(l.ppg)}</td><td class="cod">${esc(l.cod)}</td>
    <td class="c"><s class="nota-s">${l.nota}</s></td><td class="c">${l.pag}</td></tr>`).join('');
  $$('.t-exp th').forEach(th => th.className = (th.dataset.k === ordem.k ? (ordem.dir > 0 ? 'sort-asc' : 'sort-desc') : '') + (['nota', 'pag'].includes(th.dataset.k) ? ' c' : ''));
}
F.rede.addEventListener('change', () => { opcoes(); render(); });
[F.area, F.ies, F.nota].forEach(s => s.addEventListener('change', render));
F.q.addEventListener('input', render);
$('#f-limpar').addEventListener('click', () => { Object.values(F).forEach(s => s.value = s === F.nota ? '0' : ''); opcoes(); ordem = { k: null, dir: 1 }; render(); });
$$('.t-exp th').forEach(th => th.addEventListener('click', () => {
  ordem = { k: th.dataset.k, dir: ordem.k === th.dataset.k ? -ordem.dir : 1 }; render();
}));
$('#f-csv').addEventListener('click', () => {
  const q = v => `"${String(v).replace(/"/g, '""')}"`;
  const csv = ['Rede;Área temática;IES;Programa;Código CAPES;Nota;Página',
    ...filtrar().map(l => [DADOS.redes[l.rede], l.area, l.ies, l.ppg, l.cod, l.nota, l.pag].map(q).join(';'))].join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }));
  a.download = 'ppgs_redes_capes_global.csv'; a.click();
});
// clique numa célula da matriz abre o explorador filtrado
$$('[data-filtro]').forEach(td => td.addEventListener('click', () => {
  const [r, a, u] = td.dataset.filtro.split('§');
  F.rede.value = r; opcoes(); F.area.value = a; F.ies.value = u; F.nota.value = '0'; F.q.value = '';
  tip.classList.remove('on'); render(); mostrar('explorar');
}));
render();

// PNG da vista atual
$('[data-acao="png"]').addEventListener('click', async ev => {
  const b = ev.currentTarget, txt = b.textContent;
  b.disabled = true; b.textContent = 'Gerando…';
  try {
    const el = $(`.vista[data-vista="${vistaAtual}"]`);
    const url = await htmlToImage.toPng(el, { pixelRatio: 2, backgroundColor: '#fbf9f4', cacheBust: true });
    const a = document.createElement('a'); a.href = url; a.download = `UEA_redes_${vistaAtual}.png`; a.click();
    b.textContent = txt;
  } catch (err) { console.error(err); b.textContent = 'Erro — ver console'; setTimeout(() => b.textContent = txt, 2500); }
  finally { b.disabled = false; }
});

const h = location.hash.slice(1);
if ($(`.vista[data-vista="${h}"]`)) mostrar(h);
