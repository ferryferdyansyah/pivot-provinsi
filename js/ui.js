// ui.js - lapisan antarmuka: animasi loading, status, ringkasan, filter tab, navigasi.
// Dimuat SEBELUM main.js (main.js memakai Busy & UI), setelah render.js/sheets.js.
const $ = id => document.getElementById(id);
const tick = () => new Promise(r => { let d = false; const go = () => d || (d = true, r()); requestAnimationFrame(() => setTimeout(go, 30)); setTimeout(go, 120) });   // beri browser kesempatan menggambar ulang

// ===== Overlay loading =====
const Busy = {
    t0: 0, timer: 0,
    show(title, steps = []) {
        $('busy-t').textContent = title; $('busy-d').textContent = '';
        $('busy-steps').innerHTML = steps.map(s => `<li>${s}</li>`).join('');
        const tr = document.querySelector('.bar-track'); tr.classList.remove('det'); $('busy-bar').style.width = '';
        $('busy').hidden = false; $('busy').setAttribute('aria-busy', 'true'); document.body.style.overflow = 'hidden';
        this.t0 = Date.now(); clearInterval(this.timer);
        const tm = () => $('busy-time').textContent = `${Math.round((Date.now() - this.t0) / 1000)} dtk`; tm(); this.timer = setInterval(tm, 500);
        if (steps.length) this.stage(0);
    },
    detail(s) { $('busy-d').textContent = s },
    stage(i) { [...$('busy-steps').children].forEach((li, k) => { li.className = k < i ? 'done' : k === i ? 'on' : '' }) },
    progress(p) { document.querySelector('.bar-track').classList.add('det'); $('busy-bar').style.width = Math.max(0, Math.min(100, p)) + '%' },
    hide() { clearInterval(this.timer); $('busy').hidden = true; $('busy').removeAttribute('aria-busy'); document.body.style.overflow = '' }
};

// ===== Helper UI =====
const UI = {
    meta: {},                                   // mis. { excelRows }
    tick,
    drop(el, state, title, sub) {               // state: idle | done | error
        el.dataset.state = state;
        el.querySelector('.dz-t').textContent = title || el.querySelector('.dz-t').dataset.t;
        el.querySelector('.dz-s').textContent = sub || el.querySelector('.dz-s').dataset.s;
    }
};

// ===== Status bar: warna mengikuti isi teks (#status ditulis oleh banyak file) =====
(() => {
    const bar = $('statusbar'), st = $('status');
    const kind = t => /^(Gagal|Kolom HAT|File \.dbf)|harus lebih kecil/i.test(t) ? 'error'
        : /Peringatan|Unggah file (utama|PDF)/i.test(t) ? 'warn'
        : /^(Membaca|Menyiapkan|Memproses)/i.test(t) ? 'busy'
        : /baris|terisi|diubah|diisi|dikembalikan|^Tahun lama/i.test(t) ? 'ok' : 'info';
    new MutationObserver(() => { bar.dataset.kind = kind(st.textContent) }).observe(st, { childList: true, characterData: true, subtree: true });
})();

// ===== Ringkasan + filter tab (dipanggil setelah setiap render) =====
let _sumKey = '';
function animateNum(el, to, fmtFn) {
    const t0 = performance.now(), D = 700;
    const step = now => { const p = Math.min(1, (now - t0) / D), e = 1 - Math.pow(1 - p, 3); el.textContent = fmtFn(to * e); if (p < 1) requestAnimationFrame(step) };
    requestAnimationFrame(step);
}
function updateSummary() {
    const box = $('summary'), has = ROWS.length > 0;
    document.body.classList.toggle('has-data', has);
    $('drop2').dataset.locked = has ? '' : '1'; if (has) $('drop2').removeAttribute('data-locked');
    if (!has) { box.hidden = true; _sumKey = ''; return }
    const kab = COLS.includes('WADMKK') ? new Set(ROWS.map(r => r.WADMKK).filter(Boolean)).size : 0;
    const ha = ROWS.reduce((s, r) => s + (parseFloat(r.LUASHA) || 0), 0);
    const ok = document.querySelectorAll('.tab.ok').length, tot = document.querySelectorAll('.tab').length;
    const items = [[ROWS.length, 'Baris Raw Data', v => Math.round(v).toLocaleString('id')]];
    if (UI.meta.excelRows) items.unshift([UI.meta.excelRows, 'Baris Excel asli', v => Math.round(v).toLocaleString('id')]);
    if (kab) items.push([kab, 'Kabupaten/Kota', v => Math.round(v).toString()]);
    items.push([ha, 'Total luas (Ha)', v => Math.round(v).toLocaleString('id')]);
    items.push([ok, `Tabel terisi dari ${tot}`, v => Math.round(v).toString()]);
    const key = items.map(i => i[0].toFixed(2)).join('|');
    box.hidden = false; if (key === _sumKey) return; _sumKey = key;
    box.innerHTML = items.map(i => `<div class="stat"><b>0</b><span>${i[1]}</span></div>`).join('');
    items.forEach((it, k) => animateNum(box.children[k].firstChild, it[0], it[2]));
}
const grpOf = n => /Raw/i.test(n) ? 'raw' : /Tabel_III|Grafik_III/i.test(n) ? 'III' : 'IV';
let curGrp = 'all';
function buildTabs() {
    const tabs = [...document.querySelectorAll('.tab')];
    tabs.forEach(t => { t.dataset.g = grpOf(t.textContent); t.hidden = curGrp !== 'all' && t.dataset.g !== curGrp });
    const cnt = g => g === 'all' ? tabs.length : tabs.filter(t => t.dataset.g === g).length;
    $('tabfilter').innerHTML = [['all', 'Semua'], ['raw', 'Raw Data'], ['III', 'Bab III'], ['IV', 'Bab IV']].filter(([g]) => cnt(g))
        .map(([g, l]) => `<button class="chip ${g === curGrp ? 'on' : ''}" data-g="${g}">${l} · ${cnt(g)}</button>`).join('');
    paintTitle();
}
function paintTitle() { const on = document.querySelector('.tab.on'); $('paneltitle').textContent = on ? on.textContent.replace(/^\d+_/, '').replace(/_/g, ' ') : '' }
function stepTab(d) {
    const v = [...document.querySelectorAll('.tab')].filter(t => !t.hidden), i = v.findIndex(t => t.classList.contains('on'));
    const n = v[(i + d + v.length) % v.length]; if (n) { n.click(); n.scrollIntoView({ block: 'nearest', inline: 'center' }) }
}
document.addEventListener('click', e => {
    const c = e.target.closest('.chip'); if (c) { curGrp = c.dataset.g; buildTabs(); const f = [...document.querySelectorAll('.tab')].find(t => !t.hidden && t.classList.contains('on')) || [...document.querySelectorAll('.tab')].find(t => !t.hidden); if (f && !f.classList.contains('on')) f.click(); return }
    if (e.target.closest('.tab')) paintTitle();
    if (e.target.closest('#prevtab')) stepTab(-1);
    if (e.target.closest('#nexttab')) stepTab(1);
});
document.addEventListener('keydown', e => {
    const dz = e.target.closest?.('.dz'); if (dz && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); dz.click(); return }
    if (!document.body.classList.contains('has-data') || !$('busy').hidden || /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
    if (e.key === 'ArrowRight') stepTab(1); else if (e.key === 'ArrowLeft') stepTab(-1);
});

// bungkus render(): setelah tabel dibuat, perbarui ringkasan & tab
const _render = render;
render = function () { _render(); updateSummary(); buildTabs() };
