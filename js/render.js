// render.js
// Menampilkan tab, tabel, dan grafik di halaman web.

function tableHTML(d, def) {
    let h = `<table><tr><th rowspan=2>No</th><th rowspan=2>${def.h}</th><th colspan=${d.kecs.length}>Luas di Tiap Kabupaten/Kota (Ha)</th><th rowspan=2>Jumlah (Ha)</th><th rowspan=2>Persentase (%)</th></tr><tr>${d.kecs.map(k => `<th>${k}</th>`).join('')}</tr>`;
    d.body.forEach((b, i) => h += `<tr><td class=c>${i + 1}</td><td>${b.c}</td>${b.v.map(v => `<td class=n>${fmt(v)}</td>`).join('')}<td class=n><b>${fmt(b.t)}</b></td><td class=n><b>${fmt(b.p)}</b></td></tr>`);
    h += `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${d.kt.map(v => `<td class=n>${fmt(v)}</td>`).join('')}<td class=n>${fmt(d.T)}</td><td class=n>100.00</td></tr></table>`;
    return h;
}
function top5HTML(d, def) {
    return `<table><tr><th>No</th><th>${def.t5}</th><th>Jumlah (Ha)</th></tr>` + d.top.map((b, i) => `<tr><td class=c>${i + 1}</td><td>${b.c}</td><td class=n>${fmt(b.t)}</td></tr>`).join('') + '</table>';
}
let charts = [];
// angka notasi ilmiah (mis. 6.11922125787e+00) ditampilkan 3 angka di belakang koma
const SCI = /^[-+]?[0-9.]+e[-+]?[0-9]+$/i;
const fmtRaw = v => SCI.test(v) ? (+v).toFixed(2).replace('-0.00', '0.00') : v;
// ---- Raw Data: tampil per 100 baris ----
const PER = 100; let rawPage = 1;
function pagerHTML(pages, s, e) {
    const n = ROWS.length, b = (a, l, dis) => `<button data-pg="${a}" ${dis ? 'disabled' : ''}>${l}</button>`;
    return `<div class=pager>${b('first', '«', rawPage <= 1)}${b('prev', '‹', rawPage <= 1)}<span>Hal.</span><input type=number min=1 max=${pages} value=${rawPage} data-go aria-label="Nomor halaman"><span>/ ${pages.toLocaleString('id')}</span>${b('next', '›', rawPage >= pages)}${b('last', '»', rawPage >= pages)}<span class=sp>Baris ${(s + 1).toLocaleString('id')}–${e.toLocaleString('id')} dari ${n.toLocaleString('id')}</span></div>`;
}
function renderRaw() {
    const p0 = document.getElementById('p0'), n = ROWS.length, pages = Math.max(1, Math.ceil(n / PER));
    rawPage = Math.min(Math.max(rawPage, 1), pages);
    const s = (rawPage - 1) * PER, e = Math.min(s + PER, n), L = ROWS.slice(s, e), pg = pagerHTML(pages, s, e);
    p0.innerHTML = pg + `<div class=raw-scroll><table><tr><th>#</th>${COLS.map(c => `<th>${c}</th>`).join('')}</tr>${L.map((r, i) => `<tr><td class=idx>${s + i + 1}</td>${COLS.map(c => `<td>${fmtRaw(r[c])}</td>`).join('')}</tr>`).join('')}</table></div>` + pg + `<div class=note>Tampilan web dibatasi ${PER} baris per halaman. File Excel memuat semua baris.</div>`;
    p0.scrollTop = 0;
}
document.addEventListener('click', e => {
    const b = e.target.closest('#p0 [data-pg]'); if (!b) return;
    const pages = Math.ceil(ROWS.length / PER), a = b.dataset.pg;
    rawPage = a === 'first' ? 1 : a === 'last' ? pages : rawPage + (a === 'next' ? 1 : -1); renderRaw();
});
document.addEventListener('change', e => {
    if (!e.target.matches('#p0 [data-go]')) return;
    rawPage = parseInt(e.target.value) || 1; renderRaw();
});
function render() {
    charts.forEach(c => c.destroy()); charts = [];
    const tabs = document.getElementById('tabs'), pn = document.getElementById('panels'); tabs.innerHTML = pn.innerHTML = '';
    const names = SHEETS.map(sheetName);
    names.forEach((n, i) => {
        tabs.insertAdjacentHTML('beforeend', `<button class="tab ${i ? '' : 'on'}" data-i=${i}>${n}</button>`);
        pn.insertAdjacentHTML('beforeend', `<div class="panel ${i ? '' : 'on'}" id="p${i}"></div>`);
    });
    tabs.onclick = e => {
        const i = e.target.dataset.i; if (i === undefined) return;
        document.querySelectorAll('.tab').forEach(t => t.classList.toggle('on', t.dataset.i === i));
        document.querySelectorAll('.panel').forEach(p => p.classList.toggle('on', p.id === 'p' + i))
    };
    const p0 = document.getElementById('p0');
    if (!ROWS.length) p0.innerHTML = `<div class=nf>${NEW_TEXT}</div>`;
    else { rawPage = 1; renderRaw() }
    MODELS = DEFS.map((def, i) => {
        const d = agg(def.f), p = document.getElementById('p' + panelOf('def', i));
        if (!d) { p.innerHTML = `<div class=nf>${NEW_TEXT}</div>`; return null }
        p.innerHTML = `<div class=wrap><div>${tableHTML(d, def)}</div><div class=side>${def.type === 'bar' ? top5HTML(d, def) : ''}<div class=cv><canvas width=${def.type === 'pie' ? 640 : 520} height=300></canvas></div></div></div>`;
        charts.push(mkChart(p.querySelector('canvas'), def, d));
        return d;
    });
    // sheet baru (isi belum dibuat)
    SHEETS.forEach((s, i) => { if (s.kind === 'new') document.getElementById('p' + i).innerHTML = `<div class=nf>${NEW_TEXT}</div>` });
    // sheet tambahan (extra.js)
    EXTRA_MODELS = EXTRA.map((ex, j) => {
        const d = ex.build(), p = document.getElementById('p' + panelOf('extra', j));
        if (!d) { p.innerHTML = `<div class=nf>${NEW_TEXT}</div>`; return null }
        p.innerHTML = ex.html(d, ex);
        const cv = p.querySelector('canvas'); if (cv && ex.chart) charts.push(ex.chart(cv, d));
        return d;
    });
    // penanda: outline hijau = tabel terisi, merah = "Data Belum Ditemukan"
    SHEETS.forEach((_, i) => {
        const ok = !document.getElementById('p' + i).querySelector('.nf');
        [document.querySelector(`.tab[data-i="${i}"]`), document.getElementById('p' + i)].forEach(el => el.classList.add(ok ? 'ok' : 'miss'));
    });
}
