// extra.js
// Sheet tambahan (tabel & grafik yang bentuknya khusus, tidak bisa memakai agg() biasa):
//   7. Perubahan per Kabupaten/Kota      <- GQREKLAS (Berubah / Tidak Berubah) x WADMKC
//   8. Perbandingan Penggunaan Tanah <- GNAME25 (2014) vs QNAME25 (2026)
//   9. Perubahan Terluas             <- GQNAME ("A menjadi B")
//  10. Reklasifikasi                 <- GREKLAS (2014) vs QREKLAS (2026)
// Tiap sheet = satu objek di array EXTRA di bagian bawah file ini:
//   build()  -> menghitung data (null = '(Data tidak ditemukan)')
//   html(d)  -> HTML tabel + kanvas grafik untuk tampilan web
//   chart()  -> membuat grafik Chart.js (boleh kosong jika tidak ada grafik)
//   excel()  -> menulis sheet Excel

let EXTRA_MODELS = [];
const PAL2 = PAL.concat(['#ff9da7', '#9c755f', '#bab0ab', '#59a14f']);

// ---------- Helper kecil ----------
const fd = v => fmt(Math.abs(v) < 0.005 ? 0 : v);       // hindari tampilan "-0.00"
const pc = (x, t) => (t ? x / t * 100 : 0);
const th = (t, o = '') => `<th ${o}>${yr(t)}</th>`;
const td = (v, c = '') => `<td class="${c}">${v}</td>`;
const cvTag = ex => `<div class=cv><canvas width=${ex.w} height=${ex.h}></canvas></div>`;
const kecOf = r => (COLS.includes('WADMKC') && r.WADMKC) || '-';
const area = r => parseFloat(r.LUASHA) || 0;
// "Tidak Dapat Dianalisis" (dikenali dari kata "dianalis...") ditampilkan sebagai kategori tersendiri di tabel yang memakai
// kolom KSPOLA, VNAME, dan GQREKLAS; kolomnya hanya muncul bila nilai itu memang ada di data.
const isTDA = v => /analis/i.test(String(v || ''));   // "Tidak Dapat Dianalisis", "...Dianalisa", "di analisis", dst
const TDA_ALWAYS = true;   // true: kolom Tidak Dapat Dianalisis SELALU tampil di tabel KSPOLA & VNAME (walau luasnya 0); false: hanya bila ada di data
const TDA_LABEL = 'Tidak Dapat Dianalisis', TDA_COLOR = '#7f7f7f';
// kabupaten/kota sebuah baris (kunci tahun lama per kab/kota); jika kolom WADMKK tidak ada, pakai nama wilayah tabel
const kabOfRow = r => ((r.WADMKK || '').trim()) || kecOf(r);
// teks kolom "Tahun Lama - Tahun Baru", mis. "2015 - 2026"
const thRangeTxt = kab => `${thLamaOf(kab)} - ${TH_BARU}`;

// jumlah luas per kategori untuk 1 kolom -> {kategori: luas}, atau null jika kolom tidak ada/kosong
function sumBy(f) {
  if (!ROWS.length || !COLS.includes(f)) return null;
  const m = {}; let any = false;
  for (const r of ROWS) { if (!r[f]) continue; any = true; m[r[f]] = (m[r[f]] || 0) + area(r); }
  return any ? m : null;
}

// ---------- Excel helpers ----------
// tulis sel header (bisa digabung) dengan gaya biru tua
function H(ws, r1, c1, r2, c2, v) {
  if (r1 !== r2 || c1 !== c2) ws.mergeCells(r1, c1, r2, c2);
  cell(ws, r1, c1, yr(v));
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) hdr(ws.getCell(r, c));
}
// sisipkan grafik sebagai gambar PNG
function addChart(wb, ws, ex, d, col, row) {
  const cv = document.createElement('canvas'); cv.width = ex.w; cv.height = ex.h;
  const ch = ex.chart(cv, d, false), url = cv.toDataURL('image/png'); ch.destroy();
  ws.addImage(wb.addImage({ base64: url, extension: 'png' }), { tl: { col, row }, ext: { width: ex.w, height: ex.h } });
}

// ---------- Plugin label nilai untuk bar chart multi-seri ----------
const multiVals = {
  id: 'mv', afterDatasetsDraw(ch) {
    const c = ch.ctx; c.save(); c.font = '11px Arial'; c.textAlign = 'center'; c.fillStyle = '#222';
    ch.data.datasets.forEach((ds, i) => ch.getDatasetMeta(i).data.forEach((el, j) => {
      const v = ds.data[j]; c.fillText(fmt(v), el.x, v < 0 ? el.y + 13 : el.y - 5);
    })); c.restore();
  }
};
const baseOpt = (title, a, legend) => ({
  responsive: false, animation: a ? undefined : false, layout: { padding: { top: 18 } },
  plugins: { legend: { display: !!legend, position: legend, labels: { boxWidth: 10, font: { size: 10 } } }, title: { display: true, text: Array.isArray(title) ? title.map(yr) : yr(title), font: { size: 12 } } }
});

// =====================================================================
// SHEET 7 - Perubahan per Kabupaten/Kota
// =====================================================================
function build7() {
  if (!ROWS.length || !COLS.includes('GQREKLAS')) return null;
  const m = {}, kb = {}; let any = false, tda = false;
  for (const r of ROWS) {
    const g = (r.GQREKLAS || '').trim().toLowerCase(); if (!g) continue; any = true;
    const k = kecOf(r); m[k] ??= { b: 0, tb: 0, u: 0 }; kb[k] ??= kabOfRow(r);
    if (isTDA(g)) { m[k].u += area(r); tda = true; }                       // "Tidak Dapat Dianalisis"
    else if (g.startsWith('tidak')) m[k].tb += area(r); else m[k].b += area(r);   // "Tidak Berubah" vs "Berubah"
  }
  if (!any) return null;
  // thn = "Tahun Lama - Tahun Baru" (tahun lama diambil per kabupaten/kota, lihat thLamaOf di config.js)
  const rows = Object.keys(m).sort().map(k => ({ k, thn: thRangeTxt(kb[k]), b: m[k].b, tb: m[k].tb, u: m[k].u, t: m[k].b + m[k].tb + m[k].u }));
  const tot = rows.reduce((a, r) => ({ b: a.b + r.b, tb: a.tb + r.tb, u: a.u + r.u, t: a.t + r.t }), { b: 0, tb: 0, u: 0, t: 0 });
  return { rows, tot, tda };
}
const html7 = (d, ex) => {
  const cat = (r, p) => td(fd(r.b), 'n') + td(fd(pc(r.b, r.t)), 'n') + td(fd(r.tb), 'n') + td(fd(pc(r.tb, r.t)), 'n') + (d.tda ? td(fd(r.u), 'n') + td(fd(pc(r.u, r.t)), 'n') : '') + td(fd(r.t), 'n');
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=3')}${th('Kabupaten/Kota', 'rowspan=3')}${th('Tahun Lama - Tahun Baru', 'rowspan=3')}${th('Perubahan Penggunaan Tanah', `colspan=${d.tda ? 6 : 4}`)}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Berubah', 'colspan=2')}${th('Tidak Berubah', 'colspan=2')}${d.tda ? th(TDA_LABEL, 'colspan=2') : ''}</tr><tr>${(d.tda ? [1, 2, 3] : [1, 2]).map(() => th('Luas (Ha)') + th('Luas (%)')).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${td(r.thn, 'c')}${cat(r)}</tr>`);
  h += `<tr class=tot><td colspan=3 class=c>Total</td>${cat(d.tot)}</tr></table></div>${cvTag(ex)}</div>`;
  return h;
};
const chart7 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: {
    labels: d.rows.map(r => r.k), datasets: [
      { label: 'Berubah', data: d.rows.map(r => r.b), backgroundColor: '#ffc000' },
      { label: 'Tidak Berubah', data: d.rows.map(r => r.tb), backgroundColor: NAVY }].concat(d.tda ? [{ label: TDA_LABEL, data: d.rows.map(r => r.u), backgroundColor: TDA_COLOR }] : [])
  },
  options: { ...baseOpt('PERUBAHAN PENGGUNAAN TANAH PER KECAMATAN TAHUN 2014-2026 (HA)', a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel7(wb, ws, d, ex) {
  const nC = d.tda ? 3 : 2, cT = 4 + nC * 2, C = colLetter, TL = C(cT);   // kolom: 1 No, 2 Kab/Kota, 3 Tahun Lama - Tahun Baru, 4.. luas & %, terakhir Jumlah
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kabupaten/Kota'); H(ws, 1, 3, 3, 3, 'Tahun Lama - Tahun Baru'); H(ws, 1, 4, 1, cT - 1, 'Perubahan Penggunaan Tanah');
  ['Berubah', 'Tidak Berubah'].concat(d.tda ? [TDA_LABEL] : []).forEach((t, i) => H(ws, 2, 4 + i * 2, 2, 5 + i * 2, t));
  H(ws, 1, cT, 3, cT, 'Jumlah (Ha)');
  for (let i = 0; i < nC * 2; i++) H(ws, 3, 4 + i, 3, 4 + i, i % 2 ? 'Luas (%)' : 'Luas (Ha)');
  // ===== RUMUS =====
  const fL = rawRng('LUASHA'), fG = rawRng('GQREKLAS'), fK = rawRng('WADMKC'), canF = !!(fL && fG && fK);
  const r1 = 4, rL = 3 + d.rows.length, n = rL + 1;                 // baris data pertama, terakhir, total
  const sum = (x) => Array.from({ length: nC }, (_, c) => `${C(4 + c * 2)}${x}`).join('+');
  d.rows.forEach((r, i) => {
    const x = r1 + i, ok = canF && r.k !== '-', V = [r.b, r.tb, r.u];
    cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.k); cell(ws, x, 3, r.thn, { ctr: 1 });
    // Berubah = baris GQREKLAS terisi yang BUKAN "Tidak..." dan bukan "...dianalisis"; Tidak Berubah = diawali "Tidak" (bukan "...dianalisis"); Tidak Dapat Dianalisis = mengandung "dianalis"
    const F = [`SUMIFS(${fL},${fK},$B${x},${fG},"<>",${fG},"<>Tidak*",${fG},"<>*analis*")`, `SUMIFS(${fL},${fK},$B${x},${fG},"Tidak*",${fG},"<>*analis*")`, `SUMIFS(${fL},${fK},$B${x},${fG},"*analis*")`];
    for (let c = 0; c < nC; c++) {
      cell(ws, x, 4 + c * 2, ok ? FX(F[c], V[c]) : V[c], { num: 1 });
      cell(ws, x, 5 + c * 2, FX(`IF($${TL}${x}=0,0,${C(4 + c * 2)}${x}/$${TL}${x}*100)`, pc(V[c], r.t)), { num: 1 });
    }
    cell(ws, x, cT, FX(sum(x), r.t), { num: 1, b: 1 });
  });
  const t = d.tot, TV = [t.b, t.tb, t.u];
  cell(ws, n, 1, 'Total', { b: 1, ctr: 1 }); cell(ws, n, 2, null); cell(ws, n, 3, null); ws.mergeCells(n, 1, n, 3);
  for (let c = 0; c < nC; c++) {
    cell(ws, n, 4 + c * 2, FX(`SUM(${C(4 + c * 2)}${r1}:${C(4 + c * 2)}${rL})`, TV[c]), { num: 1, b: 1 });
    cell(ws, n, 5 + c * 2, FX(`IF($${TL}${n}=0,0,${C(4 + c * 2)}${n}/$${TL}${n}*100)`, pc(TV[c], t.t)), { num: 1, b: 1 });
  }
  cell(ws, n, cT, FX(`SUM(${TL}${r1}:${TL}${rL})`, t.t), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 24; ws.getColumn(3).width = 20; for (let c = 4; c <= cT; c++) ws.getColumn(c).width = 14;
  // ===== GRAFIK ASLI =====
  const q = `'${ws.name.replace(/'/g, "''")}'!`, R = c => `${q}$${c}$${r1}:$${c}$${rL}`, cats = d.rows.map(r => r.k);
  const series = [{ name: 'Berubah', cat: R('B'), catVals: cats, val: R(C(4)), vals: d.rows.map(r => r.b), color: '#ffc000' },
                  { name: 'Tidak Berubah', cat: R('B'), catVals: cats, val: R(C(6)), vals: d.rows.map(r => r.tb), color: NAVY }];
  if (d.tda) series.push({ name: TDA_LABEL, cat: R('B'), catVals: cats, val: R(C(8)), vals: d.rows.map(r => r.u), color: TDA_COLOR });
  addXChart({ sheet: ws.name, type: 'bar', title: yr('PERUBAHAN PENGGUNAAN TANAH PER KECAMATAN TAHUN 2014-2026 (HA)'), legend: 'b', from: { col: cT + 1, row: 0 }, ext: { w: ex.w, h: ex.h }, series });
}

// =====================================================================
// SHEET 8 & 10 - Perbandingan 2 kolom (lama vs baru) + selisih
// =====================================================================
function buildCmp(fa, fb) {
  const a = sumBy(fa), b = sumBy(fb); if (!a || !b) return null;
  const cats = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort((x, y) => x.localeCompare(y));
  const rows = cats.map(c => ({ c, a: a[c] || 0, b: b[c] || 0, d: (b[c] || 0) - (a[c] || 0) }));
  const tot = rows.reduce((s, r) => ({ a: s.a + r.a, b: s.b + r.b, d: s.d + r.d }), { a: 0, b: 0, d: 0 });
  return { rows, tot };
}
const cmpTable = (d, title, h3) => {
  let h = `<table><tr>${th('No')}${th(title)}${th('Tahun 2014 (Ha)')}${th('Tahun 2026 (Ha)')}${th(h3)}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.c)}${td(fmt(r.a), 'n')}${td(fmt(r.b), 'n')}${td(fd(r.d), 'n')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${td(fmt(d.tot.a), 'n')}${td(fmt(d.tot.b), 'n')}${td(fd(d.tot.d), 'n')}</tr></table>`;
};
function cmpExcel(ws, d, title, h3, fa, fb) {      // fa/fb = nama kolom tahun lama / baru di Raw Data (mis. 'GNAME25','QNAME25')
  [['No', 1], [title, 2], ['Tahun 2014 (Ha)', 3], ['Tahun 2026 (Ha)', 4], [h3, 5]].forEach(([v, c]) => H(ws, 1, c, 1, c, v));
  const fL = rawRng('LUASHA'), ra = rawRng(fa), rb = rawRng(fb), canF = !!(fL && ra && rb);
  const rL = 1 + d.rows.length, n = rL + 1;
  d.rows.forEach((r, i) => {
    const x = 2 + i; cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.c);
    cell(ws, x, 3, canF ? FX(`SUMIFS(${fL},${ra},$B${x})`, r.a) : r.a, { num: 1 });      // luas tahun lama
    cell(ws, x, 4, canF ? FX(`SUMIFS(${fL},${rb},$B${x})`, r.b) : r.b, { num: 1 });      // luas tahun baru
    cell(ws, x, 5, FX(`D${x}-C${x}`, Math.abs(r.d) < 0.005 ? 0 : r.d), { num: 1 });      // selisih
  });
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  cell(ws, n, 3, FX(`SUM(C2:C${rL})`, d.tot.a), { num: 1, b: 1 }); cell(ws, n, 4, FX(`SUM(D2:D${rL})`, d.tot.b), { num: 1, b: 1 });
  cell(ws, n, 5, FX(`SUM(E2:E${rL})`, Math.abs(d.tot.d) < 0.005 ? 0 : d.tot.d), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 44; for (let c = 3; c <= 5; c++) ws.getColumn(c).width = 20;
  return n;
}
// --- Sheet 8 ---
const build8 = () => buildCmp('GNAME25', 'QNAME25');
const html8 = d => cmpTable(d, 'Penggunaan Tanah', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)');
const excel8 = (wb, ws, d) => { cmpExcel(ws, d, 'Penggunaan Tanah', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)', 'GNAME25', 'QNAME25'); };

// =====================================================================
// SHEET 9 - Perubahan terluas (GQNAME)
// =====================================================================
function build9() {
  if (!ROWS.length || !COLS.includes('GQNAME')) return null;
  const m = {};
  for (const r of ROWS) {
    const c = r.GQNAME; if (!c) continue;
    const p = c.split(' menjadi ');
    if (p.length === 2 && p[0].trim() === p[1].trim()) continue;      // "A menjadi A" = tidak berubah, dilewati
    m[c] = (m[c] || 0) + area(r);
  }
  const top = Object.entries(m).map(([c, t]) => ({ c, t })).sort((a, b) => b.t - a.t).slice(0, 5);
  if (!top.length) return null;
  return { top, T: top.reduce((s, x) => s + x.t, 0) };
}
const html9 = (d, ex) => `<div class=wrap><div><table><tr>${th('No')}${th('Perubahan Penggunaan Tanah 2014-2026 Terluas')}${th('Luas (Ha)')}</tr>`
  + d.top.map((x, i) => `<tr>${td(i + 1, 'c')}${td(x.c)}${td(fmt(x.t), 'n')}</tr>`).join('')
  + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${td(fmt(d.T), 'n')}</tr></table></div>${cvTag(ex)}</div>`;
const chart9 = (cv, d, a) => mkChart(cv, { type: 'pie', ct: 'PERUBAHAN PENGGUNAAN TANAH DOMINAN TAHUN 2014-2026 (%)' }, { body: d.top }, a);
function excel9(wb, ws, d, ex) {
  H(ws, 1, 1, 1, 1, 'No'); H(ws, 1, 2, 1, 2, 'Perubahan Penggunaan Tanah 2014-2026 Terluas'); H(ws, 1, 3, 1, 3, 'Luas (Ha)');
  const fL = rawRng('LUASHA'), fG = rawRng('GQNAME'), nt = d.top.length, rL = 1 + nt, n = rL + 1;
  d.top.forEach((x, i) => {
    cell(ws, 2 + i, 1, i + 1, { ctr: 1 }); cell(ws, 2 + i, 2, x.c);
    cell(ws, 2 + i, 3, fL && fG ? FX(`SUMIFS(${fL},${fG},$B${2 + i})`, x.t) : x.t, { num: 1 });   // luas tiap perubahan
  });
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  cell(ws, n, 3, FX(`SUM(C2:C${rL})`, d.T), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 70; ws.getColumn(3).width = 16;
  const q = `'${ws.name.replace(/'/g, "''")}'!`;
  addXChart({ sheet: ws.name, type: 'pie', title: yr('PERUBAHAN PENGGUNAAN TANAH DOMINAN TAHUN 2014-2026 (%)'), legend: 'r', from: { col: 3, row: 1 }, ext: { w: ex.w, h: ex.h },
    pointColors: d.top.map((_, i) => PAL[i % PAL.length]),
    series: [{ cat: `${q}$B$2:$B$${rL}`, catVals: d.top.map(x => x.c), val: `${q}$C$2:$C$${rL}`, vals: d.top.map(x => x.t) }] });
}

// =====================================================================
// SHEET 10 - Reklasifikasi (GREKLAS vs QREKLAS)
// =====================================================================
const build10 = () => buildCmp('GREKLAS', 'QREKLAS');
const html10 = (d, ex) => {
  const side = `<table><tr>${th('Penggunaan Tanah Reklasifikasi')}${d.rows.map(r => th(r.c)).join('')}</tr>
   <tr><th>${yr('Perubahan Penggunaan Tanah 2014-2026 (Ha)')}</th>${d.rows.map(r => td(fd(r.d), 'n')).join('')}</tr></table>`;
  return `<div class=wrap><div>${cmpTable(d, 'Penggunaan Tanah Reklasifikasi', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)')}</div><div class=side>${side}${cvTag(ex)}</div></div>`;
};
const chart10 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: [yr('Perubahan Penggunaan Tanah 2014-2026 (Ha)')], datasets: d.rows.map((r, i) => ({ label: r.c, data: [r.d], backgroundColor: PAL2[i % PAL2.length] })) },
  options: { ...baseOpt('PERUBAHAN PENGGUNAAN TANAH REKLASIFIKASI TAHUN 2014-2026 (HA)', a, 'right'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel10(wb, ws, d, ex) {
  const n = cmpExcel(ws, d, 'Penggunaan Tanah Reklasifikasi', 'Perubahan Penggunaan Tanah 2014-2026 (Ha)', 'GREKLAS', 'QREKLAS');
  // tabel samping (transpose): kategori jadi kolom, nilainya MENGAMBIL dari kolom selisih di tabel utama
  H(ws, 1, 7, 1, 7, 'Penggunaan Tanah Reklasifikasi'); H(ws, 2, 7, 2, 7, 'Perubahan Penggunaan Tanah 2014-2026 (Ha)');
  d.rows.forEach((r, i) => { H(ws, 1, 8 + i, 1, 8 + i, r.c); cell(ws, 2, 8 + i, FX(`E${2 + i}`, Math.abs(r.d) < 0.005 ? 0 : r.d), { num: 1 }); });
  ws.getColumn(7).width = 40; for (let i = 0; i < d.rows.length; i++) ws.getColumn(8 + i).width = 16;
  const q = `'${ws.name.replace(/'/g, "''")}'!`;
  addXChart({ sheet: ws.name, type: 'bar', title: yr('PERUBAHAN PENGGUNAAN TANAH REKLASIFIKASI TAHUN 2014-2026 (HA)'), legend: 'r', from: { col: 6, row: 3 }, ext: { w: ex.w, h: ex.h },
    series: d.rows.map((r, i) => ({ name: r.c, nameRef: `${q}${A1(1, 8 + i, 1, 1)}`, cat: `${q}$G$2`, catVals: [yr('Perubahan Penggunaan Tanah 2014-2026 (Ha)')],
      val: `${q}${A1(2, 8 + i, 1, 1)}`, vals: [Math.abs(r.d) < 0.005 ? 0 : r.d], color: PAL2[i % PAL2.length] })) });
}

// =====================================================================
// SHEET 28_Tabel_IV-17 - Potensi Cadangan Karbon
//   Kelas   : STD_C_G (2014) dan STD_C_Q (2026)
//   Nilai   : TON_C_G (2014) dan TON_C_Q (2026)
// =====================================================================
function sumByVal(catCol, valCol) {
  if (!ROWS.length || !COLS.includes(catCol) || !COLS.includes(valCol)) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const c = r[catCol]; if (!c) continue; any = true;
    m[c] = (m[c] || 0) + (parseFloat(r[valCol]) || 0);
  }
  return any ? m : null;
}
const ket = v => Math.abs(v) < 0.005 ? 'Tetap' : v > 0 ? 'Bertambah' : 'Berkurang';

function build17() {
  const a = sumByVal('STD_C_G', 'TON_C_G'), b = sumByVal('STD_C_Q', 'TON_C_Q');
  if (!a || !b) return null;
  const cats = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort((x, y) => x.localeCompare(y));
  const rows = cats.map(c => ({ c, a: a[c] || 0, b: b[c] || 0, d: (b[c] || 0) - (a[c] || 0) }));
  const tot = rows.reduce((s, r) => ({ a: s.a + r.a, b: s.b + r.b, d: s.d + r.d }), { a: 0, b: 0, d: 0 });
  return { rows, tot };
}
const html17 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Kelas Penutup Lahan', 'rowspan=2')}${th('Potensi Cadangan Karbon (Ton)', 'colspan=3')}${th('Keterangan', 'rowspan=2')}</tr>
  <tr>${th('Tahun 2014')}${th('Tahun 2026')}${th('Perubahan')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.c)}${td(fmt(r.a), 'n')}${td(fmt(r.b), 'n')}${td(fd(r.d), 'n')}${td(ket(r.d), 'c')}</tr>`);
  const t = d.tot;
  h += `<tr class=tot><td colspan=2 class=c>Jumlah (Ton)</td>${td(fmt(t.a), 'n')}${td(fmt(t.b), 'n')}${td(fd(t.d), 'n')}${td(ket(t.d), 'c')}</tr></table></div>`;
  const side = `<table><tr>${th('Kelas Penutup Lahan')}${d.rows.map(r => th(r.c)).join('')}</tr>
   <tr><th>Perubahan</th>${d.rows.map(r => td(fd(r.d), 'n')).join('')}</tr></table>`;
  return h + `<div class=side>${side}${cvTag(ex)}</div></div>`;
};
const chart17 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: ['Perubahan (Ton)'], datasets: d.rows.map((r, i) => ({ label: r.c, data: [r.d], backgroundColor: PAL2[i % PAL2.length] })) },
  options: { ...baseOpt('PERUBAHAN POTENSI CADANGAN KARBON TAHUN 2014-2026 (TON)', a, 'right'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel17(wb, ws, d, ex) {
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Kelas Penutup Lahan'); H(ws, 1, 3, 1, 5, 'Potensi Cadangan Karbon (Ton)');
  H(ws, 2, 3, 2, 3, 'Tahun 2014'); H(ws, 2, 4, 2, 4, 'Tahun 2026'); H(ws, 2, 5, 2, 5, 'Perubahan'); H(ws, 1, 6, 2, 6, 'Keterangan');
  const z = v => Math.abs(v) < 0.005 ? 0 : v;
  const tG = rawRng('TON_C_G'), sG = rawRng('STD_C_G'), tQ = rawRng('TON_C_Q'), sQ = rawRng('STD_C_Q'), canF = !!(tG && sG && tQ && sQ);
  const ketF = x => `IF(ABS(E${x})<0.005,"Tetap",IF(E${x}>0,"Bertambah","Berkurang"))`;
  const r1 = 3, rL = 2 + d.rows.length, n = rL + 1;
  d.rows.forEach((r, i) => {
    const x = r1 + i;
    cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.c);
    cell(ws, x, 3, canF ? FX(`SUMIFS(${tG},${sG},$B${x})`, r.a) : r.a, { num: 1 });
    cell(ws, x, 4, canF ? FX(`SUMIFS(${tQ},${sQ},$B${x})`, r.b) : r.b, { num: 1 });
    cell(ws, x, 5, FX(`D${x}-C${x}`, z(r.d)), { num: 1 });
    cell(ws, x, 6, FX(ketF(x), ket(r.d)), { ctr: 1 });
  });
  const t = d.tot;
  cell(ws, n, 1, 'Jumlah (Ton)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  cell(ws, n, 3, FX(`SUM(C${r1}:C${rL})`, t.a), { num: 1, b: 1 }); cell(ws, n, 4, FX(`SUM(D${r1}:D${rL})`, t.b), { num: 1, b: 1 });
  cell(ws, n, 5, FX(`SUM(E${r1}:E${rL})`, z(t.d)), { num: 1, b: 1 });
  cell(ws, n, 6, FX(ketF(n), ket(t.d)), { b: 1, ctr: 1 });
  H(ws, 1, 8, 1, 8, 'Kelas Penutup Lahan'); H(ws, 2, 8, 2, 8, 'Perubahan');
  d.rows.forEach((r, i) => { H(ws, 1, 9 + i, 1, 9 + i, r.c); cell(ws, 2, 9 + i, FX(`E${r1 + i}`, z(r.d)), { num: 1 }); });   // tabel samping = ambil dari kolom Perubahan
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 32; for (let c = 3; c <= 5; c++) ws.getColumn(c).width = 16; ws.getColumn(6).width = 14;
  ws.getColumn(8).width = 24; for (let i = 0; i < d.rows.length; i++) ws.getColumn(9 + i).width = 18;
  const q = `'${ws.name.replace(/'/g, "''")}'!`;
  addXChart({ sheet: ws.name, type: 'bar', title: yr('PERUBAHAN POTENSI CADANGAN KARBON TAHUN 2014-2026 (TON)'), legend: 'r', from: { col: 7, row: 3 }, ext: { w: ex.w, h: ex.h },
    series: d.rows.map((r, i) => ({ name: r.c, nameRef: `${q}${A1(1, 9 + i, 1, 1)}`, cat: `${q}$H$2`, catVals: ['Perubahan'],
      val: `${q}${A1(2, 9 + i, 1, 1)}`, vals: [z(r.d)], color: PAL2[i % PAL2.length] })) });
}

// =====================================================================
// SHEET 25_Tabel_IV-14, 26_Tabel_IV-15, 27_Tabel_IV-16 - Ketersediaan Tanah
//   Sumber : kolom V_ARAHAN x kabupaten/kota (WADMKC), luas dari LUASHA
// =====================================================================
// ---------- PENGATURAN (edit di sini) ----------
// IV-15: 2 jenis arahan Tanaman Pangan. 'key' = kata yang HARUS ada di teks V_ARAHAN (huruf kecil, tidak peka besar/kecil).
// IV-15: semua nilai V_ARAHAN yang mengandung SALAH SATU kata kunci di bawah dianggap potensi pertanian pangan
// (1 baris per nilai V_ARAHAN). Huruf kecil, tidak peka besar/kecil. Tambah kata kunci bila istilah RTRW/RDTR daerah berbeda.
// Nilai V_ARAHAN yang diawali "tidak tersedia" tidak pernah dihitung.
const KEY_PANGAN = ['pangan', 'sawah', 'padi', 'palawija', 'hortikultura', 'lp2b', 'kp2b', 'lahan basah', 'tegalan', 'ladang', 'semusim', 'lumbung', 'pertanian'];
const isPangan = a => { const t = norm(a); return !t.startsWith('tidak tersedia') && KEY_PANGAN.some(k => t.includes(k)) };
// IV-16: kelompok potensi sektor lain. Semua V_ARAHAN yang mengandung 'key' dimasukkan ke kelompok itu (1 baris per nilai V_ARAHAN).
// const SEKTOR_LAIN = [
//   { name: 'Potensi Perumahan', key: 'perumahan' },
//   { name: 'Potensi Perkebunan', key: 'perkebunan' },
//   { name: 'Potensi Hortikultura', key: 'hortikultura' },
//   { name: 'Potensi Peternakan', key: 'peternakan' },
//   { name: 'Potensi Pariwisata', key: 'pariwisata' },
//   { name: 'Potensi Perdagangan dan Jasa', key: 'perdagangan' },
//   { name: 'Potensi Perkantoran', key: 'perkantoran' },
//   { name: 'Potensi Industri', key: 'industri' },
//   { name: 'Potensi Pertambangan', key: 'pertambangan' },
//   { name: 'Potensi Fasilitas Umum dan Sosial', key: 'fasilitas umum' },
//   { name: 'Potensi Infrastruktur Perkotaan', key: 'infrastruktur' },
//   { name: 'Potensi Transportasi', key: 'transportasi' },
//   { name: 'Potensi Pertahanan dan Keamanan', key: 'pertahanan' }
// ];
const SEKTOR_LAIN = [
  { name: 'Potensi Pengembangan Sektor Perkebunan', key: ['perkebunan', 'kebun', 'tanaman tahunan'] },
  { name: 'Potensi Pengembangan Sektor Permukiman/Perumahan', key: ['perumahan', 'permukiman', 'pemukiman', 'hunian'] },
  { name: 'Potensi Pengembangan Sektor Industri', key: ['industri', 'pergudangan'] },
  { name: 'Potensi Pengembangan Sektor Pertambangan', key: ['pertambangan', 'tambang', 'galian', 'mineral'] },
  { name: 'Potensi Pengembangan Sektor Pariwisata', key: ['pariwisata', 'wisata', 'rekreasi'] }
];
const COL_BAR = ['#ffc000', '#1f3864', '#70ad47', '#5b9bd5', '#ed7d31', '#a5a5a5', '#7030a0', '#00b0a0', '#c55a11', '#2e75b6', '#bf9000', '#548235', '#843c0c', '#44546a'];
const PIE_COL = ['#70ad47', '#5b9bd5', '#ffc000', '#ed7d31', '#7030a0', '#00b0a0', '#2e75b6', '#bf9000', '#843c0c', '#44546a', '#a5a5a5', '#1f3864', '#c55a11', '#548235', '#ff9da7'];
const SISA_COL = '#c00000';                                                           // warna "Sisa Luas"

// ---------- Helper ----------
const norm = s => String(s || '').toLowerCase().split(' ').filter(Boolean).join(' ');
const hasAll = (a, keys) => keys.every(k => norm(a).includes(k));
const sumArr = v => v.reduce((a, b) => a + b, 0);
// luas per (V_ARAHAN x kabupaten/kota) -> {kecs, m:{arahan:{kec:luas}}} atau null
function arahanMatrix() {
  if (!ROWS.length || !COLS.includes('V_ARAHAN')) return null;
  const kecs = [...new Set(ROWS.map(kecOf))].sort(), m = {};
  for (const r of ROWS) {
    const a = (r.V_ARAHAN || '').trim(); if (!a) continue;
    const k = kecOf(r); (m[a] ??= {})[k] = (m[a][k] || 0) + area(r);
  }
  return Object.keys(m).length ? { kecs, m } : null;
}
const sumKec = (mx, names) => mx.kecs.map(k => names.reduce((s, a) => s + (mx.m[a][k] || 0), 0));
// nama wilayah (kota/kabupaten) diambil dari kolom WADMKK; jika ada >1 nilai, digabung dengan koma
function wilayah() {
  if (!ROWS.length || !COLS.includes('WADMKK')) return 'Wilayah';
  const m = {}; ROWS.forEach(r => { const w = (r.WADMKK || '').trim(); if (w) m[w] = (m[w] || 0) + 1; });
  const k = Object.keys(m).sort((a, b) => m[b] - m[a]);
  if (k.length > 1) return (window.NAMA_KAB ? 'Provinsi ' + window.NAMA_KAB.replace(/_/g, ' ') : 'Provinsi');   // data gabungan banyak kab/kota -> cukup nama provinsi
  return k.length ? k[0] : 'Wilayah';
}

const tdr = v => `<td class="n"><b>${fd(v)}</b></td>`;

// ---------- IV-15 : Potensi Pertanian Pangan ----------
function build15() {
  const mx = arahanMatrix(); if (!mx) return null;
  const names = Object.keys(mx.m).filter(isPangan).sort((x, y) => x.localeCompare(y));
  if (!names.length) return null;
  const rows = names.map(a => { const v = sumKec(mx, [a]); return { label: a, v, t: sumArr(v) } });
  const kt = mx.kecs.map((_, i) => rows.reduce((s, r) => s + r.v[i], 0));
  return { kecs: mx.kecs, rows, kt, T: sumArr(kt), wil: wilayah() };
}
const html15 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Arahan Ketersediaan', 'rowspan=2')}${th('Kabupaten/Kota (Ha)', 'colspan=' + d.kecs.length)}${th(d.wil + ' (Ha)', 'rowspan=2')}</tr>
  <tr>${d.kecs.map(k => th(k)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.label)}${r.v.map(v => td(fd(v), 'n')).join('')}${tdr(r.t)}</tr>`);
  h += `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${d.kt.map(v => td(fd(v), 'n')).join('')}${td(fd(d.T), 'n')}</tr></table></div><div class=side>${cvTag(ex)}</div></div>`;
  return h;
};
const chart15 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: d.kecs, datasets: d.rows.map((r, i) => ({ label: r.label, data: r.v, backgroundColor: COL_BAR[i % COL_BAR.length] })) },
  options: { ...baseOpt('KETERSEDIAAN TANAH UNTUK POTENSI PERTANIAN PANGAN TAHUN 2026 (HA)', a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel15(wb, ws, d, ex) {
  const nk = d.kecs.length, cT = 3 + nk, C = colLetter;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Arahan Ketersediaan'); H(ws, 1, 3, 1, 2 + nk, 'Kabupaten/Kota (Ha)');
  d.kecs.forEach((k, i) => H(ws, 2, 3 + i, 2, 3 + i, k)); H(ws, 1, cT, 2, cT, d.wil + ' (Ha)');
  const fL = rawRng('LUASHA'), fV = rawRng('V_ARAHAN'), fK = rawRng('WADMKC'), canF = !!(fL && fV && fK);
  const r1 = 3, rL = 2 + d.rows.length, n = rL + 1;
  d.rows.forEach((r, i) => {
    const x = r1 + i;
    cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.label).alignment = { vertical: 'middle', wrapText: true };
    r.v.forEach((v, j) => cell(ws, x, 3 + j, canF && d.kecs[j] !== '-' ? FX(`SUMIFS(${fL},${fV},$B${x},${fK},${C(3 + j)}$2)`, v) : v, { num: 1 }));
    cell(ws, x, cT, FX(`SUM(C${x}:${C(2 + nk)}${x})`, r.t), { num: 1, b: 1 });
  });
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  d.kt.forEach((v, j) => cell(ws, n, 3 + j, FX(`SUM(${C(3 + j)}${r1}:${C(3 + j)}${rL})`, v), { num: 1, b: 1 }));
  cell(ws, n, cT, FX(`SUM(${C(cT)}${r1}:${C(cT)}${rL})`, d.T), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 55; for (let c = 3; c <= cT; c++) ws.getColumn(c).width = 15;
  const q = `'${ws.name.replace(/'/g, "''")}'!`, cat = `${q}$C$2:$${C(2 + nk)}$2`;
  addXChart({ sheet: ws.name, type: 'bar', legend: 'b', from: { col: cT + 1, row: 0 }, ext: { w: ex.w, h: ex.h }, title: 'KETERSEDIAAN TANAH UNTUK POTENSI PERTANIAN PANGAN TAHUN 2026 (HA)',
    series: d.rows.map((r, i) => ({ name: r.label, nameRef: `${q}$B$${r1 + i}`, cat, catVals: d.kecs, val: `${q}$C$${r1 + i}:$${C(2 + nk)}$${r1 + i}`, vals: r.v, color: COL_BAR[i % COL_BAR.length] })) });
}

// ---------- IV-16 : Potensi Sektor Lain ----------
function build16() {
  const mx = arahanMatrix(); if (!mx) return null;
  const all = Object.keys(mx.m), used = new Set(), groups = [];
  all.forEach(a => { if (isPangan(a)) used.add(a); });   // milik IV-15, jangan dobel
  SEKTOR_LAIN.forEach(s => {
    const names = all.filter(a => !used.has(a) && s.key.some(k => norm(a).includes(k))).sort((x, y) => x.localeCompare(y));
    if (!names.length) return;
    names.forEach(a => used.add(a));
    const items = names.map(a => { const v = sumKec(mx, [a]); return { label: a, v, t: sumArr(v) }; });
    const sub = mx.kecs.map((_, i) => items.reduce((x, r) => x + r.v[i], 0));
    groups.push({ name: s.name, items, sub, t: sumArr(sub) });
  });
  if (!groups.length) return null;
  const kt = mx.kecs.map((_, i) => groups.reduce((x, g) => x + g.sub[i], 0));
  return { kecs: mx.kecs, groups, kt, T: sumArr(kt), wil: wilayah(), unused: all.filter(a => !used.has(a) && !norm(a).startsWith('tidak tersedia')).sort() };
}
const html16 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Potensi Sektor Lain', 'rowspan=2')}${th('Arahan Ketersediaan', 'rowspan=2')}${th('Kabupaten/Kota (Ha)', 'colspan=' + d.kecs.length)}${th(d.wil + ' (Ha)', 'rowspan=2')}</tr>
  <tr>${d.kecs.map(k => th(k)).join('')}</tr>`;
  d.groups.forEach((g, gi) => {
    g.items.forEach((it, i) => h += `<tr>${i === 0 ? `<td class=c rowspan=${g.items.length + 1}>${gi + 1}</td><td class=c rowspan=${g.items.length + 1}>${g.name}</td>` : ''}${td(it.label)}${it.v.map(v => td(fd(v), 'n')).join('')}${tdr(it.t)}</tr>`);
    h += `<tr class=tot>${td(g.name + ' (Ha)')}${g.sub.map(v => td(fd(v), 'n')).join('')}${td(fd(g.t), 'n')}</tr>`;
  });
  h += `<tr class=tot><td colspan=3 class=c>Jumlah (Ha)</td>${d.kt.map(v => td(fd(v), 'n')).join('')}${td(fd(d.T), 'n')}</tr></table>`;
  if (d.unused.length) h += `<div class=note>Nilai V_ARAHAN yang tidak masuk IV-15 maupun IV-16: ${d.unused.join('; ')}</div>`;
  return h + `</div><div class=side>${cvTag(ex)}</div></div>`;
};
const chart16 = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: d.kecs, datasets: d.groups.map((g, i) => ({ label: g.name, data: g.sub, backgroundColor: COL_BAR[i % COL_BAR.length] })) },
  options: { ...baseOpt('KETERSEDIAAN TANAH UNTUK POTENSI SEKTOR LAIN TAHUN 2026 (HA)', a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excel16(wb, ws, d, ex) {
  const nk = d.kecs.length, cT = 4 + nk, C = colLetter;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Potensi Sektor Lain'); H(ws, 1, 3, 2, 3, 'Arahan Ketersediaan');
  H(ws, 1, 4, 1, 3 + nk, 'Kabupaten/Kota (Ha)'); d.kecs.forEach((k, i) => H(ws, 2, 4 + i, 2, 4 + i, k)); H(ws, 1, cT, 2, cT, d.wil + ' (Ha)');
  const mid = { horizontal: 'center', vertical: 'middle', wrapText: true }, wrap = { vertical: 'middle', wrapText: true };
  const fL = rawRng('LUASHA'), fV = rawRng('V_ARAHAN'), fK = rawRng('WADMKC'), canF = !!(fL && fV && fK), subs = [];
  let r = 3;
  d.groups.forEach((g, gi) => {
    const r0 = r, rs = r0 + g.items.length;                       // rs = baris subtotal
    for (let x = r0; x <= rs; x++) { cell(ws, x, 1, null); cell(ws, x, 2, null); }
    ws.getCell(r0, 1).value = gi + 1; ws.getCell(r0, 1).alignment = mid;
    ws.getCell(r0, 2).value = g.name; ws.getCell(r0, 2).alignment = mid;
    ws.mergeCells(r0, 1, rs, 1); ws.mergeCells(r0, 2, rs, 2);
    g.items.forEach((it, i) => {
      const n = r0 + i;
      cell(ws, n, 3, it.label).alignment = wrap;
      it.v.forEach((v, j) => cell(ws, n, 4 + j, canF && d.kecs[j] !== '-' ? FX(`SUMIFS(${fL},${fV},$C${n},${fK},${C(4 + j)}$2)`, v) : v, { num: 1 }));
      cell(ws, n, cT, FX(`SUM(D${n}:${C(3 + nk)}${n})`, it.t), { num: 1, b: 1 });
    });
    cell(ws, rs, 3, g.name + ' (Ha)', { b: 1 });
    g.sub.forEach((v, j) => cell(ws, rs, 4 + j, FX(`SUM(${C(4 + j)}${r0}:${C(4 + j)}${rs - 1})`, v), { num: 1, b: 1 }));
    cell(ws, rs, cT, FX(`SUM(${C(cT)}${r0}:${C(cT)}${rs - 1})`, g.t), { num: 1, b: 1 });
    subs.push({ r0, rs }); r = rs + 1;
  });
  cell(ws, r, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3);
  d.kt.forEach((v, j) => cell(ws, r, 4 + j, FX(subs.map(s => `${C(4 + j)}${s.rs}`).join('+'), v), { num: 1, b: 1 }));
  cell(ws, r, cT, FX(subs.map(s => `${C(cT)}${s.rs}`).join('+'), d.T), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 26; ws.getColumn(3).width = 60; for (let c = 4; c <= cT; c++) ws.getColumn(c).width = 15;
  const q = `'${ws.name.replace(/'/g, "''")}'!`, cat = `${q}$D$2:$${C(3 + nk)}$2`;
  addXChart({ sheet: ws.name, type: 'bar', legend: 'b', from: { col: cT + 1, row: 0 }, ext: { w: ex.w, h: ex.h }, title: 'KETERSEDIAAN TANAH UNTUK POTENSI SEKTOR LAIN TAHUN 2026 (HA)',
    series: d.groups.map((g, i) => ({ name: g.name, nameRef: `${q}$B$${subs[i].r0}`, cat, catVals: d.kecs, val: `${q}$D$${subs[i].rs}:$${C(3 + nk)}$${subs[i].rs}`, vals: g.sub, color: COL_BAR[i % COL_BAR.length] })) });
}

// ---------- IV-14 : Ringkasan Potensi Sektoral ----------
function build14() {
  const p = build15(), s = build16(); if (!p && !s) return null;
  const total = ROWS.reduce((a, r) => a + area(r), 0), wil = wilayah();   // luas seluruh wilayah (semua baris)           // luas seluruh wilayah (semua baris)
  const items = [{ c: 'Potensi Pertanian Pangan', t: p ? p.T : 0 }];
  if (s) s.groups.forEach(g => items.push({ c: g.name, t: g.t }));
  items.push({ c: 'Sisa Luas ' + wil, t: total - sumArr(items.map(x => x.t)) });
  items.forEach(x => x.p = pc(x.t, total));
  return { items, total, wil };
}
const html14 = (d, ex) => {
  let h = `<div class=wrap><div><table><tr>${th('No')}${th('Potensi Sektoral')}${th('Ha')}${th('% Wilayah')}</tr>`;
  d.items.forEach((x, i) => h += `<tr>${td(i + 1, 'c')}${td(x.c)}${td(fd(x.t), 'n')}${td(fd(x.p), 'n')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${td(fd(d.total), 'n')}${td('100.00', 'n')}</tr></table></div><div class=side>${cvTag(ex)}</div></div>`;
};
const pieVals = {
  id: 'pv', afterDatasetsDraw(ch) {
    const c = ch.ctx, ds = ch.data.datasets[0], tot = sumArr(ds.data);
    c.save(); c.font = 'bold 12px Arial'; c.fillStyle = '#fff'; c.textAlign = 'center';
    ch.getDatasetMeta(0).data.forEach((el, i) => { const p = tot ? ds.data[i] / tot * 100 : 0; if (p < 2) return; const t = el.tooltipPosition(); c.fillText(p.toFixed(2) + '%', t.x, t.y + 4); });
    c.restore();
  }
};
const chart14 = (cv, d, a) => {
  const n = d.items.length;
  return new Chart(cv, {
    type: 'pie',
    data: { labels: d.items.map(x => x.c), datasets: [{ data: d.items.map(x => Math.max(0, x.t)), backgroundColor: d.items.map((_, i) => i === n - 1 ? SISA_COL : PIE_COL[i % PIE_COL.length]) }] },
    options: baseOpt('KETERSEDIAAN TANAH UNTUK POTENSI SEKTORAL TAHUN 2026 (%)', a, 'right'), plugins: [bgPlugin, pieVals]
  });
};
function excel14(wb, ws, d, ex) {
  H(ws, 1, 1, 1, 1, 'No'); H(ws, 1, 2, 1, 2, 'Potensi Sektoral'); H(ws, 1, 3, 1, 3, 'Ha'); H(ws, 1, 4, 1, 4, '% Wilayah');
  // Pangan dan sektor lain mengambil total dari sheet Tabel IV-15 dan IV-16 (posisi sel dihitung dari susunan kedua sheet itu)
  const d15 = build15(), d16 = build16(), C = colLetter;
  const t15 = (EXTRA.find(e => e.build === build15) || {}).tab, t16 = (EXTRA.find(e => e.build === build16) || {}).tab;
  const ref15 = d15 ? `'${t15}'!${C(3 + d15.kecs.length)}${3 + d15.rows.length}` : null;
  const ref16 = []; if (d16) { let r = 3; d16.groups.forEach(g => { const rs = r + g.items.length; ref16.push(`'${t16}'!${C(4 + d16.kecs.length)}${rs}`); r = rs + 1; }); }
  const nI = d.items.length, rL = 1 + nI, rt = rL + 1, fL = rawRng('LUASHA');
  d.items.forEach((x, i) => {
    const n = 2 + i, last = i === nI - 1;
    const f = last ? `$C$${rt}-SUM(C2:C${n - 1})`                                  // Sisa luas = luas wilayah - jumlah potensi
      : i === 0 ? (ref15 || null) : ref16[i - 1];
    cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, x.c);
    cell(ws, n, 3, f ? FX(f, x.t) : x.t, { num: 1 });
    cell(ws, n, 4, FX(`IF($C$${rt}=0,0,C${n}/$C$${rt}*100)`, x.p), { num: 1 });
  });
  cell(ws, rt, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  cell(ws, rt, 3, fL ? FX(`SUM(${fL})`, d.total) : d.total, { num: 1, b: 1 });                  // luas seluruh wilayah dari Raw Data
  cell(ws, rt, 4, FX(`SUM(D2:D${rL})`, 100), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 40; ws.getColumn(3).width = 16; ws.getColumn(4).width = 14;
  const q = `'${ws.name.replace(/'/g, "''")}'!`;
  addXChart({ sheet: ws.name, type: 'pie', legend: 'r', from: { col: 5, row: 0 }, ext: { w: ex.w, h: ex.h }, title: 'KETERSEDIAAN TANAH UNTUK POTENSI SEKTORAL TAHUN 2026 (%)',
    pointColors: d.items.map((_, i) => i === nI - 1 ? SISA_COL : PIE_COL[i % PIE_COL.length]),
    series: [{ cat: `${q}$B$2:$B$${rL}`, catVals: d.items.map(x => x.c), val: `${q}$C$2:$C$${rL}`, vals: d.items.map(x => Math.max(0, x.t)) }] });
}

// =====================================================================
// SHEET 24_Tabel_IV-12 - Arahan Ketersediaan Tanah per Kabupaten/Kota x Pola Ruang RDTR/RTRW
//   Kabupaten/Kota = WADMKC | Pola Ruang = POLA_COL | Arahan = V_ARAHAN | Luas = LUASHA
//   Luas (%) = luas baris / jumlah seluruh baris tabel x 100
// =====================================================================
const POLA_COL = 'NAMOBJ';   // kolom Pola Ruang RDTR/RTRW (ganti jika nama kolomnya lain)

// angka kecil tampil lebih banyak desimal: 2 desimal; jika hasilnya 0.00 -> 4 desimal; jika masih 0.0000 -> 5, 6, dst
const decA = v => { let n = 2; while (n < 8 && v > 0 && Number(v.toFixed(n)) === 0) n += n === 2 ? 2 : 1; return n; };
const fmtA = v => v.toLocaleString('en-US', { minimumFractionDigits: decA(v), maximumFractionDigits: decA(v) });
const nfA = v => '#,##0.' + '0'.repeat(decA(v));

function build12() {
  if (!ROWS.length || !['V_ARAHAN', POLA_COL].every(c => COLS.includes(c))) return null;
  const m = {}; let total = 0;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), a = (r.V_ARAHAN || '').trim(); if (!p || !a) continue;
    const k = kecOf(r), v = area(r);
    m[k] ??= {}; m[k][p] ??= {}; m[k][p][a] = (m[k][p][a] || 0) + v; total += v;
  }
  if (!Object.keys(m).length) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(m).sort().map(k => {
    const polas = Object.keys(m[k]).sort(by).map(p => ({ p, items: Object.keys(m[k][p]).sort(by).map(a => ({ a, v: m[k][p][a] })) }));
    return { k, polas, n: polas.reduce((s, q) => s + q.items.length, 0) };
  });
  return { kecs, total, wil: wilayah() };
}
const html12 = d => {
  const top = 'style="vertical-align:top"';
  let h = `<div class=raw-scroll><table><tr>${th('No')}${th('Kabupaten/Kota')}${th('Pola Ruang RDTR/RTRW')}${th('Arahan Ketersediaan Tanah')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  d.kecs.forEach((kc, ki) => kc.polas.forEach((q, qi) => q.items.forEach((it, i) => {
    h += '<tr>';
    if (qi === 0 && i === 0) h += `<td class=c rowspan=${kc.n} ${top}>${ki + 1}</td><td rowspan=${kc.n} ${top}>${kc.k}</td>`;
    if (i === 0) h += `<td class=c rowspan=${q.items.length} ${top}>${q.p}</td>`;
    h += `${td(it.a)}${td(fmtA(it.v), 'n')}${td(fmtA(pc(it.v, d.total)), 'n')}</tr>`;
  })));
  return h + `<tr class=tot><td colspan=4 class=c>${d.wil}</td>${td(fd(d.total), 'n')}${td('100.00', 'n')}</tr></table></div>`;
};
function excel12(wb, ws, d) {
  ['No', 'Kabupaten/Kota', 'Pola Ruang RDTR/RTRW', 'Arahan Ketersediaan Tanah', 'Luas (Ha)', 'Luas (%)'].forEach((t, i) => H(ws, 1, i + 1, 1, i + 1, t));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, top = { vertical: 'top', wrapText: true };
  const num = (r, c, v, f) => { const x = cell(ws, r, c, f ? FX(f, v) : v, { num: 1 }); x.numFmt = nfA(v); };
  const fL = rawRng('LUASHA'), fK = rawRng('WADMKC'), fP = rawRng(POLA_COL), fV = rawRng('V_ARAHAN'), canF = !!(fL && fK && fP && fV);
  const rt = 2 + d.kecs.reduce((a, k) => a + k.n, 0);                                  // baris total
  let r = 2;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.polas.forEach(q => {
      const p0 = r;
      q.items.forEach(it => {
        for (let c = 1; c <= 3; c++) cell(ws, r, c, null);
        cell(ws, r, 4, it.a);
        num(r, 5, it.v, canF && kc.k !== '-' ? `SUMIFS(${fL},${fK},$B$${k0},${fP},$C$${p0},${fV},$D${r})` : null);   // luas per kabupaten/kota x pola ruang x arahan
        num(r, 6, pc(it.v, d.total), `IF($E$${rt}=0,0,E${r}/$E$${rt}*100)`); r++;
      });
      ws.getCell(p0, 3).value = q.p; ws.getCell(p0, 3).alignment = mid;
      if (r - 1 > p0) ws.mergeCells(p0, 3, r - 1, 3);
    });
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = top;
    if (r - 1 > k0) { ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2); }
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); for (let c = 2; c <= 4; c++) cell(ws, r, c, null); ws.mergeCells(r, 1, r, 4);
  cell(ws, r, 5, FX(`SUM(E2:E${r - 1})`, d.total), { num: 1, b: 1 }); cell(ws, r, 6, FX(`SUM(F2:F${r - 1})`, 100), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 22; ws.getColumn(3).width = 36; ws.getColumn(4).width = 75; ws.getColumn(5).width = 14; ws.getColumn(6).width = 12;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
}

// =====================================================================
// SHEET 23_Tabel_IV-11 - Ketersediaan Tanah (Tersedia / Tidak Tersedia) per Kabupaten/Kota x Pola Ruang RDTR/RTRW
//   Kabupaten/Kota = WADMKC | Pola Ruang = POLA_COL | Tersedia/Tidak Tersedia = VNAME | Luas = LUASHA
//   Luas (%) tiap baris = luas / Jumlah (Ha) baris itu x 100  (baris total: terhadap seluruh wilayah)
//   Butuh: POLA_COL dan norm() dari blok sebelumnya (IV-12 dan IV-14/15/16).
// =====================================================================
function build11() {
  const VCOL = COLS.find(c => c.toUpperCase() === 'VNAME');   // cocokkan tanpa peduli huruf besar/kecil
  if (!ROWS.length || ![VCOL, POLA_COL].every(c => COLS.includes(c))) return null;
  const m = {}; let any = false, tda = false;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), s = norm(r[VCOL]); if (!p || !s) continue;
    const i = isTDA(s) ? 2 : s.startsWith('tidak') ? 1 : s.includes('tersedia') ? 0 : -1; if (i < 0) continue;   // 0 = Tersedia, 1 = Tidak Tersedia, 2 = Tidak Dapat Dianalisis
    if (i === 2) tda = true;
    any = true; const k = kecOf(r);
    m[k] ??= {}; m[k][p] ??= [0, 0, 0]; m[k][p][i] += area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(m).sort().map(k => ({ k, rows: Object.keys(m[k]).sort(by).map(p => ({ p, ...ketMk(m[k][p][0], m[k][p][1], m[k][p][2]) })) }));
  const all = kecs.flatMap(kc => kc.rows);
  return { kecs, tot: ketMk(sumArr(all.map(r => r.a)), sumArr(all.map(r => r.b)), sumArr(all.map(r => r.u))), wil: wilayah(), tda: TDA_ALWAYS || tda };
}
const html11 = d => {
  const top = 'style="vertical-align:top"';
  let h = `<table><tr>${th('No', 'rowspan=3')}${th('Kabupaten/Kota', 'rowspan=3')}${th('Pola Ruang RDTR/RTRW', 'rowspan=3')}${th('Ketersediaan Tanah', `colspan=${d.tda ? 6 : 4}`)}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${ketGrp(d.tda)}</tr>
  <tr>${ketLuas(d.tda)}</tr>`;
  const cells = r => ketCells(r, d.tda);
  d.kecs.forEach((kc, ki) => kc.rows.forEach((r, i) => {
    h += '<tr>' + (i === 0 ? `<td class=c rowspan=${kc.rows.length} ${top}>${ki + 1}</td><td rowspan=${kc.rows.length} ${top}>${kc.k}</td>` : '') + td(r.p) + cells(r) + '</tr>';
  }));
  return h + `<tr class=tot><td colspan=3 class=c>${d.wil}</td>${cells(d.tot)}</tr></table>`;
};
function excel11(wb, ws, d) { return excelKetPola('Pola Ruang RDTR/RTRW', POLA_COL, true)(wb, ws, d); }

// =====================================================================
// SHEET 21_Tabel_IV-9  - Ketersediaan Tanah per Kabupaten/Kota (+ grafik)
// SHEET 22_Tabel_IV-10 - Ketersediaan Tanah per Kabupaten/Kota x Penggunaan Tanah (QNAME25)
//   Tersedia / Tidak Tersedia diambil dari kolom Vname (huruf besar/kecil bebas)
// =====================================================================
const vCol = () => COLS.find(c => c.toUpperCase() === 'VNAME');
// 0 = Tersedia, 1 = Tidak Tersedia, 2 = Tidak Dapat Dianalisis, -1 = tidak dikenali
const ketIdx = v => {
  const s = norm(String(v || '').replace(/[\s_\-]+/g, ' ')); if (!s) return -1;
  if (isTDA(s)) return 2;
  return /(^|\s)(tidak|tdk|belum)(\s|$)/.test(s) ? 1 : s.includes('tersedia') ? 0 : -1;
};
const ketMk = (a, b, u = 0) => { const t = a + b + u; return { a, b, u, t, pa: pc(a, t), pb: pc(b, t), pu: pc(u, t) }; };
// bagian tampilan: kolom "Tidak Dapat Dianalisis" hanya muncul bila ada di data (tda = true)
const ketGrp = tda => th('Tersedia', 'colspan=2') + th('Tidak Tersedia', 'colspan=2') + (tda ? th(TDA_LABEL, 'colspan=2') : '');
const ketLuas = tda => (tda ? [1, 2, 3] : [1, 2]).map(() => th('Luas (Ha)') + th('Luas (%)')).join('');
const ketCells = (r, tda) => `${td(fd(r.a), 'n')}${td(fd(r.pa), 'n')}${td(fd(r.b), 'n')}${td(fd(r.pb), 'n')}${tda ? td(fd(r.u), 'n') + td(fd(r.pu), 'n') : ''}${tdr(r.t)}`;
// bagian Excel: header "Ketersediaan Tanah" (3 baris) mulai kolom c0; mengembalikan nomor kolom Jumlah
function ketHeads(ws, c0, tda) {
  const nC = tda ? 3 : 2, cT = c0 + nC * 2;
  H(ws, 1, c0, 1, cT - 1, 'Ketersediaan Tanah');
  ['Tersedia', 'Tidak Tersedia'].concat(tda ? [TDA_LABEL] : []).forEach((t, i) => H(ws, 2, c0 + i * 2, 2, c0 + i * 2 + 1, t));
  for (let i = 0; i < nC * 2; i++) H(ws, 3, c0 + i, 3, c0 + i, i % 2 ? 'Luas (%)' : 'Luas (Ha)');
  H(ws, 1, cT, 3, cT, 'Jumlah (Ha)');
  return cT;
}
// tulis 1 baris data mulai kolom c0. f = [rumus Tersedia, Tidak Tersedia, TDA] (null = nilai statis)
function ketPut(ws, r, c0, o, f, tda, b) {
  const nC = tda ? 3 : 2, cT = c0 + nC * 2, C = colLetter, TL = C(cT), V = [o.a, o.b, o.u], P = [o.pa, o.pb, o.pu], S = [];
  for (let c = 0; c < nC; c++) {
    cell(ws, r, c0 + c * 2, f && f[c] ? FX(f[c], V[c]) : V[c], { num: 1, b });
    cell(ws, r, c0 + c * 2 + 1, FX(`IF($${TL}${r}=0,0,${C(c0 + c * 2)}${r}/$${TL}${r}*100)`, P[c]), { num: 1, b });
    S.push(`${C(c0 + c * 2)}${r}`);
  }
  cell(ws, r, cT, FX(S.join('+'), o.t), { num: 1, b: 1 });
  return cT;
}
const ketFs = (ex2, tda, strict) => ex2 ? [ketSumF(0, ex2, strict), ketSumF(1, ex2, strict)].concat(tda ? [ketSumF(2, ex2, strict)] : []) : null;
const ketSums = (c0, a, z, tda) => (tda ? [0, 2, 4] : [0, 2]).map(i => `SUM(${colLetter(c0 + i)}${a}:${colLetter(c0 + i)}${z})`);

// ---------- IV-9 : per kabupaten/kota ----------
function buildKetKec() {
  const V = vCol(); if (!ROWS.length || !V) return null;
  const m = {}; let any = false, tda = false;
  for (const r of ROWS) {
    const i = ketIdx(r[V]); if (i < 0) continue;
    if (i === 2) tda = true;
    any = true; const k = kecOf(r); m[k] ??= [0, 0, 0]; m[k][i] += area(r);
  }
  if (!any) return null;
  const rows = Object.keys(m).sort().map(k => ({ k, ...ketMk(m[k][0], m[k][1], m[k][2]) }));
  return { rows, tot: ketMk(sumArr(rows.map(r => r.a)), sumArr(rows.map(r => r.b)), sumArr(rows.map(r => r.u))), wil: wilayah(), tda: TDA_ALWAYS || tda };
}
const htmlKetKec = (d, ex) => {
  const cells = r => ketCells(r, d.tda);
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=3')}${th('Kabupaten/Kota', 'rowspan=3')}${th('Ketersediaan Tanah', `colspan=${d.tda ? 6 : 4}`)}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${ketGrp(d.tda)}</tr>
  <tr>${ketLuas(d.tda)}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${cells(r)}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${cells(d.tot)}</tr></table></div>${cvTag(ex)}</div>`;
};
const chartKetKec = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: {
    labels: d.rows.map(r => r.k), datasets: [
      { label: 'Tersedia', data: d.rows.map(r => r.a), backgroundColor: '#ffc000' },
      { label: 'Tidak Tersedia', data: d.rows.map(r => r.b), backgroundColor: NAVY }].concat(d.tda ? [{ label: TDA_LABEL, data: d.rows.map(r => r.u), backgroundColor: TDA_COLOR }] : [])
  },
  options: { ...baseOpt(['KETERSEDIAAN TANAH PER KECAMATAN', d.wil.toUpperCase() + ' TAHUN 2026 (HA)'], a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excelKetKec(wb, ws, d, ex) {
  const tda = d.tda, C = colLetter;
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kabupaten/Kota'); const cT = ketHeads(ws, 3, tda);
  const fK = rawRng('WADMKC'), r1 = 4, rL = 3 + d.rows.length, n = rL + 1;
  d.rows.forEach((r, i) => {
    const x = r1 + i, ex2 = fK && r.k !== '-' ? [[fK, `$B${x}`]] : null;
    cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.k); ketPut(ws, x, 3, r, ketFs(ex2, tda), tda);
  });
  cell(ws, n, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  ketPut(ws, n, 3, d.tot, ketSums(3, r1, rL, tda), tda, true);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 24; for (let c = 3; c <= cT; c++) ws.getColumn(c).width = 14;
  const q = `'${ws.name.replace(/'/g, "''")}'!`, cat = `${q}$B$${r1}:$B$${rL}`, cats = d.rows.map(r => r.k), R = c => `${q}$${C(c)}$${r1}:$${C(c)}$${rL}`;
  const series = [{ name: 'Tersedia', cat, catVals: cats, val: R(3), vals: d.rows.map(r => r.a), color: '#ffc000' },
                  { name: 'Tidak Tersedia', cat, catVals: cats, val: R(5), vals: d.rows.map(r => r.b), color: NAVY }];
  if (tda) series.push({ name: TDA_LABEL, cat, catVals: cats, val: R(7), vals: d.rows.map(r => r.u), color: TDA_COLOR });
  addXChart({ sheet: ws.name, type: 'bar', legend: 'b', from: { col: cT + 1, row: 0 }, ext: { w: ex.w, h: ex.h }, title: 'KETERSEDIAAN TANAH PER KECAMATAN ' + d.wil.toUpperCase() + ' TAHUN 2026 (HA)', series });
}

// ---------- IV-10 : per kabupaten/kota x Penggunaan Tanah (kolom col, mis. QNAME25) ----------
function buildKetPola(col) {
  const V = vCol(); if (!ROWS.length || !V || !COLS.includes(col)) return null;
  const m = {}; let any = false, tda = false;
  for (const r of ROWS) {
    const p = (r[col] || '').trim(), i = ketIdx(r[V]); if (!p || i < 0) continue;
    if (i === 2) tda = true;
    any = true; const k = kecOf(r); m[k] ??= {}; m[k][p] ??= [0, 0, 0]; m[k][p][i] += area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(m).sort().map(k => ({ k, rows: Object.keys(m[k]).sort(by).map(p => ({ p, ...ketMk(m[k][p][0], m[k][p][1], m[k][p][2]) })) }));
  const all = kecs.flatMap(kc => kc.rows);
  return { kecs, tot: ketMk(sumArr(all.map(r => r.a)), sumArr(all.map(r => r.b)), sumArr(all.map(r => r.u))), wil: wilayah(), tda: TDA_ALWAYS || tda };
}
const htmlKetPola = label => d => {
  const top = 'style="vertical-align:top"';
  const cells = r => ketCells(r, d.tda);
  let h = `<div class=raw-scroll><table><tr>${th('No', 'rowspan=3')}${th('Kabupaten/Kota', 'rowspan=3')}${th(label, 'rowspan=3')}${th('Ketersediaan Tanah', `colspan=${d.tda ? 6 : 4}`)}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${ketGrp(d.tda)}</tr>
  <tr>${ketLuas(d.tda)}</tr>`;
  d.kecs.forEach((kc, ki) => kc.rows.forEach((r, i) => {
    h += '<tr>' + (i === 0 ? `<td class=c rowspan=${kc.rows.length} ${top}>${ki + 1}</td><td rowspan=${kc.rows.length} ${top}>${kc.k}</td>` : '') + td(r.p) + cells(r) + '</tr>';
  }));
  return h + `<tr class=tot><td colspan=3 class=c>${d.wil}</td>${cells(d.tot)}</tr></table></div>`;
};
const excelKetPola = (label, col = 'QNAME25', strict = false) => (wb, ws, d) => {
  const tda = d.tda;
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kabupaten/Kota'); H(ws, 1, 3, 3, 3, label); const cT = ketHeads(ws, 4, tda);
  const fK = rawRng('WADMKC'), fP = rawRng(col);
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, left = { vertical: 'top', wrapText: true };
  let r = 4; const r1 = 4;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.rows.forEach(x => {
      cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, x.p).alignment = left;
      const ex2 = fK && fP && kc.k !== '-' ? [[fK, `$B$${k0}`], [fP, `$C${r}`]] : null;     // kabupaten/kota (sel gabungan) + kategori baris ini
      ketPut(ws, r, 4, x, ketFs(ex2, tda, strict), tda); r++;
    });
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = left;
    if (r - 1 > k0) { ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2); }
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3);
  ketPut(ws, r, 4, d.tot, ketSums(4, r1, r - 1, tda), tda, true);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 20; ws.getColumn(3).width = 38; for (let c = 4; c < cT; c++) ws.getColumn(c).width = 12; ws.getColumn(cT).width = 15;
  ws.views = [{ state: 'frozen', ySplit: 3 }];
};

// =====================================================================
// SHEET 20_Tabel_IV-8 - Kesesuaian Penggunaan Tanah terhadap RDTR per Kabupaten/Kota x Pola Ruang
// SHEET 19_Tabel_IV-7 - Matriks Penggunaan Tanah (QNAME25) x Arahan Fungsi Kawasan (NAMOBJ) = S / T / M
//   Kesesuaian diambil dari kolom KSPOLA (atau KS_POLA): Sesuai (S), Tidak Sesuai (T), Mendukung (M)
// =====================================================================
const MENDUKUNG_SESUAI = true;   // true: "Mendukung" ikut dihitung ke kolom "Sesuai" di IV-8; false: "Mendukung" tidak dihitung
const ksCol = () => COLS.find(c => ['KSPOLA', 'KS_POLA'].includes(c.toUpperCase()));
const ksCode = v => {            // -> 'S' | 'T' | 'M' | 'D' (Tidak Dapat Dianalisis) | '' (tidak dikenali)
  const s = norm(String(v || '').replace(/[\s_\-]+/g, ' ')); if (!s) return '';
  if (isTDA(s)) return 'D';      // dicek lebih dulu: "Tidak Dapat Dianalisis" mengandung kata "tidak" tetapi BUKAN "Tidak Sesuai"
  if (/(^|\s)(tidak|tdk)(\s|$)/.test(s)) return 'T';
  if (s.includes('mendukung')) return 'M';
  return s.includes('sesuai') ? 'S' : '';
};

// ---------- IV-8 ----------
// Kategori: Sesuai (S), Mendukung (M), Tidak Sesuai (T), Tidak Dapat Dianalisis (D, kolomnya muncul bila ada di data). x = Tidak Sesuai, u = Tidak Dapat Dianalisis, t = jumlah.
const ksMk = (s, m, x, u = 0) => { const t = s + m + x + u; return { s, m, x, u, t, ps: pc(s, t), pm: pc(m, t), px: pc(x, t), pu: pc(u, t) }; };
function buildKes() {
  const K = ksCol(); if (!ROWS.length || !K || !COLS.includes(POLA_COL)) return null;
  const g = {}; let any = false, tda = false;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), c = ksCode(r[K]); if (!p || !c) continue;
    const i = c === 'S' ? 0 : c === 'M' ? 1 : c === 'T' ? 2 : 3;
    if (c === 'D') tda = true;
    any = true; const k = kecOf(r); g[k] ??= {}; g[k][p] ??= [0, 0, 0, 0]; g[k][p][i] += area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  const kecs = Object.keys(g).sort().map(k => {
    const rows = Object.keys(g[k]).sort(by).map(p => ({ p, ...ksMk(...g[k][p]) }));
    return { k, rows, sub: ksMk(sumArr(rows.map(r => r.s)), sumArr(rows.map(r => r.m)), sumArr(rows.map(r => r.x)), sumArr(rows.map(r => r.u))) };
  });
  const tot = ksMk(sumArr(kecs.map(c => c.sub.s)), sumArr(kecs.map(c => c.sub.m)), sumArr(kecs.map(c => c.sub.x)), sumArr(kecs.map(c => c.sub.u)));
  return { kecs, tot, wil: wilayah(), tda: TDA_ALWAYS || tda };
}
const ksCells = (r, tda) => `${td(fd(r.s), 'n')}${td(fd(r.ps), 'n')}${td(fd(r.m), 'n')}${td(fd(r.pm), 'n')}${td(fd(r.x), 'n')}${td(fd(r.px), 'n')}${tda ? td(fd(r.u), 'n') + td(fd(r.pu), 'n') : ''}${tdr(r.t)}`;
const ksHead = (kolom, tda) => `<tr>${th('No', 'rowspan=3')}${th('Kabupaten/Kota', 'rowspan=3')}${kolom ? th(kolom, 'rowspan=3') : ''}${th('Kesesuaian Penggunaan Tanah terhadap RDTR/RTRW', `colspan=${tda ? 8 : 6}`)}${th('Jumlah (Ha)', 'rowspan=3')}</tr>
  <tr>${th('Sesuai', 'colspan=2')}${th('Mendukung', 'colspan=2')}${th('Tidak Sesuai', 'colspan=2')}${tda ? th(TDA_LABEL, 'colspan=2') : ''}</tr>
  <tr>${(tda ? [1, 2, 3, 4] : [1, 2, 3]).map(() => th('Luas (Ha)') + th('Luas (%)')).join('')}</tr>`;
const htmlKes = (d, ex) => {
  const top = 'style="vertical-align:top"';
  // ringkasan per kabupaten/kota + grafik (di atas), lalu tabel rinci
  let h = `<div class=wrap><div><table>${ksHead('', d.tda)}`;
  d.kecs.forEach((kc, i) => h += `<tr>${td(i + 1, 'c')}${td(kc.k)}${ksCells(kc.sub, d.tda)}</tr>`);
  h += `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${ksCells(d.tot, d.tda)}</tr></table></div>${cvTag(ex)}</div><br>`;
  h += `<div class=raw-scroll><table>${ksHead('Pola Ruang RDTR/RTRW', d.tda)}`;
  d.kecs.forEach((kc, ki) => {
    const span = kc.rows.length + 1;
    kc.rows.forEach((r, i) => {
      h += '<tr>' + (i === 0 ? `<td class=c rowspan=${span} ${top}>${ki + 1}</td><td rowspan=${span} ${top}>${kc.k}</td>` : '') + td(r.p) + ksCells(r, d.tda) + '</tr>';
    });
    h += `<tr class=tot><td class=c>Jumlah</td>${ksCells(kc.sub, d.tda)}</tr>`;
  });
  return h + `<tr class=tot><td colspan=3 class=c>${d.wil}</td>${ksCells(d.tot, d.tda)}</tr></table></div>`;
};
const chartKes = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: {
    labels: d.kecs.map(k => k.k), datasets: [
      { label: 'Sesuai', data: d.kecs.map(k => k.sub.s), backgroundColor: NAVY },
      { label: 'Mendukung', data: d.kecs.map(k => k.sub.m), backgroundColor: '#ffc000' },
      { label: 'Tidak Sesuai', data: d.kecs.map(k => k.sub.x), backgroundColor: '#c00000' }].concat(d.tda ? [{ label: TDA_LABEL, data: d.kecs.map(k => k.sub.u), backgroundColor: TDA_COLOR }] : [])
  },
  options: { ...baseOpt(['KESESUAIAN PENGGUNAAN TANAH TERHADAP FUNGSI KAWASAN', 'PER KECAMATAN DI ' + d.wil.toUpperCase() + ' TAHUN 2026 (HA)'], a, 'bottom'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excelKes(wb, ws, d, ex) {
  const tda = d.tda, KS = ['s', 'm', 'x'].concat(tda ? ['u'] : []), PK = ['ps', 'pm', 'px', 'pu'], CODE = { s: 'S', m: 'M', x: 'T', u: 'D' }, nK = KS.length;
  const heads = (c0, kolom) => {      // header 3 baris mulai kolom c0; kolom = true -> ada kolom Pola Ruang
    const o = kolom ? 3 : 2;          // jumlah kolom sebelum data angka
    H(ws, 1, c0, 3, c0, 'No'); H(ws, 1, c0 + 1, 3, c0 + 1, 'Kabupaten/Kota'); if (kolom) H(ws, 1, c0 + 2, 3, c0 + 2, 'Pola Ruang RDTR/RTRW');
    H(ws, 1, c0 + o, 1, c0 + o + nK * 2 - 1, 'Kesesuaian Penggunaan Tanah terhadap RDTR/RTRW');
    ['Sesuai', 'Mendukung', 'Tidak Sesuai'].concat(tda ? [TDA_LABEL] : []).forEach((t, i) => H(ws, 2, c0 + o + i * 2, 2, c0 + o + i * 2 + 1, t));
    for (let i = 0; i < nK * 2; i++) H(ws, 3, c0 + o + i, 3, c0 + o + i, i % 2 ? 'Luas (%)' : 'Luas (Ha)');
    H(ws, 1, c0 + o + nK * 2, 3, c0 + o + nK * 2, 'Jumlah (Ha)');
    return c0 + o;                    // kolom angka pertama
  };
  const C = colLetter, fK = rawRng('WADMKC'), fP = rawRng(POLA_COL);
  // tulis sel: tiap kategori (luas, %), lalu jumlah. f = {s,m,x,u} rumus luas (null = nilai statis); persen & jumlah selalu rumus
  const put = (r, c, o, f, b) => {
    const T = `${C(c + nK * 2)}${r}`;
    KS.forEach((k, i) => {
      cell(ws, r, c + i * 2, f[k] ? FX(f[k], o[k]) : o[k], { num: 1, b });
      cell(ws, r, c + i * 2 + 1, FX(`IF(${T}=0,0,${C(c + i * 2)}${r}/${T}*100)`, o[PK[i]]), { num: 1, b });
    });
    cell(ws, r, c + nK * 2, FX(KS.map((_, i) => `${C(c + i * 2)}${r}`).join('+'), o.t), { num: 1, b: 1 });
  };
  const sumF = (c, a, z) => { const f = {}; KS.forEach((k, i) => f[k] = `SUM(${C(c + i * 2)}${a}:${C(c + i * 2)}${z})`); return f; };
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true }, left = { vertical: 'top', wrapText: true };
  // ---- tabel rinci (kiri): kolom 1.. ----
  const n1 = heads(1, true), subRow = [];
  let r = 4;
  d.kecs.forEach((kc, ki) => {
    const k0 = r;
    kc.rows.forEach(x => {
      cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, x.p).alignment = left;
      const ex2 = fK && fP && kc.k !== '-' ? [[fK, `$B$${k0}`], [fP, `$C${r}`]] : null;
      const f = {}; if (ex2) KS.forEach(k => f[k] = ksSumF(CODE[k], ex2));
      put(r, n1, x, f); r++;
    });
    cell(ws, r, 1, null); cell(ws, r, 2, null); cell(ws, r, 3, 'Jumlah', { b: 1, ctr: 1 }); put(r, n1, kc.sub, sumF(n1, k0, r - 1), true); subRow.push(r); r++;
    ws.getCell(k0, 1).value = ki + 1; ws.getCell(k0, 1).alignment = mid;
    ws.getCell(k0, 2).value = kc.k; ws.getCell(k0, 2).alignment = left;
    ws.mergeCells(k0, 1, r - 1, 1); ws.mergeCells(k0, 2, r - 1, 2);
  });
  cell(ws, r, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, r, 2, null); cell(ws, r, 3, null); ws.mergeCells(r, 1, r, 3);
  const tf = {}; KS.forEach((k, i) => tf[k] = `SUMIF($C$4:$C$${r - 1},"Jumlah",${C(n1 + i * 2)}$4:${C(n1 + i * 2)}$${r - 1})`);   // total = jumlahkan baris "Jumlah" tiap kabupaten/kota
  put(r, n1, d.tot, tf, true);
  const cTot = n1 + nK * 2;           // kolom Jumlah tabel rinci
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 20; ws.getColumn(3).width = 38; for (let c = n1; c < cTot; c++) ws.getColumn(c).width = 12; ws.getColumn(cTot).width = 15;
  // ---- ringkasan per kabupaten/kota (kanan): mulai 2 kolom setelah tabel rinci, mengambil dari baris "Jumlah" tabel rinci ----
  const c0 = cTot + 2, n2 = heads(c0, false);
  d.kecs.forEach((kc, i) => {
    const n = 4 + i, sr = subRow[i]; cell(ws, n, c0, i + 1, { ctr: 1 }); cell(ws, n, c0 + 1, kc.k);
    const f = {}; KS.forEach((k, j) => f[k] = `${C(n1 + j * 2)}${sr}`);
    put(n, n2, kc.sub, f);
  });
  const n = 4 + d.kecs.length;
  cell(ws, n, c0, d.wil, { b: 1, ctr: 1 }); cell(ws, n, c0 + 1, null); ws.mergeCells(n, c0, n, c0 + 1); put(n, n2, d.tot, sumF(n2, 4, n - 1), true);
  ws.getColumn(c0).width = 6; ws.getColumn(c0 + 1).width = 20; for (let c = c0 + 2; c <= c0 + 1 + nK * 2; c++) ws.getColumn(c).width = 12; ws.getColumn(c0 + 2 + nK * 2).width = 15;
  const q = `'${ws.name.replace(/'/g, "''")}'!`, cat = `${q}$${C(c0 + 1)}$4:$${C(c0 + 1)}$${n - 1}`, cats = d.kecs.map(k => k.k), R = c => `${q}$${C(c)}$4:$${C(c)}$${n - 1}`;
  const series = [{ name: 'Sesuai', cat, catVals: cats, val: R(n2), vals: d.kecs.map(k => k.sub.s), color: NAVY },
                  { name: 'Mendukung', cat, catVals: cats, val: R(n2 + 2), vals: d.kecs.map(k => k.sub.m), color: '#ffc000' },
                  { name: 'Tidak Sesuai', cat, catVals: cats, val: R(n2 + 4), vals: d.kecs.map(k => k.sub.x), color: '#c00000' }];
  if (tda) series.push({ name: TDA_LABEL, cat, catVals: cats, val: R(n2 + 6), vals: d.kecs.map(k => k.sub.u), color: TDA_COLOR });
  addXChart({ sheet: ws.name, type: 'bar', legend: 'b', from: { col: c0 - 1, row: n + 1 }, ext: { w: ex.w, h: ex.h },
    title: 'KESESUAIAN PENGGUNAAN TANAH TERHADAP FUNGSI KAWASAN PER KECAMATAN DI ' + d.wil.toUpperCase() + ' TAHUN 2026 (HA)', series });
  ws.views = [{ state: 'frozen', ySplit: 3 }];
}

// ---------- IV-7 : matriks S / T / M ----------
function buildMatKs() {
  const K = ksCol(); if (!ROWS.length || !K || !COLS.includes('QNAME25') || !COLS.includes(POLA_COL)) return null;
  const m = {}, cols = new Set(); let any = false, tda = false;
  for (const r of ROWS) {
    const q = (r.QNAME25 || '').trim(), p = (r[POLA_COL] || '').trim(), c = ksCode(r[K]); if (!q || !p || !c) continue;
    any = true; if (c === 'D') tda = true; cols.add(p); (m[q] ??= {})[p] ??= new Set(); m[q][p].add(c);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  return { cols: [...cols].sort(by), rows: Object.keys(m).sort(by).map(q => ({ q, v: m[q] })), tda };
}
const matCell = (v, p) => v[p] ? [...v[p]].sort((a, b) => 'STMD'.indexOf(a) - 'STMD'.indexOf(b)).join('/') : '-';   // jika satu pasangan punya >1 nilai, tampil mis. "S/T"
const ksLegend = d => 'S = Sesuai, T = Tidak Sesuai, M = Mendukung' + (d.tda ? ', D = ' + TDA_LABEL : '') + ', - = tidak ada data';
const htmlMatKs = d => {
  const vert = 'style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:normal;height:170px;min-width:30px;padding:6px 2px;text-align:left"';
  let h = `<div class=raw-scroll><table><tr>${th('No', 'rowspan=2')}${th('Penggunaan Tanah', 'rowspan=2')}${th('Arahan Fungsi Kawasan pada Rencana Detail Tata Ruang', `colspan=${d.cols.length}`)}</tr>
  <tr>${d.cols.map(p => th(p, vert)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.q)}${d.cols.map(p => td(matCell(r.v, p), 'c')).join('')}</tr>`);
  return h + `</table></div><div class=note>${ksLegend(d)}</div>`;
};
function excelMatKs(wb, ws, d) {
  const nc = d.cols.length, C = colLetter;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Penggunaan Tanah'); H(ws, 1, 3, 1, 2 + nc, 'Arahan Fungsi Kawasan pada Rencana Detail Tata Ruang');
  d.cols.forEach((p, j) => { H(ws, 2, 3 + j, 2, 3 + j, p); ws.getCell(2, 3 + j).alignment = { textRotation: 90, horizontal: 'center', vertical: 'middle', wrapText: true }; });
  ws.getRow(2).height = 150;
  // Blok hitung (kanan, kolom abu-abu): jumlah baris data S / T / M per sel. Matriks utama membaca blok ini.
  const fQ = rawRng('QNAME25'), fP = rawRng(POLA_COL), canF = !!(fQ && fP && ksCol());
  const base = 3 + nc + 1, blk = i => base + i * (nc + 1);
  const GREY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } };
  const cnt0 = {}; { const K = ksCol(); if (K) for (const r of ROWS) { const q = (r.QNAME25 || '').trim(), p = (r[POLA_COL] || '').trim(), c = ksCode(r[K]); if (q && p && c) { const k = q + '\u0001' + p + '\u0001' + c; cnt0[k] = (cnt0[k] || 0) + 1 } } }   // jumlah baris per sel & kode
  const CODES = ['S', 'T', 'M'].concat(d.tda ? ['D'] : []);
  if (canF) CODES.forEach((k, i) => {
    H(ws, 1, blk(i), 1, blk(i) + nc - 1, 'Blok hitung ' + k + ' (jumlah baris data, dipakai matriks di kiri)');
    d.cols.forEach((p, j) => H(ws, 2, blk(i) + j, 2, blk(i) + j, p));
  });
  d.rows.forEach((r, i) => {
    const n = 3 + i; cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, r.q);
    d.cols.forEach((p, j) => {
      const cached = matCell(r.v, p);
      if (!canF) { cell(ws, n, 3 + j, cached, { ctr: 1 }); return }
      const ex2 = [[fQ, `$B${n}`], [fP, `${C(3 + j)}$2`]], cnt = [];
      CODES.forEach((k, bi) => { const x = cell(ws, n, blk(bi) + j, FX(ksCntF(k, ex2), cnt0[r.q + '\u0001' + p + '\u0001' + k] || 0)); x.fill = GREY; cnt.push(`${C(blk(bi) + j)}${n}`); });
      const order = CODES.filter(k => r.v[p] && r.v[p].has(k)).join('/') || '-';
      cell(ws, n, 3 + j, FX(`IF(${cnt.join('+')}=0,"-",SUBSTITUTE(TRIM(${CODES.map((k, q) => `IF(${cnt[q]}>0,"${k} ","")`).join('&')})," ","/"))`, order), { ctr: 1 });
    });
  });
  const n = 4 + d.rows.length; ws.getCell(n, 2).value = ksLegend(d);
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 42; for (let c = 3; c <= 2 + nc; c++) ws.getColumn(c).width = 6;
  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
}

// =====================================================================
// SHEET 16_Tabel_IV-4 - Perubahan Penggunaan Tanah (GQNAME) per Arahan Fungsi Kawasan dalam RDTR/RTRW (NAMOBJ)
//   Luas (%) = luas baris / luas total arahan fungsi kawasan itu x 100
// SHEET 14_Tabel_IV-3 - Matriks Penggunaan Tanah Lama (GNAME25) x Penggunaan Tanah Baru (QNAME25), satuan Ha
//   Butuh: POLA_COL, fmtA, nfA (dari blok IV-12) dan fd, td, th, tdr, cell, H, hdr.
// =====================================================================
// ---------- IV-4 ----------
function buildGQ() {
  if (!ROWS.length || !['GQNAME', POLA_COL].every(c => COLS.includes(c))) return null;
  const m = {}; let any = false;
  for (const r of ROWS) {
    const p = (r[POLA_COL] || '').trim(), g = (r.GQNAME || '').trim(); if (!p || !g) continue;
    any = true; (m[p] ??= {})[g] = (m[p][g] || 0) + area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y);
  return {
    groups: Object.keys(m).sort(by).map(p => {
      const items = Object.keys(m[p]).sort(by).map(g => ({ g, v: m[p][g] }));
      return { p, items, t: sumArr(items.map(i => i.v)) };
    })
  };
}
const htmlGQ = d => {
  const top = 'style="vertical-align:top"';
  let h = `<div class=raw-scroll><table><tr>${th('No')}${th('Arahan Fungsi Kawasan dalam RDTR/RTRW')}${th('Luas (Ha)')}${th('Perubahan Penggunaan Tanah')}${th('Luas (Ha)')}${th('Luas (%)')}</tr>`;
  d.groups.forEach((gr, gi) => gr.items.forEach((it, i) => {
    h += '<tr>' + (i === 0 ? `<td class=c rowspan=${gr.items.length} ${top}>${gi + 1}</td><td class=c rowspan=${gr.items.length} ${top}>${gr.p}</td><td class=n rowspan=${gr.items.length} ${top}>${fmtA(gr.t)}</td>` : '') +
      `${td(it.g)}${td(fmtA(it.v), 'n')}${td(fmtA(pc(it.v, gr.t)), 'n')}</tr>`;
  }));
  return h + '</table></div>';
};
function excelGQ(wb, ws, d) {
  ['No', 'Arahan Fungsi Kawasan dalam RDTR/RTRW', 'Luas (Ha)', 'Perubahan Penggunaan Tanah', 'Luas (Ha)', 'Luas (%)'].forEach((t, i) => H(ws, 1, i + 1, 1, i + 1, t));
  const mid = { horizontal: 'center', vertical: 'top', wrapText: true };
  const fL = rawRng('LUASHA'), fG = rawRng('GQNAME'), fP = rawRng(POLA_COL), canF = !!(fL && fG && fP);
  const num = (r, c, v, f) => { const x = cell(ws, r, c, f ? FX(f, v) : v, { num: 1 }); x.numFmt = nfA(v); };
  let r = 2;
  d.groups.forEach((gr, gi) => {
    const g0 = r;
    gr.items.forEach(it => {
      for (let c = 1; c <= 3; c++) cell(ws, r, c, null); cell(ws, r, 4, it.g);
      num(r, 5, it.v, canF ? `SUMIFS(${fL},${fP},$B$${g0},${fG},$D${r})` : null);          // luas perubahan tertentu pada arahan ini
      num(r, 6, pc(it.v, gr.t), `IF($C$${g0}=0,0,E${r}/$C$${g0}*100)`); r++;
    });
    ws.getCell(g0, 1).value = gi + 1; ws.getCell(g0, 2).value = gr.p;
    ws.getCell(g0, 3).value = FX(`SUM(E${g0}:E${r - 1})`, gr.t); ws.getCell(g0, 3).numFmt = nfA(gr.t);                // luas total arahan
    for (let c = 1; c <= 3; c++) ws.getCell(g0, c).alignment = mid;
    if (r - 1 > g0) for (let c = 1; c <= 3; c++) ws.mergeCells(g0, c, r - 1, c);
  });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 34; ws.getColumn(3).width = 14; ws.getColumn(4).width = 70; ws.getColumn(5).width = 14; ws.getColumn(6).width = 12;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
}

// ---------- IV-3 ----------
function buildMatLB() {
  if (!ROWS.length || !['GNAME25', 'QNAME25'].every(c => COLS.includes(c))) return null;
  const m = {}, cs = new Set(); let any = false;
  for (const r of ROWS) {
    const g = (r.GNAME25 || '').trim(), q = (r.QNAME25 || '').trim(); if (!g || !q) continue;
    any = true; cs.add(q); (m[g] ??= {})[q] = (m[g][q] || 0) + area(r);
  }
  if (!any) return null;
  const by = (x, y) => x.localeCompare(y), cols = [...cs].sort(by);
  const rows = Object.keys(m).sort(by).map(g => ({ g, v: m[g], t: sumArr(Object.values(m[g])) }));
  const kt = cols.map(q => sumArr(rows.map(r => r.v[q] || 0)));
  return { cols, rows, kt, T: sumArr(kt) };
}
const htmlMatLB = d => {
  const vert = 'style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:normal;height:170px;min-width:30px;padding:6px 2px;text-align:left"';
  let h = `<div style="overflow-x:auto"><table><tr>${th('No', 'rowspan=2')}${th('Penggunaan Tanah Lama (Tahun 2014)', 'rowspan=2')}${th('Penggunaan Tanah Baru (Tahun 2026)', `colspan=${d.cols.length}`)}${th('Jumlah (Ha)', 'rowspan=2')}</tr>
  <tr>${d.cols.map(q => th(q, vert)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.g)}${d.cols.map(q => td(fd(r.v[q] || 0), 'n')).join('')}${tdr(r.t)}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Ha)</td>${d.kt.map(v => tdr(v)).join('')}${tdr(d.T)}</tr></table></div>`;
};
function excelMatLB(wb, ws, d) {
  const nc = d.cols.length;
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Penggunaan Tanah Lama (Tahun 2014)'); H(ws, 1, 3, 1, 2 + nc, 'Penggunaan Tanah Baru (Tahun 2026)'); H(ws, 1, 3 + nc, 2, 3 + nc, 'Jumlah (Ha)');
  d.cols.forEach((q, j) => { H(ws, 2, 3 + j, 2, 3 + j, q); ws.getCell(2, 3 + j).alignment = { textRotation: 90, horizontal: 'center', vertical: 'middle', wrapText: true }; });
  ws.getRow(2).height = 150;
  const fL = rawRng('LUASHA'), fG = rawRng('GNAME25'), fQ = rawRng('QNAME25'), canF = !!(fL && fG && fQ);
  const r1 = 3, rL = 2 + d.rows.length, n = rL + 1, C = colLetter, cT = 3 + nc;
  d.rows.forEach((r, i) => {
    const x = r1 + i; cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.g);
    d.cols.forEach((q, j) => cell(ws, x, 3 + j, canF ? FX(`SUMIFS(${fL},${fG},$B${x},${fQ},${C(3 + j)}$2)`, r.v[q] || 0) : (r.v[q] || 0), { num: 1 }));   // baris = lama, kolom = baru
    cell(ws, x, cT, FX(`SUM(${C(3)}${x}:${C(2 + nc)}${x})`, r.t), { num: 1, b: 1 });
  });
  cell(ws, n, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, n, 2, null); ws.mergeCells(n, 1, n, 2);
  d.kt.forEach((v, j) => cell(ws, n, 3 + j, FX(`SUM(${C(3 + j)}${r1}:${C(3 + j)}${rL})`, v), { num: 1, b: 1 }));
  cell(ws, n, cT, FX(`SUM(${C(cT)}${r1}:${C(cT)}${rL})`, d.T), { num: 1, b: 1 });
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 36; for (let c = 3; c <= 2 + nc; c++) ws.getColumn(c).width = 10; ws.getColumn(3 + nc).width = 13;
  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2 }];
}

// =====================================================================
// SHEET 01_Tabel_III-1 - Luas wilayah administrasi per kabupaten/kota + Jumlah Penduduk (diisi manual)
//   Luas (Ha) dan % Wilayah dihitung dari data (LUASHA per WADMKC).
//   Kolom Jiwa DIISI MANUAL (kotak isian di tabel). Setelah diisi, % Total Penduduk dan
//   Kepadatan (Jiwa/km2 = Jiwa / (Ha/100)) langsung terhitung. Isian disimpan di browser (localStorage).
//   Di Excel: sel Jiwa berwarna kuning, % dan Kepadatan memakai rumus sehingga ikut terhitung saat diisi.
// =====================================================================
const POP = (() => { try { return JSON.parse(localStorage.getItem('djpa_pop') || '{}') } catch (e) { return {} } })();
const popKey = k => wilayah() + '|' + k;
const popGet = k => { const v = POP[popKey(k)]; return typeof v === 'number' && isFinite(v) ? v : null };
const popSet = (k, v) => {
  if (v === null || isNaN(v)) delete POP[popKey(k)]; else POP[popKey(k)] = v;
  try { localStorage.setItem('djpa_pop', JSON.stringify(POP)) } catch (e) { }
};
const fInt = v => Math.round(v).toLocaleString('en-US');
let T3 = null;   // data tabel yang sedang tampil (dipakai saat mengetik)
function popCalc(rows, totHa) {
  const vals = rows.map(r => popGet(r.k)), filled = vals.filter(v => v !== null), tj = sumArr(filled);
  return {
    vals, tj, n: filled.length,
    pp: vals.map(v => v !== null && tj ? pc(v, tj) : null),
    kp: vals.map((v, i) => v !== null && rows[i].ha ? v / (rows[i].ha / 100) : null),
    tkp: filled.length && totHa ? tj / (totHa / 100) : null
  };
}
function buildT3() {
  if (!ROWS.length || !COLS.includes('WADMKC')) return null;
  const m = {}; let any = false;
  for (const r of ROWS) { const k = kecOf(r); any = true; m[k] = (m[k] || 0) + area(r); }
  if (!any) return null;
  const totHa = sumArr(Object.values(m));
  const rows = Object.keys(m).sort().map(k => ({ k, ha: m[k], pw: pc(m[k], totHa) }));
  return (T3 = { rows, totHa, wil: wilayah() });
}
const esc3 = s => String(s).replace(/"/g, '&quot;');
const htmlT3 = (d, ex) => {
  const p = popCalc(d.rows, d.totHa), dash = v => v === null ? '-' : v;
  const inp = 'style="width:120px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  let h = `<div class=wrap><div><table><tr>${th('No', 'rowspan=2')}${th('Kabupaten/Kota', 'rowspan=2')}${th('Luas', 'colspan=2')}${th('Jumlah Penduduk', 'colspan=2')}${th('Kepadatan Penduduk (Jiwa/km²)', 'rowspan=2')}</tr>
  <tr>${th('Ha')}${th('% Wilayah')}${th('Jiwa')}${th('% Total Penduduk')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${td(fd(r.ha), 'n')}${td(fd(r.pw), 'n')}<td class=n><input class=pop-in type=number min=0 step=any placeholder="isi jiwa" data-k="${esc3(r.k)}" value="${p.vals[i] ?? ''}" ${inp}></td>${td(dash(p.pp[i] === null ? null : fd(p.pp[i])), 'n pp')}${td(dash(p.kp[i] === null ? null : fInt(p.kp[i])), 'n kp')}</tr>`);
  h += `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${td(fd(d.totHa), 'n')}${td('100.00', 'n')}${td(p.n ? fInt(p.tj) : '-', 'n tj')}${td(p.n && p.tj ? '100.00' : '-', 'n tpp')}${td(p.tkp === null ? '-' : fInt(p.tkp), 'n tkp')}</tr></table>
  <div class=note>Isi kolom <b>Jiwa</b> secara manual (angka tanpa titik/koma). % Total Penduduk dan Kepadatan terhitung otomatis; isian tersimpan di browser ini.</div></div>${cvTag(ex)}</div>`;
  return h;
};
function popRefresh(root) {
  if (!T3 || !root) return;
  root.querySelectorAll('.pop-in').forEach(el => { const v = el.value === '' ? null : parseFloat(el.value); popSet(el.dataset.k, v); });
  const p = popCalc(T3.rows, T3.totHa), rows = [...root.querySelectorAll('.pop-in')].map(e => e.closest('tr'));
  rows.forEach((tr, i) => {
    tr.querySelector('.pp').textContent = p.pp[i] === null ? '-' : fd(p.pp[i]);
    tr.querySelector('.kp').textContent = p.kp[i] === null ? '-' : fInt(p.kp[i]);
  });
  root.querySelector('.tj').textContent = p.n ? fInt(p.tj) : '-';
  root.querySelector('.tpp').textContent = p.n && p.tj ? '100.00' : '-';
  root.querySelector('.tkp').textContent = p.tkp === null ? '-' : fInt(p.tkp);
}
document.addEventListener('input', e => {
  if (e.target.matches && e.target.matches('.pop-in')) popRefresh(e.target.closest('.panel'));
});
const chartT3 = (cv, d, a) => new Chart(cv, {
  type: 'pie',
  data: { labels: d.rows.map(r => r.k), datasets: [{ data: d.rows.map(r => r.ha), backgroundColor: PAL }] },
  options: baseOpt('LUAS WILAYAH ADMINISTRASI (%)', a, 'right'),
  plugins: [bgPlugin, valPlugin]
});
function excelT3(wb, ws, d, ex) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Kabupaten/Kota'); H(ws, 1, 3, 1, 4, 'Luas'); H(ws, 1, 5, 1, 6, 'Jumlah Penduduk'); H(ws, 1, 7, 2, 7, 'Kepadatan Penduduk (Jiwa/km2)');
  ['Ha', '% Wilayah', 'Jiwa', '% Total Penduduk'].forEach((t, i) => H(ws, 2, 3 + i, 2, 3 + i, t));
  const n = d.rows.length, r0 = 3, r1 = 2 + n, rt = 3 + n, p = popCalc(d.rows, d.totHa);
  const fx = (r, c, formula, result, f, b) => { const x = cell(ws, r, c, null, { b }); x.value = { formula, result }; x.numFmt = f; return x; };
  d.rows.forEach((row, i) => {
    const r = r0 + i;
    cell(ws, r, 1, i + 1, { ctr: 1 }); cell(ws, r, 2, row.k); cell(ws, r, 3, rawRng('LUASHA') && rawRng('WADMKC') && row.k !== '-' ? FX(`SUMIFS(${rawRng('LUASHA')},${rawRng('WADMKC')},$B${r})`, row.ha) : row.ha, { num: 1 });
    fx(r, 4, `C${r}/C$${rt}*100`, row.pw, '#,##0.00');
    const e = cell(ws, r, 5, p.vals[i]); e.numFmt = '#,##0'; e.fill = YEL;
    fx(r, 6, `IF(OR(E${r}="",E$${rt}=0),"",E${r}/E$${rt}*100)`, p.pp[i] === null ? '' : p.pp[i], '#,##0.00');
    fx(r, 7, `IF(OR(E${r}="",C${r}=0),"",E${r}/(C${r}/100))`, p.kp[i] === null ? '' : p.kp[i], '#,##0');
  });
  cell(ws, rt, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  fx(rt, 3, `SUM(C${r0}:C${r1})`, d.totHa, '#,##0.00', 1);
  fx(rt, 4, `SUM(D${r0}:D${r1})`, 100, '#,##0.00', 1);
  fx(rt, 5, `IF(COUNT(E${r0}:E${r1})=0,"",SUM(E${r0}:E${r1}))`, p.n ? p.tj : '', '#,##0', 1);
  fx(rt, 6, `IF(COUNT(E${r0}:E${r1})=0,"",100)`, p.n ? 100 : '', '#,##0.00', 1);
  fx(rt, 7, `IF(OR(E${rt}="",C${rt}=0),"",E${rt}/(C${rt}/100))`, p.tkp === null ? '' : p.tkp, '#,##0', 1);
  ws.getCell(rt + 2, 1).value = 'Sel kuning (kolom Jiwa) dapat diisi manual; % Total Penduduk dan Kepadatan Penduduk terhitung otomatis.';
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 24; ws.getColumn(3).width = 14; ws.getColumn(4).width = 14; ws.getColumn(5).width = 16; ws.getColumn(6).width = 16; ws.getColumn(7).width = 20;
  const qq = `'${ws.name.replace(/'/g, "''")}'!`;
  addXChart({ sheet: ws.name, type: 'pie', title: 'LUAS WILAYAH ADMINISTRASI (%)', legend: 'r', from: { col: 8, row: 0 }, ext: { w: ex.w, h: ex.h },
    pointColors: d.rows.map((_, i) => PAL[i % PAL.length]),
    series: [{ cat: `${qq}$B$${r0}:$B$${r1}`, catVals: d.rows.map(x => x.k), val: `${qq}$C$${r0}:$C$${r1}`, vals: d.rows.map(x => x.ha) }] });
}

// =====================================================================
// SHEET 02_Tabel_III-2 - Jumlah Penduduk, Laju Pertumbuhan Penduduk, Rasio Jenis Kelamin per Kabupaten/Kota
//   Jumlah Penduduk (Jiwa) DIAMBIL OTOMATIS dari isian Tabel III-1 (tidak perlu diisi dua kali).
//   Laju Pertumbuhan (%/Tahun) dan Rasio Jenis Kelamin diisi MANUAL (kotak isian).
//   Baris total Laju dan Rasio = rata-rata tertimbang menurut Jumlah Penduduk (sama dengan contoh).
//   Butuh blok 01_Tabel_III-1 (popKey, popGet, fInt, buildT3, ...) sudah terpasang.
// =====================================================================
const T2S = (() => { try { return JSON.parse(localStorage.getItem('djpa_pop2') || '{}') } catch (e) { return {} } })();   // {laju:{kunci:v}, rasio:{kunci:v}}
const t2Get = (f, k) => { const v = (T2S[f] || {})[popKey(k)]; return typeof v === 'number' && isFinite(v) ? v : null };
const t2Set = (f, k, v) => {
  const m = (T2S[f] ??= {}); if (v === null || isNaN(v)) delete m[popKey(k)]; else m[popKey(k)] = v;
  try { localStorage.setItem('djpa_pop2', JSON.stringify(T2S)) } catch (e) { }
};
// rata-rata tertimbang: hanya baris yang jiwa DAN nilainya terisi
const wavg = (w, v) => { let n = 0, d = 0; w.forEach((x, i) => { if (x !== null && v[i] !== null) { n += x * v[i]; d += x } }); return d ? n / d : null };
let T2D = null;
function t2Calc(rows) {
  const j = rows.map(r => popGet(r.k)), l = rows.map(r => t2Get('laju', r.k)), s = rows.map(r => t2Get('rasio', r.k)), fj = j.filter(v => v !== null);
  return { j, l, s, tj: fj.length ? sumArr(fj) : null, tl: wavg(j, l), ts: wavg(j, s) };
}
function buildT2() {
  if (!ROWS.length || !COLS.includes('WADMKC')) return null;
  const ks = [...new Set(ROWS.map(kecOf))].sort();
  return (T2D = { rows: ks.map(k => ({ k })), wil: wilayah() });
}
const htmlT2 = d => {
  const c = t2Calc(d.rows);
  const inp = 'style="width:110px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  const box = (f, k, v, ph) => `<input class=t2-in type=number step=any data-f=${f} data-k="${String(k).replace(/"/g, '&quot;')}" placeholder="${ph}" value="${v ?? ''}" ${inp}>`;
  let h = `<div class=t2root><table><tr>${th('No')}${th('Kabupaten/Kota')}${th('Jumlah Penduduk (Jiwa)')}${th('Laju Pertumbuhan Penduduk (%/Tahun)')}${th('Rasio Jenis Kelamin Penduduk')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${td(c.j[i] === null ? '-' : fInt(c.j[i]), 'n t2j')}<td class=n>${box('laju', r.k, c.l[i], 'isi %')}</td><td class=n>${box('rasio', r.k, c.s[i], 'isi rasio')}</td></tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${td(c.tj === null ? '-' : fInt(c.tj), 'n t2tj')}${td(c.tl === null ? '-' : c.tl.toFixed(2), 'n t2tl')}${td(c.ts === null ? '-' : c.ts.toFixed(1), 'n t2ts')}</tr></table>
  <div class=note>Jumlah Penduduk otomatis dari Tabel III-1. Isi Laju Pertumbuhan dan Rasio Jenis Kelamin secara manual; baris total dihitung sebagai rata-rata tertimbang menurut jumlah penduduk.</div></div>`;
};
function t2Refresh() {
  const root = document.querySelector('.t2root'); if (!root || !T2D) return;
  const c = t2Calc(T2D.rows);
  root.querySelectorAll('.t2j').forEach((el, i) => el.textContent = c.j[i] === null ? '-' : fInt(c.j[i]));
  root.querySelector('.t2tj').textContent = c.tj === null ? '-' : fInt(c.tj);
  root.querySelector('.t2tl').textContent = c.tl === null ? '-' : c.tl.toFixed(2);
  root.querySelector('.t2ts').textContent = c.ts === null ? '-' : c.ts.toFixed(1);
}
document.addEventListener('input', e => {
  const t = e.target; if (!t.matches) return;
  if (t.matches('.t2-in')) { t2Set(t.dataset.f, t.dataset.k, t.value === '' ? null : parseFloat(t.value)); t2Refresh(); }
  else if (t.matches('.pop-in')) setTimeout(t2Refresh, 0);   // Jiwa diubah di Tabel III-1 -> segarkan tabel ini
});
function excelT2(wb, ws, d) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  const s1 = (EXTRA.find(e => e.build === buildT3) || { tab: '01_Tabel_III-1' }).tab;   // nama sheet Tabel III-1 (sumber Jiwa)
  ['No', 'Kabupaten/Kota', 'Jumlah Penduduk (Jiwa)', 'Laju Pertumbuhan Penduduk (%/Tahun)', 'Rasio Jenis Kelamin Penduduk'].forEach((t, i) => H(ws, 1, i + 1, 1, i + 1, t));
  const n = d.rows.length, r0 = 2, r1 = 1 + n, rt = 2 + n, c = t2Calc(d.rows);
  const fx = (r, col, formula, result, f) => { const x = cell(ws, r, col, null, { b: r === rt }); x.value = { formula, result }; x.numFmt = f; return x; };
  d.rows.forEach((row, i) => {
    const r = r0 + i, src = `'${s1}'!E${3 + i}`;                 // sel Jiwa di Tabel III-1 (baris data mulai 3)
    cell(ws, r, 1, i + 1, { ctr: 1 }); cell(ws, r, 2, row.k);
    fx(r, 3, `IF(${src}="","",${src})`, c.j[i] === null ? '' : c.j[i], '#,##0');
    const l = cell(ws, r, 4, c.l[i]); l.numFmt = '0.00'; l.fill = YEL;
    const s = cell(ws, r, 5, c.s[i]); s.numFmt = '0.0'; s.fill = YEL;
  });
  cell(ws, rt, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  const C = `C${r0}:C${r1}`, wt = col => `SUMPRODUCT(${C},--ISNUMBER(${col}${r0}:${col}${r1}))`, nm = col => `SUMPRODUCT(${C},${col}${r0}:${col}${r1})`;
  fx(rt, 3, `IF(COUNT(${C})=0,"",SUM(${C}))`, c.tj === null ? '' : c.tj, '#,##0');
  fx(rt, 4, `IF(${wt('D')}=0,"",${nm('D')}/${wt('D')})`, c.tl === null ? '' : c.tl, '0.00');
  fx(rt, 5, `IF(${wt('E')}=0,"",${nm('E')}/${wt('E')})`, c.ts === null ? '' : c.ts, '0.0');
  ws.getCell(rt + 2, 1).value = 'Jumlah Penduduk mengikuti sheet Tabel III-1. Sel kuning (Laju dan Rasio) diisi manual; baris total = rata-rata tertimbang menurut jumlah penduduk.';
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 24; ws.getColumn(3).width = 22; ws.getColumn(4).width = 26; ws.getColumn(5).width = 26;
  ws.getRow(1).height = 34;
}

// =====================================================================
// SHEET 03_Tabel_III-3 - Status Pekerjaan Utama Penduduk per Jenis Kelamin (diisi MANUAL)
//   Isi kolom Laki-Laki dan Perempuan; Jumlah per baris, baris total, dan grafik terhitung otomatis.
//   Isian disimpan di browser (localStorage). Di Excel: sel Laki-Laki/Perempuan berwarna kuning,
//   Jumlah memakai rumus (grafik di Excel berupa gambar sesuai isian saat file diunduh).
// =====================================================================
const PK_TAHUN = 2025;   // tahun data untuk judul grafik
const PK_STATUS = ['Berusaha sendiri', 'Buruh tidak dibayar', 'Buruh dibayar', 'Buruh/Karyawan/Pegawai', 'Pekerja bebas', 'Pekerja keluarga/tak dibayar'];
const PKS = (() => { try { return JSON.parse(localStorage.getItem('djpa_pek') || '{}') } catch (e) { return {} } })();
const pkKey = (i, j) => wilayah() + '|' + PK_STATUS[i] + '|' + j;
const pkGet = (i, j) => { const v = PKS[pkKey(i, j)]; return typeof v === 'number' && isFinite(v) ? v : null };
const pkSet = (i, j, v) => {
  if (v === null || isNaN(v)) delete PKS[pkKey(i, j)]; else PKS[pkKey(i, j)] = v;
  try { localStorage.setItem('djpa_pek', JSON.stringify(PKS)) } catch (e) { }
};
const pkF = v => Math.round(v).toLocaleString('en-US');
function pkCalc() {
  const l = PK_STATUS.map((_, i) => pkGet(i, 'l')), p = PK_STATUS.map((_, i) => pkGet(i, 'p'));
  const j = l.map((x, i) => x === null && p[i] === null ? null : (x || 0) + (p[i] || 0));
  const s = a => { const f = a.filter(v => v !== null); return f.length ? sumArr(f) : null };
  return { l, p, j, tl: s(l), tp: s(p), tj: s(j) };
}
let PKC = null;   // grafik yang sedang tampil (diperbarui saat mengetik)
function buildPk() { return ROWS.length ? { wil: wilayah() } : null }
const htmlPk = (d, ex) => {
  const c = pkCalc(), dash = v => v === null ? '-' : pkF(v);
  const inp = 'style="width:110px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  const box = (i, j, v) => `<input class=pk-in type=number min=0 step=any data-i=${i} data-j=${j} placeholder="isi" value="${v ?? ''}" ${inp}>`;
  let h = `<div class=wrap><div class=pkroot><table><tr>${th('No', 'rowspan=2')}${th('Status Pekerjaan Utama', 'rowspan=2')}${th('Jumlah Penduduk per-Jenis Kelamin (Jiwa)', 'colspan=2')}${th('Jumlah Penduduk (Jiwa)', 'rowspan=2')}</tr>
  <tr>${th('Laki-Laki')}${th('Perempuan')}</tr>`;
  PK_STATUS.forEach((s, i) => h += `<tr>${td(i + 1, 'c')}${td(s)}<td class=n>${box(i, 'l', c.l[i])}</td><td class=n>${box(i, 'p', c.p[i])}</td>${td(dash(c.j[i]), 'n pk-j')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Jumlah (Jiwa)</td>${td(dash(c.tl), 'n pk-tl')}${td(dash(c.tp), 'n pk-tp')}${td(dash(c.tj), 'n pk-tj')}</tr></table>
  <div class=note>Isi kolom Laki-Laki dan Perempuan secara manual (angka tanpa titik/koma). Jumlah, total, dan grafik terhitung otomatis; isian tersimpan di browser ini.</div></div>${cvTag(ex)}</div>`;
};
function pkRefresh() {
  const root = document.querySelector('.pkroot'); if (!root) return;
  const c = pkCalc(), dash = v => v === null ? '-' : pkF(v);
  root.querySelectorAll('.pk-j').forEach((el, i) => el.textContent = dash(c.j[i]));
  root.querySelector('.pk-tl').textContent = dash(c.tl); root.querySelector('.pk-tp').textContent = dash(c.tp); root.querySelector('.pk-tj').textContent = dash(c.tj);
  if (PKC && PKC.canvas && document.body.contains(PKC.canvas)) {
    PKC.data.datasets[0].data = c.l.map(v => v || 0); PKC.data.datasets[1].data = c.p.map(v => v || 0); PKC.update('none');
  }
}
document.addEventListener('input', e => {
  const t = e.target; if (!t.matches || !t.matches('.pk-in')) return;
  pkSet(+t.dataset.i, t.dataset.j, t.value === '' ? null : parseFloat(t.value)); pkRefresh();
});
// label angka di ujung batang horizontal (angka 0 tidak ditampilkan)
const hbarVals = {
  id: 'hbv', afterDatasetsDraw(ch) {
    const c = ch.ctx; c.save(); c.font = '10px Arial'; c.fillStyle = '#222'; c.textAlign = 'left'; c.textBaseline = 'middle';
    ch.data.datasets.forEach((ds, i) => ch.getDatasetMeta(i).data.forEach((el, j) => { const v = ds.data[j]; if (v > 0) c.fillText(pkF(v), el.x + 4, el.y); }));
    c.restore();
  }
};
const chartPk = (cv, d, a) => {
  const c = pkCalc();
  const ch = new Chart(cv, {
    type: 'bar',
    data: {
      labels: PK_STATUS, datasets: [
        { label: 'Laki-Laki', data: c.l.map(v => v || 0), backgroundColor: NAVY },
        { label: 'Perempuan', data: c.p.map(v => v || 0), backgroundColor: '#ffc000' }]
    },
    options: { ...baseOpt(['STATUS PEKERJAAN UTAMA PENDUDUK', d.wil.toUpperCase() + ' TAHUN ' + PK_TAHUN + ' (JIWA)'], a, 'bottom'), indexAxis: 'y', layout: { padding: { top: 18, right: 56 } }, scales: { x: { beginAtZero: true } } },
    plugins: [bgPlugin, hbarVals]
  });
  if (a !== false) PKC = ch;   // addChart (Excel) memanggil dengan a=false -> bukan grafik tampilan
  return ch;
};
function excelPk(wb, ws, d, ex) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Status Pekerjaan Utama'); H(ws, 1, 3, 1, 4, 'Jumlah Penduduk per-Jenis Kelamin (Jiwa)'); H(ws, 1, 5, 2, 5, 'Jumlah Penduduk (Jiwa)');
  H(ws, 2, 3, 2, 3, 'Laki-Laki'); H(ws, 2, 4, 2, 4, 'Perempuan');
  const n = PK_STATUS.length, r0 = 3, r1 = 2 + n, rt = 3 + n, c = pkCalc();
  const fx = (r, col, formula, result) => { const x = cell(ws, r, col, null, { b: r === rt || col === 5 }); x.value = { formula, result }; x.numFmt = '#,##0'; return x; };
  PK_STATUS.forEach((s, i) => {
    const r = r0 + i; cell(ws, r, 1, i + 1, { ctr: 1 }); cell(ws, r, 2, s);
    const l = cell(ws, r, 3, c.l[i]); l.numFmt = '#,##0'; l.fill = YEL;
    const p = cell(ws, r, 4, c.p[i]); p.numFmt = '#,##0'; p.fill = YEL;
    fx(r, 5, `IF(COUNT(C${r}:D${r})=0,"",SUM(C${r}:D${r}))`, c.j[i] === null ? '' : c.j[i]);
  });
  cell(ws, rt, 1, 'Jumlah (Jiwa)', { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  [['C', 'tl'], ['D', 'tp'], ['E', 'tj']].forEach(([L, k], i) => fx(rt, 3 + i, `IF(COUNT(${L}${r0}:${L}${r1})=0,"",SUM(${L}${r0}:${L}${r1}))`, c[k] === null ? '' : c[k]));
  ws.getCell(rt + 2, 1).value = 'Sel kuning (Laki-Laki dan Perempuan) diisi manual; Jumlah terhitung otomatis. Grafik ikut berubah saat sel kuning diisi.';
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 34; ws.getColumn(3).width = 16; ws.getColumn(4).width = 16; ws.getColumn(5).width = 20;
  const qq = `'${ws.name.replace(/'/g, "''")}'!`, cat = `${qq}$B$${r0}:$B$${r1}`;
  addXChart({ sheet: ws.name, type: 'bar', dir: 'bar', valFmt: '#,##0', legend: 'b', from: { col: 6, row: 0 }, ext: { w: ex.w, h: ex.h },
    title: 'STATUS PEKERJAAN UTAMA PENDUDUK ' + d.wil.toUpperCase() + ' TAHUN ' + PK_TAHUN + ' (JIWA)',
    series: [{ name: 'Laki-Laki', cat, catVals: PK_STATUS, val: `${qq}$C$${r0}:$C$${r1}`, vals: c.l.map(v => v || 0), color: NAVY },
             { name: 'Perempuan', cat, catVals: PK_STATUS, val: `${qq}$D$${r0}:$D$${r1}`, vals: c.p.map(v => v || 0), color: '#ffc000' }] });
}

// =====================================================================
// SHEET 04_Tabel_III-4 - PDRB Atas Dasar Harga Berlaku menurut Lapangan Usaha (diisi MANUAL)
//   Isi nilai tiap lapangan usaha per tahun; baris total (Produk Domestik Regional Bruto) terhitung otomatis.
//   Isian disimpan di browser (localStorage). Di Excel: sel nilai berwarna kuning, total memakai rumus.
//   Ganti tahun lewat array PD_TAHUN di bawah ini.
// =====================================================================
const PD_TAHUN = [2021, 2022, 2023, 2024, 2025];
const PD_USAHA = [
  ['A', 'Pertanian, Kehutanan, dan Perikanan'], ['B', 'Pertambangan dan Penggalian'], ['C', 'Industri Pengolahan'],
  ['D', 'Pengadaan Listrik dan Gas'], ['E', 'Pengadaan Air; Pengelolaan Sampah, Limbah, dan Daur Ulang'], ['F', 'Konstruksi'],
  ['G', 'Perdagangan Besar dan Eceran; Reparasi Mobil dan Sepeda Motor'], ['H', 'Transportasi dan Pergudangan'],
  ['I', 'Penyediaan Akomodasi dan Makan Minum'], ['J', 'Informasi dan Komunikasi'], ['K', 'Jasa Keuangan dan Asuransi'],
  ['L', 'Real Estat'], ['M, N', 'Jasa Perusahaan'], ['O', 'Administrasi Pemerintahan, Pertahanan, dan Jaminan Sosial Wajib'],
  ['P', 'Jasa Pendidikan'], ['Q', 'Jasa Kesehatan dan Kegiatan Sosial'], ['R, S, T, U', 'Jasa Lainnya']
];
const PDS = (() => { try { return JSON.parse(localStorage.getItem('djpa_pdrb') || '{}') } catch (e) { return {} } })();
const pdKey = (i, y) => wilayah() + '|' + PD_USAHA[i][0] + '|' + PD_TAHUN[y];
const pdGet = (i, y) => { const v = PDS[pdKey(i, y)]; return typeof v === 'number' && isFinite(v) ? v : null };
const pdSet = (i, y, v) => {
  if (v === null || isNaN(v)) delete PDS[pdKey(i, y)]; else PDS[pdKey(i, y)] = v;
  try { localStorage.setItem('djpa_pdrb', JSON.stringify(PDS)) } catch (e) { }
};
function pdCalc() {
  const v = PD_USAHA.map((_, i) => PD_TAHUN.map((_, y) => pdGet(i, y)));
  const t = PD_TAHUN.map((_, y) => { const f = v.map(r => r[y]).filter(x => x !== null); return f.length ? sumArr(f) : null });
  return { v, t };
}
const pdShow = v => v === null ? '–' : fmt(v);
function buildPd() { return ROWS.length ? { wil: wilayah() } : null }
const htmlPd = () => {
  const c = pdCalc();
  const inp = 'style="width:110px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  let h = `<div class=pdroot style="overflow-x:auto"><table><tr>${th('Kode')}${th('Lapangan Usaha')}${PD_TAHUN.map(y => th(y)).join('')}</tr>`;
  PD_USAHA.forEach(([k, n], i) => h += `<tr>${td(k, 'c')}${td(n)}${PD_TAHUN.map((_, y) => `<td class=n><input class=pd-in type=number min=0 step=any data-i=${i} data-y=${y} placeholder="isi" value="${c.v[i][y] ?? ''}" ${inp}></td>`).join('')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Produk Domestik Regional Bruto</td>${PD_TAHUN.map((_, y) => td(pdShow(c.t[y]), 'n pd-t')).join('')}</tr></table>
  <div class=note>Isi nilai tiap lapangan usaha secara manual (angka tanpa pemisah ribuan; kosongkan jika tidak ada). Total terhitung otomatis; isian tersimpan di browser ini.</div></div>`;
};
document.addEventListener('input', e => {
  const t = e.target; if (!t.matches || !t.matches('.pd-in')) return;
  pdSet(+t.dataset.i, +t.dataset.y, t.value === '' ? null : parseFloat(t.value));
  const root = t.closest('.pdroot'), c = pdCalc();
  root.querySelectorAll('.pd-t').forEach((el, y) => el.textContent = pdShow(c.t[y]));
});
function excelPd(wb, ws) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  H(ws, 1, 1, 1, 1, 'Kode'); H(ws, 1, 2, 1, 2, 'Lapangan Usaha'); PD_TAHUN.forEach((y, j) => H(ws, 1, 3 + j, 1, 3 + j, y));
  const n = PD_USAHA.length, r0 = 2, r1 = 1 + n, rt = 2 + n, c = pdCalc();
  PD_USAHA.forEach(([k, nm], i) => {
    const r = r0 + i; cell(ws, r, 1, k, { ctr: 1 }); cell(ws, r, 2, nm).alignment = { wrapText: true, vertical: 'middle' };
    PD_TAHUN.forEach((_, y) => { const x = cell(ws, r, 3 + y, c.v[i][y]); x.numFmt = '#,##0.00'; x.fill = YEL; });
  });
  cell(ws, rt, 1, 'Produk Domestik Regional Bruto', { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  PD_TAHUN.forEach((_, y) => {
    const L = String.fromCharCode(67 + y), x = cell(ws, rt, 3 + y, null, { b: 1 });
    x.value = { formula: `IF(COUNT(${L}${r0}:${L}${r1})=0,"",SUM(${L}${r0}:${L}${r1}))`, result: c.t[y] === null ? '' : c.t[y] }; x.numFmt = '#,##0.00';
  });
  ws.getCell(rt + 2, 1).value = 'Sel kuning diisi manual; baris total terhitung otomatis.';
  ws.getColumn(1).width = 10; ws.getColumn(2).width = 56; for (let k = 0; k < PD_TAHUN.length; k++) ws.getColumn(3 + k).width = 14;
}

// =====================================================================
// SHEET 05_Tabel_III-5 - sama seperti 04_Tabel_III-4 (Lapangan Usaha x Tahun), data TERPISAH, diisi MANUAL
//   Memakai daftar PD_USAHA, PD_TAHUN, dan pdShow dari blok 04_Tabel_III-4 (blok itu harus sudah terpasang).
//   Isian disimpan di browser (localStorage 'djpa_pdrb2'). Di Excel: sel nilai kuning, total memakai rumus.
// =====================================================================
const PD2S = (() => { try { return JSON.parse(localStorage.getItem('djpa_pdrb2') || '{}') } catch (e) { return {} } })();
const pd2Key = (i, y) => wilayah() + '|' + PD_USAHA[i][0] + '|' + PD_TAHUN[y];
const pd2Get = (i, y) => { const v = PD2S[pd2Key(i, y)]; return typeof v === 'number' && isFinite(v) ? v : null };
const pd2Set = (i, y, v) => {
  if (v === null || isNaN(v)) delete PD2S[pd2Key(i, y)]; else PD2S[pd2Key(i, y)] = v;
  try { localStorage.setItem('djpa_pdrb2', JSON.stringify(PD2S)) } catch (e) { }
};
function pd2Calc() {
  const v = PD_USAHA.map((_, i) => PD_TAHUN.map((_, y) => pd2Get(i, y)));
  const t = PD_TAHUN.map((_, y) => { const f = v.map(r => r[y]).filter(x => x !== null); return f.length ? sumArr(f) : null });
  return { v, t };
}
function buildPd2() { return ROWS.length ? { wil: wilayah() } : null }
const htmlPd2 = () => {
  const c = pd2Calc();
  const inp = 'style="width:110px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  let h = `<div class=pd2root style="overflow-x:auto"><table><tr>${th('Kode')}${th('Lapangan Usaha')}${PD_TAHUN.map(y => th(y)).join('')}</tr>`;
  PD_USAHA.forEach(([k, n], i) => h += `<tr>${td(k, 'c')}${td(n)}${PD_TAHUN.map((_, y) => `<td class=n><input class=pd2-in type=number min=0 step=any data-i=${i} data-y=${y} placeholder="isi" value="${c.v[i][y] ?? ''}" ${inp}></td>`).join('')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>Produk Domestik Regional Bruto</td>${PD_TAHUN.map((_, y) => td(pdShow(c.t[y]), 'n pd2-t')).join('')}</tr></table>
  <div class=note>Isi nilai tiap lapangan usaha secara manual (angka tanpa pemisah ribuan; kosongkan jika tidak ada). Total terhitung otomatis; isian tersimpan di browser ini.</div></div>`;
};
document.addEventListener('input', e => {
  const t = e.target; if (!t.matches || !t.matches('.pd2-in')) return;
  pd2Set(+t.dataset.i, +t.dataset.y, t.value === '' ? null : parseFloat(t.value));
  const root = t.closest('.pd2root'), c = pd2Calc();
  root.querySelectorAll('.pd2-t').forEach((el, y) => el.textContent = pdShow(c.t[y]));
});
function excelPd2(wb, ws) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  H(ws, 1, 1, 1, 1, 'Kode'); H(ws, 1, 2, 1, 2, 'Lapangan Usaha'); PD_TAHUN.forEach((y, j) => H(ws, 1, 3 + j, 1, 3 + j, y));
  const n = PD_USAHA.length, r0 = 2, r1 = 1 + n, rt = 2 + n, c = pd2Calc();
  PD_USAHA.forEach(([k, nm], i) => {
    const r = r0 + i; cell(ws, r, 1, k, { ctr: 1 }); cell(ws, r, 2, nm).alignment = { wrapText: true, vertical: 'middle' };
    PD_TAHUN.forEach((_, y) => { const x = cell(ws, r, 3 + y, c.v[i][y]); x.numFmt = '#,##0.00'; x.fill = YEL; });
  });
  cell(ws, rt, 1, 'Produk Domestik Regional Bruto', { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  PD_TAHUN.forEach((_, y) => {
    const L = String.fromCharCode(67 + y), x = cell(ws, rt, 3 + y, null, { b: 1 });
    x.value = { formula: `IF(COUNT(${L}${r0}:${L}${r1})=0,"",SUM(${L}${r0}:${L}${r1}))`, result: c.t[y] === null ? '' : c.t[y] }; x.numFmt = '#,##0.00';
  });
  ws.getCell(rt + 2, 1).value = 'Sel kuning diisi manual; baris total terhitung otomatis.';
  ws.getColumn(1).width = 10; ws.getColumn(2).width = 56; for (let k = 0; k < PD_TAHUN.length; k++) ws.getColumn(3 + k).width = 14;
}

// =====================================================================
// SHEET 18_Tabel_IV-6 - Laju Perubahan Penggunaan Tanah Reklasifikasi
//   Kategori   : GREKLAS (tahun lama, 2014) dan QREKLAS (tahun baru, 2026)
//   Luas Selisih = Luas Tahun Baru - Luas Tahun Lama
//   Laju Perubahan (V) % = Luas Selisih / Luas Tahun Lama x 100   ("-" jika luas tahun lama = 0)
//   Tahun Baru-Tahun Lama = TH_BARU - tahun lama (tahun lama bisa berbeda per kab/kota -> lihat thLamaOf di config.js)
//   Laju Perubahan per Tahun % = V / (Tahun Baru-Tahun Lama)
//   Tabel samping + grafik hanya memuat kategori yang laju per tahunnya tidak nol / tidak "-".
// =====================================================================
// const TH_LAMA = 2014, TH_BARU = 2026;
function buildRK() {
  if (!ROWS.length || !['GREKLAS', 'QREKLAS'].every(c => COLS.includes(c))) return null;
  // Tahun penggunaan tanah lama bisa berbeda tiap kabupaten/kota (lihat kotak input di bawah "Tahun penggunaan tanah").
  // Karena itu laju per tahun dihitung PER kabupaten/kota (selisih luas / selisih tahun kab/kota tsb), lalu dijumlahkan.
  const g = {}, q = {}, gk = {}, qk = {}, dtk = {}; let any = false;
  for (const r of ROWS) {
    const a = (r.GREKLAS || '').trim(), b = (r.QREKLAS || '').trim(); if (!a && !b) continue;
    any = true; const k = r.WADMKK || '', ha = area(r);
    if (!(k in dtk)) dtk[k] = TH_BARU - thLamaOf(k);
    if (a) { g[a] = (g[a] || 0) + ha; (gk[a] ??= {})[k] = (gk[a][k] || 0) + ha; }
    if (b) { q[b] = (q[b] || 0) + ha; (qk[b] ??= {})[k] = (qk[b][k] || 0) + ha; }
  }
  if (!any) return null;
  const dts = Object.values(dtk), uniform = dts.every(x => x === dts[0]), dt0 = TH_BARU - TH_LAMA;
  const rows = [...new Set([...Object.keys(g), ...Object.keys(q)])].sort((x, y) => x.localeCompare(y)).map(c => {
    const lama = g[c] || 0, baru = q[c] || 0, sel = baru - lama, v = lama ? sel / lama * 100 : null;
    let dt, py;
    if (uniform) { dt = dts[0]; py = v === null ? null : v / dt; }
    else {
      const ks = new Set([...Object.keys(gk[c] || {}), ...Object.keys(qk[c] || {})]);
      let ann = 0, wsum = 0, w = 0;                                   // ann = selisih luas per tahun (Ha/thn), dijumlah semua kab/kota
      ks.forEach(k => { const a = (gk[c] || {})[k] || 0, b = (qk[c] || {})[k] || 0; ann += (b - a) / dtk[k]; wsum += (a + b) * dtk[k]; w += a + b; });
      py = lama ? ann / lama * 100 : null;
      dt = py !== null && Math.abs(py) > 1e-12 ? v / py : (w ? wsum / w : dt0);   // selisih tahun "efektif" agar py = V / dt
    }
    return { c, lama, baru, sel, v, dt, py };
  });
  return { rows, dt: uniform ? dts[0] : null, uniform, chart: rows.filter(r => r.py !== null && Math.abs(r.py) >= 0.005), wil: wilayah() };
}
const rkN = v => v === null ? '-' : fd(v);
const rkName = r => r.c + (r.lama === 0 ? '*' : '');   // tanda * = tidak ada di tahun lama (laju tidak terdefinisi)
const rkNote = d => [d.rows.some(r => r.lama === 0) ? '* Tidak ada pada tahun lama sehingga laju perubahan tidak dapat dihitung (-).' : '',
  d.uniform ? '' : 'Tahun penggunaan tanah lama berbeda antar kabupaten/kota: laju per tahun dihitung per kabupaten/kota lalu dijumlahkan; kolom Tahun Baru-Tahun Lama adalah selisih tahun efektif (rata-rata tertimbang).'].filter(Boolean).join(' ');
const dtTxt = v => String(+v.toFixed(2));
const htmlRK = (d, ex) => {
  let h = `<table><tr>${['No', 'Penggunaan Tanah Reklasifikasi', 'Luas Selisih (Lt-Lt-1) (Ha)', 'Luas Tahun Baru (Lt) (Ha)', 'Luas Tahun Lama (Lt-1) (Ha)', 'Laju Perubahan (V) (%)', 'Tahun Baru-Tahun Lama', 'Laju Perubahan per Tahun (%)'].map(t => th(t)).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(rkName(r))}${td(fd(r.sel), 'n')}${td(fd(r.baru), 'n')}${td(fd(r.lama), 'n')}${td(rkN(r.v), 'n')}${td(dtTxt(r.dt), 'c')}${td(rkN(r.py), 'n')}</tr>`);
  h += '</table>' + (rkNote(d) ? `<div class=note>${rkNote(d)}</div>` : '');
  const side = `<table><tr>${th('Penggunaan Tanah Reklasifikasi')}${d.chart.map(r => th(r.c)).join('')}</tr>
   <tr>${th('Laju Pertumbuhan (%)')}${d.chart.map(r => td(fd(r.py), 'n')).join('')}</tr></table>`;
  return `<div class=wrap><div>${h}</div><div class=side><div style="overflow-x:auto">${side}</div>${cvTag(ex)}</div></div>`;
};
const chartRK = (cv, d, a) => new Chart(cv, {
  type: 'bar',
  data: { labels: [''], datasets: d.chart.map((r, i) => ({ label: r.c, data: [r.py], backgroundColor: PAL2[i % PAL2.length] })) },
  options: { ...baseOpt(['LAJU PERUBAHAN PENGGUNAAN TANAH REKLASIFIKASI', d.wil.toUpperCase() + ' TAHUN ' + yr('2014-2026') + ' (%)'], a, 'right'), scales: { y: { beginAtZero: true } } },
  plugins: [bgPlugin, multiVals]
});
function excelRK(wb, ws, d, ex) {
  ['No', 'Penggunaan Tanah Reklasifikasi', 'Luas Selisih (Lt-Lt-1) (Ha)', 'Luas Tahun Baru (Lt) (Ha)', 'Luas Tahun Lama (Lt-1) (Ha)', 'Laju Perubahan (V) (%)', 'Tahun Baru-Tahun Lama', 'Laju Perubahan per Tahun (%)'].forEach((t, i) => H(ws, 1, i + 1, 1, i + 1, t));
  const fx = (r, c, formula, result) => { const x = cell(ws, r, c, null, { num: 1 }); x.value = { formula, result }; x.alignment = { horizontal: 'right' }; return x; };
  d.rows.forEach((r, i) => {
    const n = 2 + i;
    cell(ws, n, 1, i + 1, { ctr: 1 }); cell(ws, n, 2, rkName(r));
    fx(n, 3, `D${n}-E${n}`, r.sel);
    const fL = rawRng('LUASHA'), fG = rawRng('GREKLAS'), fQ = rawRng('QREKLAS'), canF = !!(fL && fG && fQ);
    cell(ws, n, 4, canF ? FX(`SUMIFS(${fL},${fQ},${qs(r.c)})`, r.baru) : r.baru, { num: 1 });   // luas tahun baru
    cell(ws, n, 5, canF ? FX(`SUMIFS(${fL},${fG},${qs(r.c)})`, r.lama) : r.lama, { num: 1 });   // luas tahun lama
    fx(n, 6, `IF(E${n}=0,"-",C${n}/E${n}*100)`, r.v === null ? '-' : r.v);
    const g = cell(ws, n, 7, r.dt, { ctr: 1 }); if (d.uniform) g.value = { formula: `${TH_BARU}-${TH_LAMA}`, result: r.dt }; else g.numFmt = '0.00';
    fx(n, 8, `IF(ISNUMBER(F${n}),F${n}/G${n},"-")`, r.py === null ? '-' : r.py);
  });
  const nt = rkNote(d); if (nt) ws.getCell(d.rows.length + 3, 2).value = nt;
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 28; for (let c = 3; c <= 8; c++) ws.getColumn(c).width = 22;
  // tabel samping (kategori jadi kolom) + grafik
  H(ws, 1, 10, 1, 10, 'Penggunaan Tanah Reklasifikasi'); H(ws, 2, 10, 2, 10, 'Laju Pertumbuhan (%)');
  d.chart.forEach((r, j) => {
    H(ws, 1, 11 + j, 1, 11 + j, r.c);
    const x = cell(ws, 2, 11 + j, null, { num: 1 }); x.value = { formula: `H${2 + d.rows.indexOf(r)}`, result: r.py };
  });
  ws.getColumn(10).width = 32; for (let j = 0; j < d.chart.length; j++) ws.getColumn(11 + j).width = 20;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  const qq = `'${ws.name.replace(/'/g, "''")}'!`;
  addXChart({ sheet: ws.name, type: 'bar', legend: 'r', from: { col: 9, row: 3 }, ext: { w: ex.w, h: ex.h },
    title: 'LAJU PERUBAHAN PENGGUNAAN TANAH REKLASIFIKASI ' + d.wil.toUpperCase() + ' TAHUN ' + yr('2014-2026') + ' (%)',
    series: d.chart.map((r, j) => ({ name: r.c, nameRef: `${qq}${A1(1, 11 + j, 1, 1)}`, cat: `${qq}$J$2`, catVals: ['Laju Pertumbuhan (%)'], val: `${qq}${A1(2, 11 + j, 1, 1)}`, vals: [r.py], color: PAL2[j % PAL2.length] })) });
}

// =====================================================================
// SHEET 29_Tabel_IV-18 - Potensi Sosial Ekonomi per Kabupaten/Kota
//   Ketersediaan Tanah  : dari kolom Vname (Tersedia / Tidak Tersedia)
//   Penguasaan Tanah    : Belum Ada HAT / Ada HAT -> BELUM ADA DATA, dibiarkan KOSONG (kotak isian manual).
//                         Luas (%) = luas HAT / luas kabupaten/kota x 100 (luas kabupaten/kota = Tersedia + Tidak Tersedia)
//   Kepadatan Penduduk  : diambil otomatis dari Tabel III-1 (Jiwa / (Ha/100))
//   Butuh blok IV-9/IV-10 (vCol, ketIdx) dan blok III-1 (popGet, popCalc, fInt) sudah terpasang.
// =====================================================================
const SKS = (() => { try { return JSON.parse(localStorage.getItem('djpa_hat') || '{}') } catch (e) { return {} } })();   // {kunci: Ha}
const skKey = (k, j) => wilayah() + '|' + k + '|' + j;
const skGet = (k, j) => { const v = SKS[skKey(k, j)]; return typeof v === 'number' && isFinite(v) ? v : null };
const skSet = (k, j, v) => {
  if (v === null || isNaN(v)) delete SKS[skKey(k, j)]; else SKS[skKey(k, j)] = v;
  try { localStorage.setItem('djpa_hat', JSON.stringify(SKS)) } catch (e) { }
};
let SKD = null, SKC = null;   // data tabel + grafik yang sedang tampil
function buildSK() {
  const V = vCol(); if (!ROWS.length || !V || !COLS.includes('WADMKC')) return null;
  const m = {};
  for (const r of ROWS) {
    const k = kecOf(r); m[k] ??= { a: 0, b: 0, u: 0, ha: 0 }; m[k].ha += area(r);
    const i = ketIdx(r[V]); if (i === 0) m[k].a += area(r); else if (i === 1) m[k].b += area(r); else if (i === 2) m[k].u += area(r);   // 2 = Tidak Dapat Dianalisis
  }
  const rows = Object.keys(m).sort().map(k => { const t = m[k].a + m[k].b + m[k].u; return { k, a: m[k].a, b: m[k].b, u: m[k].u, t, pa: pc(m[k].a, t), pb: pc(m[k].b, t), pu: pc(m[k].u, t), ha: m[k].ha } });
  const a = sumArr(rows.map(r => r.a)), b = sumArr(rows.map(r => r.b)), u = sumArr(rows.map(r => r.u)), t = a + b + u;
  return (SKD = { rows, tot: { a, b, u, t, pa: pc(a, t), pb: pc(b, t), pu: pc(u, t) }, totHa: sumArr(rows.map(r => r.ha)), wil: wilayah() });
}
function skCalc(d) {
  const hb = d.rows.map(r => skGet(r.k, 'b')), hh = d.rows.map(r => skGet(r.k, 'a'));
  const pct = (v, r) => v !== null && r.t ? v / r.t * 100 : null;
  const sm = a => { const f = a.filter(v => v !== null); return f.length ? sumArr(f) : null };
  const tb = sm(hb), th_ = sm(hh), pp = popCalc(d.rows.map(r => ({ k: r.k, ha: r.ha })), d.totHa);
  return {
    hb, hh, pb: hb.map((v, i) => pct(v, d.rows[i])), ph: hh.map((v, i) => pct(v, d.rows[i])),
    tb, th: th_, tbp: tb !== null && d.tot.t ? tb / d.tot.t * 100 : null, thp: th_ !== null && d.tot.t ? th_ / d.tot.t * 100 : null,
    kp: pp.kp, tkp: pp.tkp
  };
}
const skD2 = v => v === null ? '-' : fd(v), skI = v => v === null ? '-' : fInt(v);
const htmlSK = (d, ex) => {
  const c = skCalc(d);
  const inp = 'style="width:100px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  const box = (k, j, v) => `<input class=sk-in type=number min=0 step=any placeholder="isi" data-k="${String(k).replace(/"/g, '&quot;')}" data-j=${j} value="${v ?? ''}" ${inp}>`;
  let h = `<div class=skroot><div style="overflow-x:auto"><table><tr>${th('No', 'rowspan=3')}${th('Kabupaten/Kota', 'rowspan=3')}${th('Ketersediaan Tanah', 'colspan=6')}${th('Penguasaan Tanah', 'colspan=4')}${th('Kepadatan Penduduk (jiwa/km2)', 'rowspan=3')}</tr>
  <tr>${th('Tersedia', 'colspan=2')}${th('Tidak Tersedia', 'colspan=2')}${th(TDA_LABEL, 'colspan=2')}${th('Belum Ada HAT', 'colspan=2')}${th('Ada HAT', 'colspan=2')}</tr>
  <tr>${[1, 2, 3, 4, 5].map(() => th('Luas (Ha)') + th('Luas (%)')).join('')}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}${td(fd(r.a), 'n')}${td(fd(r.pa), 'n')}${td(fd(r.b), 'n')}${td(fd(r.pb), 'n')}${td(fd(r.u), 'n')}${td(fd(r.pu), 'n')}<td class=n>${box(r.k, 'b', c.hb[i])}</td>${td(skD2(c.pb[i]), 'n sk-pb')}<td class=n>${box(r.k, 'a', c.hh[i])}</td>${td(skD2(c.ph[i]), 'n sk-ph')}${td(skI(c.kp[i]), 'n sk-kp')}</tr>`);
  h += `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${td(fd(d.tot.a), 'n')}${td(fd(d.tot.pa), 'n')}${td(fd(d.tot.b), 'n')}${td(fd(d.tot.pb), 'n')}${td(fd(d.tot.u), 'n')}${td(fd(d.tot.pu), 'n')}${td(skD2(c.tb), 'n sk-tb')}${td(skD2(c.tbp), 'n sk-tbp')}${td(skD2(c.th), 'n sk-th')}${td(skD2(c.thp), 'n sk-thp')}${td(skI(c.tkp), 'n sk-tkp')}</tr></table>
  <div class=note>Data Penguasaan Tanah (HAT) belum tersedia: kolom <b>Belum Ada HAT</b> dan <b>Ada HAT</b> (Ha) dikosongkan dan dapat diisi manual. Luas (%) HAT dihitung terhadap luas kabupaten/kota. Kepadatan Penduduk otomatis dari Tabel III-1 (isi Jumlah Penduduk di sana).</div></div>${cvTag(ex)}</div>`;
  return h;
};
function skRefresh() {
  const root = document.querySelector('.skroot'); if (!root || !SKD) return;
  const c = skCalc(SKD), set = (cls, arr) => root.querySelectorAll(cls).forEach((el, i) => el.textContent = arr[i]);
  set('.sk-pb', c.pb.map(skD2)); set('.sk-ph', c.ph.map(skD2)); set('.sk-kp', c.kp.map(skI));
  const one = (cls, v) => root.querySelector(cls).textContent = v;
  one('.sk-tb', skD2(c.tb)); one('.sk-tbp', skD2(c.tbp)); one('.sk-th', skD2(c.th)); one('.sk-thp', skD2(c.thp)); one('.sk-tkp', skI(c.tkp));
  if (SKC && SKC.canvas && document.body.contains(SKC.canvas)) {
    SKC.data.datasets[1].data = c.hh.map(v => v || 0); SKC.data.datasets[2].data = c.kp.map(v => v || 0); SKC.update('none');
  }
}
document.addEventListener('input', e => {
  const t = e.target; if (!t.matches) return;
  if (t.matches('.sk-in')) { skSet(t.dataset.k, t.dataset.j, t.value === '' ? null : parseFloat(t.value)); skRefresh(); }
  else if (t.matches('.pop-in')) setTimeout(skRefresh, 0);   // Jumlah Penduduk diubah di Tabel III-1
});
// label angka di atas batang (Kepadatan = bilangan bulat, lainnya 2 desimal; nilai 0 tidak ditampilkan)
const skVals = {
  id: 'skv', afterDatasetsDraw(ch) {
    const c = ch.ctx; c.save(); c.font = '10px Arial'; c.fillStyle = '#222'; c.textAlign = 'center';
    ch.data.datasets.forEach((ds, i) => ch.getDatasetMeta(i).data.forEach((el, j) => { const v = ds.data[j]; if (v > 0) c.fillText(i === 2 ? fInt(v) : fmt(v), el.x, el.y - 5); }));
    c.restore();
  }
};
const chartSK = (cv, d, a) => {
  const c = skCalc(d);
  const ch = new Chart(cv, {
    type: 'bar',
    data: {
      labels: d.rows.map(r => r.k), datasets: [
        { label: 'Tersedia', data: d.rows.map(r => r.a), backgroundColor: '#ffc000' },
        { label: 'Ada HAT', data: c.hh.map(v => v || 0), backgroundColor: NAVY },
        { label: 'Kepadatan Penduduk (jiwa/km2)', data: c.kp.map(v => v || 0), backgroundColor: '#c00000' }]
    },
    options: { ...baseOpt('POTENSI SOSIAL EKONOMI ' + d.wil.toUpperCase() + ' TAHUN 2026', a, 'bottom'), scales: { y: { beginAtZero: true } } },
    plugins: [bgPlugin, skVals]
  });
  if (a !== false) SKC = ch;   // addChart (Excel) memanggil dengan a=false -> bukan grafik tampilan
  return ch;
};
function excelSK(wb, ws, d, ex) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  const s1 = (EXTRA.find(e => e.build === buildT3) || { tab: '01_Tabel_III-1' }).tab;   // sheet Tabel III-1 (sumber kepadatan)
  // kolom: C,D Tersedia | E,F Tidak Tersedia | G,H Tidak Dapat Dianalisis | I,J Belum Ada HAT | K,L Ada HAT | M Kepadatan
  H(ws, 1, 1, 3, 1, 'No'); H(ws, 1, 2, 3, 2, 'Kabupaten/Kota'); H(ws, 1, 3, 1, 8, 'Ketersediaan Tanah'); H(ws, 1, 9, 1, 12, 'Penguasaan Tanah'); H(ws, 1, 13, 3, 13, 'Kepadatan Penduduk (jiwa/km2)');
  [['Tersedia', 3], ['Tidak Tersedia', 5], [TDA_LABEL, 7], ['Belum Ada HAT', 9], ['Ada HAT', 11]].forEach(([t, c]) => H(ws, 2, c, 2, c + 1, t));
  for (let i = 0; i < 10; i++) H(ws, 3, 3 + i, 3, 3 + i, i % 2 ? 'Luas (%)' : 'Luas (Ha)');
  const n = d.rows.length, r0 = 4, r1 = 3 + n, rt = 4 + n, c = skCalc(d);
  const fx = (r, col, formula, result, f, b) => { const x = cell(ws, r, col, null, { b }); x.value = { formula, result }; x.numFmt = f; return x; };
  const inb = (r, col, v) => { const x = cell(ws, r, col, v); x.numFmt = '#,##0.00'; x.fill = YEL; };
  const den = x => `(C${x}+E${x}+G${x})`;   // luas kabupaten/kota = Tersedia + Tidak Tersedia + Tidak Dapat Dianalisis
  d.rows.forEach((r, i) => {
    const x = r0 + i, src = `'${s1}'!G${3 + i}`;                   // sel Kepadatan di Tabel III-1 (baris data mulai 3)
    cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.k);
    const fK = rawRng('WADMKC'), ex = fK && r.k !== '-' ? [[fK, `$B${x}`]] : null, fa = ex && ketSumF(0, ex), fb = ex && ketSumF(1, ex), fu = ex && ketSumF(2, ex);
    cell(ws, x, 3, fa ? FX(fa, r.a) : r.a, { num: 1 }); cell(ws, x, 4, FX(`IF(${den(x)}=0,0,C${x}/${den(x)}*100)`, r.pa), { num: 1 });
    cell(ws, x, 5, fb ? FX(fb, r.b) : r.b, { num: 1 }); cell(ws, x, 6, FX(`IF(${den(x)}=0,0,E${x}/${den(x)}*100)`, r.pb), { num: 1 });
    cell(ws, x, 7, fu ? FX(fu, r.u) : r.u, { num: 1 }); cell(ws, x, 8, FX(`IF(${den(x)}=0,0,G${x}/${den(x)}*100)`, r.pu), { num: 1 });
    inb(x, 9, c.hb[i]); fx(x, 10, `IF(I${x}="","",I${x}/${den(x)}*100)`, c.pb[i] === null ? '' : c.pb[i], '#,##0.00');
    inb(x, 11, c.hh[i]); fx(x, 12, `IF(K${x}="","",K${x}/${den(x)}*100)`, c.ph[i] === null ? '' : c.ph[i], '#,##0.00');
    fx(x, 13, `IF(${src}="","",${src})`, c.kp[i] === null ? '' : c.kp[i], '#,##0');
  });
  cell(ws, rt, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  cell(ws, rt, 3, FX(`SUM(C${r0}:C${r1})`, d.tot.a), { num: 1, b: 1 }); cell(ws, rt, 5, FX(`SUM(E${r0}:E${r1})`, d.tot.b), { num: 1, b: 1 }); cell(ws, rt, 7, FX(`SUM(G${r0}:G${r1})`, d.tot.u), { num: 1, b: 1 });
  cell(ws, rt, 4, FX(`IF(${den(rt)}=0,0,C${rt}/${den(rt)}*100)`, d.tot.pa), { num: 1, b: 1 }); cell(ws, rt, 6, FX(`IF(${den(rt)}=0,0,E${rt}/${den(rt)}*100)`, d.tot.pb), { num: 1, b: 1 }); cell(ws, rt, 8, FX(`IF(${den(rt)}=0,0,G${rt}/${den(rt)}*100)`, d.tot.pu), { num: 1, b: 1 });
  fx(rt, 9, `IF(COUNT(I${r0}:I${r1})=0,"",SUM(I${r0}:I${r1}))`, c.tb === null ? '' : c.tb, '#,##0.00', 1);
  fx(rt, 10, `IF(I${rt}="","",I${rt}/${den(rt)}*100)`, c.tbp === null ? '' : c.tbp, '#,##0.00', 1);
  fx(rt, 11, `IF(COUNT(K${r0}:K${r1})=0,"",SUM(K${r0}:K${r1}))`, c.th === null ? '' : c.th, '#,##0.00', 1);
  fx(rt, 12, `IF(K${rt}="","",K${rt}/${den(rt)}*100)`, c.thp === null ? '' : c.thp, '#,##0.00', 1);
  const st = `'${s1}'!G${3 + n}`;
  fx(rt, 13, `IF(${st}="","",${st})`, c.tkp === null ? '' : c.tkp, '#,##0', 1);
  ws.getCell(rt + 1, 1).value = 'Sel kuning (Belum Ada HAT / Ada HAT, Ha) belum ada datanya dan dapat diisi manual. Kepadatan Penduduk mengikuti sheet Tabel III-1.';
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 20; for (let k = 3; k <= 12; k++) ws.getColumn(k).width = 14; ws.getColumn(13).width = 20;
  const qq = `'${ws.name.replace(/'/g, "''")}'!`, cat = `${qq}$B$${r0}:$B$${r1}`, R = L => `${qq}$${L}$${r0}:$${L}$${r1}`, cats = d.rows.map(x => x.k);
  addXChart({ sheet: ws.name, type: 'bar', legend: 'b', from: { col: 2, row: rt + 2 }, ext: { w: ex.w, h: ex.h }, title: 'POTENSI SOSIAL EKONOMI ' + d.wil.toUpperCase() + ' TAHUN 2026',
    series: [{ name: 'Tersedia', cat, catVals: cats, val: R('C'), vals: d.rows.map(x => x.a), color: '#ffc000' },
             { name: 'Ada HAT', cat, catVals: cats, val: R('K'), vals: c.hh.map(v => v || 0), color: NAVY },
             { name: 'Kepadatan Penduduk (jiwa/km2)', cat, catVals: cats, val: R('M'), vals: c.kp.map(v => v || 0), color: '#c00000' }] });
}

// =====================================================================
// SHEET 09_Tabel_III-9 - Galat Tanah Belum Terdaftar Hasil Unduhan Geo-KKP   (KERANGKA)
//   Jumlah Bidang dan Luas (Ha) per kabupaten/kota: belum ada data (nanti dari file HAT) -> dikosongkan, kotak isian manual.
//   % terhadap Tanah Belum Terdaftar = Luas / total "Belum Ada HAT" (dari Tabel IV-18)  x 100
//   % terhadap Luas Wilayah          = Luas / luas wilayah (dari data LUASHA)           x 100
//   Butuh blok IV-18 (skGet, SKS) dan blok III-1 (fInt) sudah terpasang.
// =====================================================================
const GLS = (() => { try { return JSON.parse(localStorage.getItem('djpa_galat') || '{}') } catch (e) { return {} } })();
const glKey = (k, j) => wilayah() + '|' + k + '|' + j;                       // j: 'b' = jumlah bidang, 'l' = luas (Ha)
const glGet = (k, j) => { const v = GLS[glKey(k, j)]; return typeof v === 'number' && isFinite(v) ? v : null };
const glSet = (k, j, v) => {
  if (v === null || isNaN(v)) delete GLS[glKey(k, j)]; else GLS[glKey(k, j)] = v;
  try { localStorage.setItem('djpa_galat', JSON.stringify(GLS)) } catch (e) { }
};
let GLD = null;
function buildGl() {
  if (!ROWS.length || !COLS.includes('WADMKC')) return null;
  const ks = [...new Set(ROWS.map(kecOf))].sort();
  return (GLD = { rows: ks.map(k => ({ k })), totHa: sumArr(ROWS.map(area)), wil: wilayah() });
}
function glCalc(d) {
  const sm = a => { const f = a.filter(v => v !== null); return f.length ? sumArr(f) : null };
  const b = d.rows.map(r => glGet(r.k, 'b')), l = d.rows.map(r => glGet(r.k, 'l'));
  const bt = sm(d.rows.map(r => skGet(r.k, 'b')));                            // total Belum Ada HAT (Tabel IV-18)
  const p1 = v => v !== null && bt ? v / bt * 100 : null, p2 = v => v !== null && d.totHa ? v / d.totHa * 100 : null;
  const tb = sm(b), tl = sm(l);
  return { b, l, e: l.map(p1), f: l.map(p2), tb, tl, te: p1(tl), tf: p2(tl), bt };
}
const glI = v => v === null ? '-' : fInt(v), glD = v => v === null ? '-' : fd(v);
const htmlGl = d => {
  const c = glCalc(d);
  const inp = 'style="width:100px;text-align:right;padding:4px 8px;border:1px solid var(--bd);border-radius:6px;background:transparent;color:inherit;font:inherit"';
  const box = (k, j, v) => `<input class=gl-in type=number min=0 step=any placeholder="isi" data-k="${String(k).replace(/"/g, '&quot;')}" data-j=${j} value="${v ?? ''}" ${inp}>`;
  let h = `<div class=glroot style="overflow-x:auto"><table><tr>${th('No', 'rowspan=2')}${th('Kabupaten/Kota', 'rowspan=2')}${th('Galat Tanah Belum Terdaftar Hasil Unduhan Geo-KKP', 'colspan=4')}</tr>
  <tr>${th('Jumlah Bidang')}${th('Luas (Ha)')}${th('% terhadap Tanah Belum Terdaftar')}${th('% terhadap Luas ' + d.wil)}</tr>`;
  d.rows.forEach((r, i) => h += `<tr>${td(i + 1, 'c')}${td(r.k)}<td class=n>${box(r.k, 'b', c.b[i])}</td><td class=n>${box(r.k, 'l', c.l[i])}</td>${td(glD(c.e[i]), 'n gl-e')}${td(glD(c.f[i]), 'n gl-f')}</tr>`);
  return h + `<tr class=tot><td colspan=2 class=c>${d.wil}</td>${td(glI(c.tb), 'n gl-tb')}${td(glD(c.tl), 'n gl-tl')}${td(glD(c.te), 'n gl-te')}${td(glD(c.tf), 'n gl-tf')}</tr></table>
  <div class=note>Kerangka: data HAT belum tersedia, jadi Jumlah Bidang dan Luas (Ha) dikosongkan (dapat diisi manual; nanti terisi otomatis dari file HAT). % terhadap Tanah Belum Terdaftar memakai total <b>Belum Ada HAT</b> dari Tabel IV-18, jadi baru terhitung setelah kolom itu terisi.</div></div>`;
};
function glRefresh() {
  const root = document.querySelector('.glroot'); if (!root || !GLD) return;
  const c = glCalc(GLD), set = (cls, arr) => root.querySelectorAll(cls).forEach((el, i) => el.textContent = arr[i]);
  set('.gl-e', c.e.map(glD)); set('.gl-f', c.f.map(glD));
  const one = (cls, v) => root.querySelector(cls).textContent = v;
  one('.gl-tb', glI(c.tb)); one('.gl-tl', glD(c.tl)); one('.gl-te', glD(c.te)); one('.gl-tf', glD(c.tf));
}
document.addEventListener('input', e => {
  const t = e.target; if (!t.matches) return;
  if (t.matches('.gl-in')) { glSet(t.dataset.k, t.dataset.j, t.value === '' ? null : parseFloat(t.value)); glRefresh(); }
  else if (t.matches('.sk-in')) setTimeout(glRefresh, 0);   // "Belum Ada HAT" diubah di Tabel IV-18
});
function excelGl(wb, ws, d) {
  const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  const s18 = (EXTRA.find(e => e.build === buildSK) || { tab: '29_Tabel_IV-18' }).tab, q = `'${s18}'!`;
  const rt18 = 4 + d.rows.length, bel = `${q}I${rt18}`, kota = `(${q}C${rt18}+${q}E${rt18}+${q}G${rt18})`;   // total Belum Ada HAT dan luas wilayah di Tabel IV-18
  H(ws, 1, 1, 2, 1, 'No'); H(ws, 1, 2, 2, 2, 'Kabupaten/Kota'); H(ws, 1, 3, 1, 6, 'Galat Tanah Belum Terdaftar Hasil Unduhan Geo-KKP');
  ['Jumlah Bidang', 'Luas (Ha)', '% terhadap Tanah Belum Terdaftar', '% terhadap Luas ' + d.wil].forEach((t, i) => H(ws, 2, 3 + i, 2, 3 + i, t));
  const n = d.rows.length, r0 = 3, r1 = 2 + n, rt = 3 + n, c = glCalc(d);
  const fx = (r, col, formula, result, f, b) => { const x = cell(ws, r, col, null, { b }); x.value = { formula, result }; x.numFmt = f; return x; };
  const inb = (r, col, v, f) => { const x = cell(ws, r, col, v); x.numFmt = f; x.fill = YEL; };
  d.rows.forEach((r, i) => {
    const x = r0 + i; cell(ws, x, 1, i + 1, { ctr: 1 }); cell(ws, x, 2, r.k);
    inb(x, 3, c.b[i], '#,##0'); inb(x, 4, c.l[i], '#,##0.00');
    fx(x, 5, `IF(OR(D${x}="",${bel}="",${bel}=0),"",D${x}/${bel}*100)`, c.e[i] === null ? '' : c.e[i], '#,##0.00');
    fx(x, 6, `IF(OR(D${x}="",${kota}=0),"",D${x}/${kota}*100)`, c.f[i] === null ? '' : c.f[i], '#,##0.00');
  });
  cell(ws, rt, 1, d.wil, { b: 1, ctr: 1 }); cell(ws, rt, 2, null); ws.mergeCells(rt, 1, rt, 2);
  fx(rt, 3, `IF(COUNT(C${r0}:C${r1})=0,"",SUM(C${r0}:C${r1}))`, c.tb === null ? '' : c.tb, '#,##0', 1);
  fx(rt, 4, `IF(COUNT(D${r0}:D${r1})=0,"",SUM(D${r0}:D${r1}))`, c.tl === null ? '' : c.tl, '#,##0.00', 1);
  fx(rt, 5, `IF(OR(D${rt}="",${bel}="",${bel}=0),"",D${rt}/${bel}*100)`, c.te === null ? '' : c.te, '#,##0.00', 1);
  fx(rt, 6, `IF(OR(D${rt}="",${kota}=0),"",D${rt}/${kota}*100)`, c.tf === null ? '' : c.tf, '#,##0.00', 1);
  ws.getCell(rt + 2, 1).value = 'Kerangka: Jumlah Bidang dan Luas (sel kuning) belum ada datanya (nanti dari file HAT). % terhadap Tanah Belum Terdaftar memakai total Belum Ada HAT di sheet Tabel IV-18.';
  ws.getColumn(1).width = 6; ws.getColumn(2).width = 22; for (let k = 3; k <= 6; k++) ws.getColumn(k).width = 22;
  ws.getRow(2).height = 36;
}

// =====================================================================
// DAFTAR SHEET TAMBAHAN (urutan = urutan sheet 7, 8, 9, 10)
// =====================================================================
const EXTRA = [
  { tab: '12_Tabel_IV-1', build: build7, html: html7, chart: chart7, excel: excel7, w: 640, h: 340 },
  { tab: '13_Tabel_IV-2', build: build8, html: html8, chart: null, excel: excel8 },
  { tab: '15_Grafik_IV-1', build: build9, html: html9, chart: chart9, excel: excel9, w: 640, h: 340 },
  { tab: '17_Tabel_IV-5', build: build10, html: html10, chart: chart10, excel: excel10, w: 640, h: 340 },
  { tab: '28_Tabel_IV-17', build: build17, html: html17, chart: chart17, excel: excel17, w: 760, h: 380 },
  { tab: '25_Tabel_IV-14', build: build14, html: html14, chart: chart14, excel: excel14, w: 760, h: 340 },
  { tab: '26_Tabel_IV-15', build: build15, html: html15, chart: chart15, excel: excel15, w: 760, h: 400 },
  { tab: '27_Tabel_IV-16', build: build16, html: html16, chart: chart16, excel: excel16, w: 760, h: 400 },
  { tab: '24_Tabel_IV-12', build: build12, html: html12, chart: null, excel: excel12 },
  { tab: '23_Tabel_IV-11', build: build11, html: html11, chart: null, excel: excel11 },
  { tab: '22_Tabel_IV-10', build: () => buildKetPola('QNAME25'), html: htmlKetPola('Penggunaan Tanah'), chart: null, excel: excelKetPola('Penggunaan Tanah') },
  { tab: '21_Tabel_IV-9', build: buildKetKec, html: htmlKetKec, chart: chartKetKec, excel: excelKetKec, w: 640, h: 340 },
  { tab: '20_Tabel_IV-8', build: buildKes, html: htmlKes, chart: chartKes, excel: excelKes, w: 640, h: 340 },
  { tab: '19_Tabel_IV-7', build: buildMatKs, html: htmlMatKs, chart: null, excel: excelMatKs },
  { tab: '16_Tabel_IV-4', build: buildGQ, html: htmlGQ, chart: null, excel: excelGQ },
  { tab: '14_Tabel_IV-3', build: buildMatLB, html: htmlMatLB, chart: null, excel: excelMatLB },
  { tab: '01_Tabel_III-1', build: buildT3, html: htmlT3, chart: chartT3, excel: excelT3, w: 640, h: 300 },
  { tab: '02_Tabel_III-2', build: buildT2, html: htmlT2, chart: null, excel: excelT2 },
  { tab: '03_Tabel_III-3', build: buildPk, html: htmlPk, chart: chartPk, excel: excelPk, w: 640, h: 420 },
  { tab: '04_Tabel_III-4', build: buildPd, html: htmlPd, chart: null, excel: excelPd },
  { tab: '05_Tabel_III-5', build: buildPd2, html: htmlPd2, chart: null, excel: excelPd2 },
  { tab: '18_Tabel_IV-6', build: buildRK, html: htmlRK, chart: chartRK, excel: excelRK, w: 720, h: 340 },
  { tab: '29_Tabel_IV-18', build: buildSK, html: htmlSK, chart: chartSK, excel: excelSK, w: 720, h: 380 },
  { tab: '09_Tabel_III-9', build: buildGl, html: htmlGl, chart: null, excel: excelGl }
];
