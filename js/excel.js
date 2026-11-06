// excel.js
// Membuat file Excel 31 sheet (urutan mengikuti SHEETS di sheets.js) memakai ExcelJS (grafik disisipkan sebagai gambar PNG).

const thin = { style: 'thin', color: { argb: 'FF000000' } }, BORD = { top: thin, left: thin, bottom: thin, right: thin };
function hdr(c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; c.border = BORD }
function cell(ws, r, c, v, o = {}) { const x = ws.getCell(r, c); x.value = v; x.border = BORD; if (o.num) x.numFmt = '#,##0.00'; if (o.b) x.font = { bold: true }; if (o.ctr) x.alignment = { horizontal: 'center' }; return x }
function chartPNG(def, d) {
    const cv = document.createElement('canvas'); cv.width = def.type === 'pie' ? 760 : 600; cv.height = 340;
    const ch = mkChart(cv, def, d, false), url = cv.toDataURL('image/png'); ch.destroy(); return url;
}
function sheetRaw(wb, name) {
    const s1 = wb.addWorksheet(name);
    if (ROWS.length) {
        s1.addRow(COLS).eachCell(hdr);
        const num = new Set(COLS.filter(c => /^(KDGN|TON|LUASHA|Shape)/.test(c) && ROWS.some(r => r[c] !== '' && !isNaN(r[c]))));   // kolom angka (dicek sampai ada baris berisi angka, bukan hanya baris 1)
        ROWS.forEach(r => {
            // angka notasi ilmiah (SCI, ada di render.js) dibulatkan 3 angka di belakang koma
            const row = s1.addRow(COLS.map(c => SCI.test(r[c]) ? +r[c] : num.has(c) ? (parseFloat(r[c]) || 0) : r[c]));
            COLS.forEach((c, i) => { if (SCI.test(r[c])) row.getCell(i + 1).numFmt = '0.00' });
        });
        s1.views = [{ state: 'frozen', ySplit: 1 }]; s1.columns.forEach(c => c.width = 16);
    } else s1.getCell(1, 1).value = '(Data tidak ditemukan)';
}
function sheetDef(wb, def, d, name) {
    const ws = wb.addWorksheet(name);
    if (!d) { ws.getCell(1, 1).value = '(Data tidak ditemukan)'; return }
    const nk = d.kecs.length, cJ = nk + 3, cP = nk + 4;
    [[1, 1, 'No'], [1, 2, def.h], [1, 3, 'Luas di Tiap Kabupaten/Kota (Ha)'], [1, cJ, 'Jumlah (Ha)'], [1, cP, 'Persentase (%)']].forEach(([r, c, v]) => hdr(cell(ws, r, c, v)));
    d.kecs.forEach((k, j) => hdr(cell(ws, 2, 3 + j, k)));
    ws.mergeCells(1, 1, 2, 1); ws.mergeCells(1, 2, 2, 2); ws.mergeCells(1, 3, 1, 2 + nk); ws.mergeCells(1, cJ, 2, cJ); ws.mergeCells(1, cP, 2, cP);
    for (let c = 1; c <= cP; c++)for (let r = 1; r <= 2; r++)hdr(ws.getCell(r, c));

    // ===== RUMUS (baru) =====
    const r1 = 3, rL = 2 + d.body.length, tr = rL + 1;                 // baris data pertama, terakhir, baris total
    const fLuas = rawRng('LUASHA'), fKat = rawRng(def.f), fKec = rawRng('WADMKC');
    const L = colLetter, canF = !!(fLuas && fKat && fKec);
    d.body.forEach((b, i) => {
        const r = r1 + i; cell(ws, r, 1, i + 1, { ctr: 1 }); cell(ws, r, 2, b.c);
        b.v.forEach((v, j) => {
            // luas = SUMIFS(LUASHA, kolom kategori = kategori di kolom B, kolom WADMKC = nama kabupaten/kota di baris 2)
            const f = canF && d.kecs[j] !== '-' ? FX(`SUMIFS(${fLuas},${fKat},$B${r},${fKec},${L(3 + j)}$2)`, v) : v;
            cell(ws, r, 3 + j, f, { num: 1 });
        });
        cell(ws, r, cJ, FX(`SUM(${A1(r, 3)}:${A1(r, 2 + nk)})`, b.t), { num: 1, b: 1 });                                   // Jumlah (Ha)
        cell(ws, r, cP, FX(`IF(${A1(tr, cJ, 1, 1)}=0,0,${A1(r, cJ)}/${A1(tr, cJ, 1, 1)}*100)`, b.p), { num: 1, b: 1 });    // Persentase (%)
    });
    cell(ws, tr, 1, 'Jumlah (Ha)', { b: 1, ctr: 1 }); cell(ws, tr, 2, null); ws.mergeCells(tr, 1, tr, 2);
    d.kt.forEach((v, j) => cell(ws, tr, 3 + j, FX(`SUM(${A1(r1, 3 + j)}:${A1(rL, 3 + j)})`, v), { num: 1, b: 1 }));        // total per kabupaten/kota
    cell(ws, tr, cJ, FX(`SUM(${A1(r1, cJ)}:${A1(rL, cJ)})`, d.T), { num: 1, b: 1 });
    cell(ws, tr, cP, FX(`SUM(${A1(r1, cP)}:${A1(rL, cP)})`, 100), { num: 1, b: 1 });
    ws.getColumn(1).width = 6; ws.getColumn(2).width = 48; for (let c = 3; c <= cP; c++)ws.getColumn(c).width = 16;

    // ===== GRAFIK ASLI (baru) =====
    const q = `'${name.replace(/'/g, "''")}'!`, rng = (c, a, b) => q + A1(a, c, 1, 1) + ':' + A1(b, c, 1, 1);
    if (def.type === 'bar') {
        const c0 = cP + 2, nt = d.top.length;
        [['No', 0], [def.t5, 1], ['Jumlah (Ha)', 2]].forEach(([v, o]) => { hdr(cell(ws, 1, c0 + o, v)); ws.mergeCells(1, c0 + o, 2, c0 + o); hdr(ws.getCell(2, c0 + o)) });
        d.top.forEach((b, i) => {
            const r = 3 + i, jr = `${A1(r1, cJ, 1, 1)}:${A1(rL, cJ, 1, 1)}`;
            cell(ws, r, c0, i + 1, { ctr: 1 });
            cell(ws, r, c0 + 2, FX(`LARGE(${jr},${i + 1})`, b.t), { num: 1 });                                              // luas terbesar ke-i
            cell(ws, r, c0 + 1, FX(`INDEX(${A1(r1, 2, 1, 1)}:${A1(rL, 2, 1, 1)},MATCH(${A1(r, c0 + 2)},${jr},0))`, b.c));   // nama kategorinya
        });
        ws.getColumn(c0).width = 6; ws.getColumn(c0 + 1).width = 44; ws.getColumn(c0 + 2).width = 16;
        addXChart({ sheet: name, type: 'bar', title: yr(def.ct), legend: null, from: { col: c0 - 1, row: 9 }, ext: { w: 600, h: 340 },
            series: [{ cat: rng(c0 + 1, 3, 2 + nt), catVals: d.top.map(x => x.c), val: rng(c0 + 2, 3, 2 + nt), vals: d.top.map(x => x.t), color: NAVY }] });
    } else {
        addXChart({ sheet: name, type: 'pie', title: yr(def.ct), legend: 'r', from: { col: 0, row: tr + 1 }, ext: { w: 760, h: 340 },
            pointColors: d.body.map((_, i) => PAL[i % PAL.length]),
            series: [{ cat: rng(2, r1, rL), catVals: d.body.map(x => x.c), val: rng(cJ, r1, rL), vals: d.body.map(x => x.t) }] });
    }
}
function sheetExtra(wb, ex, d, name) {
    const ws = wb.addWorksheet(name);
    if (!d) { ws.getCell(1, 1).value = '(Data tidak ditemukan)'; return }
    ex.excel(wb, ws, d, ex);
}
function sheetNew(wb, name) {
    const ws = wb.addWorksheet(name);
    ws.getCell(1, 1).value = NEW_TEXT; ws.getColumn(1).width = 28;
}
async function downloadXLSX(shouldDownload = true) {
    const wb = new ExcelJS.Workbook(); resetXCharts(); wb.calcProperties = { fullCalcOnLoad: true };   // Excel hitung ulang semua rumus saat file dibuka
    // urutan sheet mengikuti SHEETS (sheets.js)
    SHEETS.forEach(s => {
        const name = sheetName(s);
        if (s.kind === 'raw') sheetRaw(wb, name);
        else if (s.kind === 'def') sheetDef(wb, DEFS[s.i], MODELS[s.i], name);
        else if (s.kind === 'extra') sheetExtra(wb, EXTRA[s.i], EXTRA_MODELS[s.i], name);
        else sheetNew(wb, name);
    });
    const buf = await wb.xlsx.writeBuffer(), blob = await injectCharts(buf);   // sisipkan grafik asli Excel
    if (shouldDownload) {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'NPGT_' + (window.NAMA_KAB || 'Jawa_Timur') + '_2026.xlsx'; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    }
    return blob;
}
