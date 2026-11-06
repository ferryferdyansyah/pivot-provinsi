// xlsx-stream.js - membaca .xlsx BESAR (ratusan MB XML) secara streaming, tanpa memuat semuanya ke memori.
// Memakai DecompressionStream bawaan browser. Hasil: {cols, rows:[array], total(baris sumber)}.
async function readXlsxStream(file, onProgress, opt = {}) {
    // opt.sum: kolom angka yang dijumlahkan | opt.keep: hanya kolom ini yang dipakai (urutan asli) | opt.drop: kolom yang dibuang | opt.unit: kolom wilayah analisis (menimpa WADMKC)
    // Baris dengan atribut yang sama (per wilayah) digabung -> Raw Data jauh lebih kecil, total luas tetap sama.
    const u16 = (b, o) => b[o] | b[o + 1] << 8, u32 = (b, o) => (b[o] | b[o + 1] << 8 | b[o + 2] << 16 | b[o + 3] << 24) >>> 0;
    const tail = new Uint8Array(await file.slice(Math.max(0, file.size - 70000)).arrayBuffer());
    let e = tail.length - 22; while (e >= 0 && u32(tail, e) !== 0x06054b50) e--;
    if (e < 0) throw new Error('Bukan file .xlsx yang valid');
    const n = u16(tail, e + 10); let p = u32(tail, e + 16) - (file.size - tail.length), ent = null;
    const dec = new TextDecoder();
    for (let i = 0; i < n; i++) {
        const nl = u16(tail, p + 28), el = u16(tail, p + 30), cl = u16(tail, p + 32), nm = dec.decode(tail.subarray(p + 46, p + 46 + nl));
        if (/^xl\/worksheets\/sheet\d+\.xml$/.test(nm) && (!ent || nm < ent.nm)) ent = { nm, m: u16(tail, p + 10), cs: u32(tail, p + 20), off: u32(tail, p + 42) };
        p += 46 + nl + el + cl;
    }
    if (!ent) throw new Error('Sheet data tidak ditemukan di dalam xlsx');
    const lh = new Uint8Array(await file.slice(ent.off, ent.off + 30).arrayBuffer());
    const ds = ent.off + 30 + u16(lh, 26) + u16(lh, 28);
    let stream = file.slice(ds, ds + ent.cs).stream();
    if (ent.m === 8) stream = stream.pipeThrough(new DecompressionStream('deflate-raw'));
    const reader = stream.pipeThrough(new TextDecoderStream()).getReader();

    const cellRe = /<c r="([A-Z]+)\d+"[^>]*?(?:\/>|>([\s\S]*?)<\/c>)/g, valRe = /<(?:t|v)[^>]*>([\s\S]*?)<\/(?:t|v)>/;
    const unesc = s => s.indexOf('&') < 0 ? s : s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
    const colNum = s => { let x = 0; for (let i = 0; i < s.length; i++) x = x * 26 + s.charCodeAt(i) - 64; return x - 1 };
    let cols = null, ik = -1, buf = '', total = 0, pool = [], keep = [], sumI = new Set(), iu = -1, iW = -1;
    const grp = new Map();
    const parseRow = r => {
        const a = []; let m; cellRe.lastIndex = 0;
        while ((m = cellRe.exec(r))) { if (!m[2]) continue; const v = valRe.exec(m[2]); if (v) a[colNum(m[1])] = m[2].charCodeAt(1) === 118 ? String(+v[1]) : unesc(v[1]).trim() }   // angka disalin ulang agar tidak menahan teks XML besar di memori
        return a;
    };
    for (; ;) {
        const { done, value } = await reader.read(); if (done) break;
        buf += value; const parts = buf.split('</row>'); buf = parts.pop();
        for (const r of parts) {
            const a = parseRow(r);
            if (!cols) {
                cols = []; for (let i = 0; i < a.length; i++) cols.push(a[i] || ''); ik = cols.indexOf('WADMKK'); pool = cols.map(() => new Map());
                const drop = new Set(opt.drop || []), kp = opt.keep ? new Set(opt.keep) : null; cols.forEach((c, i) => { if (!drop.has(c) && (!kp || kp.has(c))) keep.push(i); if ((opt.sum || []).includes(c)) sumI.add(i) });
                iu = opt.unit ? cols.indexOf(opt.unit) : -1; iW = cols.indexOf('WADMKC'); continue
            }
            const kb = a[ik]; if (!kb) continue; total++;
            if (iu >= 0 && iW >= 0) a[iW] = a[iu];          // kolom wilayah analisis = kabupaten/kota
            let key = ''; for (const i of keep) if (!sumI.has(i)) key += (a[i] || '') + '\u0001';
            const g = grp.get(key);
            if (g) { for (const i of sumI) g[i] += parseFloat(a[i]) || 0; continue }
            const row = new Array(cols.length);   // nilai berulang dipakai bersama agar hemat memori
            for (const i of keep) { let v = a[i]; if (sumI.has(i)) { row[i] = parseFloat(v) || 0; continue } if (v === undefined) v = ''; else if (v.length < 400) { const pm = pool[i], s = pm.get(v); if (s === undefined) pm.set(v, v); else v = s } row[i] = v }
            grp.set(key, row);
        }
        if (onProgress) onProgress(total);
    }
    if (!cols || ik < 0) throw new Error('Kolom WADMKK (kabupaten/kota) tidak ada di baris judul');
    const outCols = keep.map(i => cols[i]), rows = [...grp.values()].map(r => keep.map(i => sumI.has(i) ? String(r[i]) : r[i]));
    return { cols: outCols, rows, total };
}
