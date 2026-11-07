// bps-parse.js - fungsi murni (tanpa DOM): teks hasil OCR -> baris + angka. Dipakai bps-image.js.
const BPS = (() => {
    // Angka format BPS: "58,8" | "112.009" | "15.205,35" | "15 907.7" (spasi ribuan + titik desimal) | "102," (terpotong)
    const NUM_SRC = String.raw`(?<![\w.,])-?(?:\d{1,3}(?: \d{3})+[.,]\d+|\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+[.,]\d+|\d+,(?!\d)|\d+)(?![\w])`;
    const stripNoise = s => s.replace(/\(\s*\d+\s*\)/g, ' ').replace(/[|_]/g, ' ');   // "(1) (2)" = nomor kolom BPS
    function parseNum(raw) {
        let s = raw.replace(/\s/g, ''), warn = '';
        if (/,$/.test(s)) { s = s.slice(0, -1); warn = 'Angka terpotong di gambar (diakhiri koma). Periksa dan lengkapi.' }
        const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
        if (c >= 0 && d >= 0) s = c > d ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
        else if (c >= 0) s = s.replace(',', '.');
        else if (d >= 0 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');   // 112.009 = ribuan
        const v = parseFloat(s), m = /^-?(\d+)(?:\.(\d+))?$/.exec(s) || [], fr = m[2] || '';
        const bad = fr.length >= 4;                       // mis. "15.20535": koma desimal hilang/salah -> dianggap tanpa tanda desimal
        return { v: isNaN(v) ? null : v, raw, warn, sep: fr.length > 0 && !bad, dec: fr.length, bad, digits: (m[1] || '') + fr, sign: s.startsWith('-') ? -1 : 1 };
    }
    // Jumlah desimal tiap kolom = modus dari angka yang punya tanda desimal (jika >=50% angka kolom punya tanda desimal; selain itu 0)
    function detectDecimals(rows) {
        const nc = Math.max(0, ...rows.map(r => r.nums.length)), d = [], share = [];
        for (let c = 0; c < nc; c++) {
            const col = rows.map(r => r.nums[c]).filter(Boolean), sep = col.filter(n => n.sep), cnt = {};
            sep.forEach(n => cnt[n.dec] = (cnt[n.dec] || 0) + 1);
            const top = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0];
            if (col.length >= 4 && top && sep.length >= col.length * 0.5) { d.push(+top[0]); share.push(top[1] / sep.length) } else { d.push(0); share.push(0) }
        }
        return { d, share };
    }
    // Hitung n.val dari n.v: kembalikan tanda desimal yang hilang (588 -> 58,8). rowwise: pilih jumlah desimal yang membuat nilainya paling dekat dgn angka lain di baris yg sama (deret tahun)
    function applyDec(rows, det, rowwise) {
        rows.forEach(r => {
            const good = r.nums.filter(n => n.sep && n.v > 0).map(n => n.v).sort((a, b) => a - b), med = good.length >= 2 ? good[good.length >> 1] : null;
            r.nums.forEach((n, c) => {
                if (n.edited) return;
                const D = det.d[c] || 0; n.val = n.v; n.fixed = false; n.flag = n.warn || '';
                if (D > 0 && (!n.sep || n.bad) && n.digits.length > D) {
                    let best = D;
                    if (rowwise && med) { let bd = 1e9; [D, D - 1, D + 1].forEach(x => { if (x < 0 || n.digits.length <= x) return; const e = Math.abs(Math.log10(+n.digits / 10 ** x / med)); if (e < bd) { bd = e; best = x } }) }
                    n.val = n.sign * +n.digits / 10 ** best; n.fixed = true; n.flag = n.flag || `Tanda desimal hilang di gambar, dikembalikan otomatis (${best} desimal). Periksa.`;
                } else if (D > 0 && n.sep && n.dec !== D && (det.share[c] || 0) >= 0.7) n.flag = n.flag || `Jumlah desimal berbeda dari kolom (${D}). Periksa.`;
            });
        });
    }
    function parseText(text) {
        const rows = []; let years = [];
        for (const line0 of String(text).split(/\r?\n/)) {
            const c = stripNoise(line0.trim()); if (!c.trim()) continue;
            const toks = [...c.matchAll(new RegExp(NUM_SRC, 'g'))]; if (!toks.length) continue;
            const t = toks.map(m => m[0]);
            if (t.length >= 2 && t.every(x => /^20\d\d$/.test(x))) { if (!years.length) years = t.map(Number); continue }   // baris judul tahun
            rows.push({ label: c.slice(0, toks[0].index).replace(/\s+/g, ' ').trim(), nums: t.map(parseNum) });
        }
        return { rows, years };
    }
    // ---- pencocokan nama kabupaten/kota ----
    const norm = s => String(s).toLowerCase().replace(/\bkab(upaten|\.)?\b\.?/g, ' ').replace(/[^a-z]/g, '');
    function lev(a, b) {
        const m = a.length, n = b.length, d = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);
        for (let j = 1; j <= n; j++) d[0][j] = j;
        for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        return d[m][n];
    }
    function matchName(label, names) {
        const x = norm(label); if (x.length < 3) return null;
        const ex = names.find(n => norm(n) === x); if (ex) return ex;
        const sc = names.map(n => [lev(x, norm(n)), n]).sort((a, b) => a[0] - b[0]);
        const lim = Math.max(1, Math.floor(x.length * 0.2));
        return sc[0][0] <= lim && (!sc[1] || sc[1][0] > sc[0][0]) ? sc[0][1] : null;   // toleransi salah baca OCR, hanya jika pemenang jelas
    }
    // ---- kode lapangan usaha (A, B, "M, N", ...) ----
    const codeOf = label => { const m = /^\s*([A-U](?:\s*,\s*[A-U])*)(?![A-Za-z])/.exec(label); return m ? m[1].replace(/\s/g, '') : '' };
    // ---- tebak jenis data dari teks OCR ----
    function guess(text) {
        const t = String(text).toLowerCase().replace(/[^a-z]/g, '');
        if (/rasiojeniskelamin|sexratio/.test(t)) return 'rasio';
        if (/statuspekerjaan|berusahasendiri|ownaccount/.test(t)) return 'pk';
        if (/lapanganusaha|pertaniankehutanan|industry/.test(t)) return 'pd';
        if (/jumlahpenduduk|lajupertumbuhan|population/.test(t)) return /jumlahpenduduk/.test(t) || !/laju/.test(t) ? 'pop' : 'laju';
        return 'pop';
    }
    const unitOf = text => /ribu|thousand/i.test(text) ? 1000 : 1;
    return { parseNum, parseText, matchName, codeOf, guess, unitOf, detectDecimals, applyDec, NUM_SRC };
})();
