// bps-pdf-parse.js - fungsi murni (tanpa DOM) untuk membaca tabel BPS dari teks PDF (hasil pdf.js getTextContent).
// Dipakai bps-pdf.js. Bisa juga diuji di Node (module.exports di bagian bawah).
//
// Masukan utama: "baris" = [{y, items:[{x, s}]}] per halaman, hasil rowsFromItems().
const BPSPDF = (() => {
    // ---------- angka format BPS ----------
    // "586.110" (ribuan) | "0,10" (desimal) | "99.30" (desimal titik) | "–" (kosong)
    const NUM = /^-?(?:\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:,\d+)?|\d+\.\d+)$/;
    const isNum = s => NUM.test(s);
    const parseNum = s => {
        if (/^[–—−-]$/.test(s)) return null;
        if (s.includes(',')) return Number(s.replace(/\./g, '').replace(',', '.'));
        return Number(/^-?\d{1,3}(?:\.\d{3})+$/.test(s) ? s.replace(/\./g, '') : s);
    };
    const norm = s => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');

    // ---------- item pdf.js -> baris ----------
    // Watermark BPS ("https://jatim.bps.go.id") adalah teks MIRING; dibuang lewat matriks transform.
    function rowsFromItems(items) {
        const it = items.filter(i => i.str && i.str.trim() && Math.abs(i.transform[1]) < 0.01 && Math.abs(i.transform[2]) < 0.01)
            .map(i => ({ x: i.transform[4], y: i.transform[5], s: i.str.trim() })).sort((a, b) => b.y - a.y);
        const rows = [];
        for (const t of it) { const r = rows[rows.length - 1]; if (r && Math.abs(r.y - t.y) <= 3.5) r.items.push(t); else rows.push({ y: t.y, items: [t] }) }
        rows.forEach(r => { r.items.sort((a, b) => a.x - b.x); r.text = r.items.map(i => i.s).join(' ') });
        return rows;
    }

    // ---------- nomor halaman cetak (baris paling bawah halaman) ----------
    function printedNo(rows) {
        for (const r of rows.slice(-3).reverse()) {
            for (const it of [r.items[r.items.length - 1], r.items[0]]) {
                const m = it && /^(?:[A-Z][A-Z0-9 ,&\/-]*\s)?(\d{1,4})$/.exec(it.s);
                if (m && r.items.length <= 3) return Number(m[1]);
            }
        }
        return null;
    }

    // ---------- "75-80, 96" -> [75..80, 96] ----------
    function parsePages(spec) {
        const out = [];
        for (const part of String(spec || '').split(/[;,\s]+/).filter(Boolean)) {
            const m = /^(\d+)\s*[-–—]\s*(\d+)$/.exec(part);
            if (m) { const a = +m[1], b = +m[2]; if (b < a || b - a > 40) throw new Error(`Rentang halaman tidak valid: ${part}`); for (let p = a; p <= b; p++) out.push(p) }
            else if (/^\d+$/.test(part)) out.push(+part);
            else throw new Error(`Format halaman tidak valid: "${part}". Contoh: 75-80 atau 96`);
        }
        return [...new Set(out)];
    }

    // ---------- Tabel penduduk per kabupaten/kota (mis. Tabel 3.1.1) ----------
    // Blok dikenali dari baris judul: "Jumlah Penduduk (jiwa)", "Laju Pertumbuhan Penduduk per tahun (%)", "Rasio Jenis Kelamin".
    const yearLabel = s => { let m = /^(20\d\d)\s*[–—-]\s*(20\d\d)\d?$/.exec(s); if (m) return `${m[1]}–${m[2]}`; m = /^(20\d\d)\d?\**$/.exec(s); return m ? m[1] : null };
    function parseAreas(pages) {
        const B = { pop: null, laju: null, rasio: null }, mk = () => ({ labels: [], rows: [], total: null });
        let cur = null, sect = null, pend = null;
        for (const rows of pages) for (const r of rows) {
            const t = norm(r.text);
            if (t === 'jumlahpendudukjiwa' || /^jumlahpenduduk(jiwa)?$/.test(t)) { cur = 'pop'; sect = pend = null; B.pop ??= mk(); continue }
            if (/^lajupertumbuhanpendudukpertahun/.test(t)) { cur = 'laju'; sect = pend = null; B.laju ??= mk(); continue }
            if (t === 'rasiojeniskelamin') { cur = 'rasio'; sect = pend = null; B.rasio ??= mk(); continue }
            if (/^(kepadatan|distribusi)penduduk/.test(t)) { cur = null; pend = null; continue }
            if (!cur) continue;
            const blk = B[cur];
            const yi = r.items.filter(i => !/^\d$/.test(i.s)), ylab = yi.map(i => yearLabel(i.s));   // item 1 digit = angka catatan kaki
            if (yi.length >= 2 && ylab.every(Boolean)) { if (!blk.labels.length) blk.labels = ylab; continue }   // judul kolom: tahun / periode
            const nums = r.items.filter(i => isNum(i.s)), nameItems = r.items.filter(i => !isNum(i.s) && !/^\(\d+\)$/.test(i.s));
            const name = nameItems.map(i => i.s).join(' ').replace(/\s+/g, ' ').trim();
            const need = Math.max(1, blk.labels.length);
            if (!nums.length) {   // penanda bagian atau nama baris yang angkanya ada di baris lain
                if (/^kabupaten/.test(t)) { sect = 'kab'; pend = null; continue }
                if (/^kota(municipality)?$/.test(t) || /^kotamunicipality/.test(t)) { sect = 'kota'; pend = null; continue }
                if (pend && pend.type === 'nums' && name) { blk.rows.push({ sect, name, v: pend.v }); pend = null; continue }   // angka di atas nama
                pend = name ? { type: 'name', name } : null; continue;
            }
            const vals = nums.map(i => parseNum(i.s));
            if (name && vals.length >= need) { blk.rows.push({ sect, name, v: vals }); pend = null; continue }
            if (!name && vals.length >= need) {            // baris angka tanpa nama: pasangkan dengan nama sebelumnya
                if (pend && pend.type === 'name') { blk.rows.push({ sect, name: pend.name, v: vals }); pend = null } else pend = { type: 'nums', v: vals };
            }
        }
        // baris nama yang muncul SETELAH baris angkanya (urutan y terbalik) dipasangkan di sini
        return B;
    }
    // cocokkan dengan nama wilayah di aplikasi ("Kota Kediri" / "Kediri")
    function matchAreas(blk, names) {
        const by = new Map();
        blk.rows.forEach(r => { const k = (r.sect === 'kota' ? 'kota|' : r.sect === 'kab' ? 'kab|' : '?|') + norm(r.name); if (!by.has(k)) by.set(k, r) });
        const out = {}, miss = [];
        for (const n of names) {
            const isKota = /^kota\s/i.test(n), base = norm(n.replace(/^kota\s+/i, ''));
            const r = by.get((isKota ? 'kota|' : 'kab|') + base) || by.get('?|' + base) || (isKota ? null : by.get('kab|' + norm(n)));
            if (r) out[n] = r.v; else miss.push(n);
        }
        const known = new Set(names.map(n => norm(n.replace(/^kota\s+/i, ''))));
        const tot = [...blk.rows].reverse().find(r => !known.has(norm(r.name)) && r.v.length >= Math.max(1, blk.labels.length) && /^(jawa|provinsi|dki|di |banten|bali|sumatera|kalimantan|sulawesi)/i.test(r.name));
        return { values: out, miss, total: tot ? tot.v : null };
    }

    // ---------- Status pekerjaan utama x jenis kelamin (mis. Tabel 3.2.6) ----------
    const PK_DEF = [
        ['Berusaha sendiri', /berusaha\s+sendiri/i],
        ['Buruh tidak dibayar', /berusaha\s+dibantu\s+(?:buruh\s+)?tidak\s+tetap(?:\s*\/?\s*(?:buruh\s+)?tidak\s+dibayar)?|berusaha\s+dibantu\s+buruh\s+tidak\s+dibayar/i],
        ['Buruh dibayar', /berusaha\s+dibantu\s+(?:buruh\s+)?(?:tetap|dibayar)(?:\s*\/?\s*(?:buruh\s+)?(?:tetap|dibayar))?/i],
        ['Buruh/Karyawan/Pegawai', /buruh\s*\/?\s*karyawan\s*\/?\s*pegawai|karyawan\s*\/?\s*pegawai/i],
        ['Pekerja bebas', /pekerja\s+bebas/i],
        ['Pekerja keluarga/tak dibayar', /pekerja\s+keluarga(?:\s*\/?\s*(?:tak|tidak)\s+dibayar)?/i]
    ];
    function parseEmployment(pages) {
        const text = pages.flat().map(r => r.text).join(' ').replace(/\s+/g, ' ');
        const hits = PK_DEF.map(([name, re]) => { const m = re.exec(text); return m ? { name, i: m.index, e: m.index + m[0].length } : null });
        if (hits.some(h => !h)) return { rows: [], err: 'Label status pekerjaan tidak lengkap: ' + PK_DEF.filter((_, k) => !hits[k]).map(d => d[0]).join(', ') };
        const order = [...hits].sort((a, b) => a.i - b.i), toks = s => s.split(/\s+/).filter(isNum).map(parseNum);
        const rows = hits.map(h => { const nx = order[order.indexOf(h) + 1], v = toks(text.slice(h.e, nx ? nx.i : text.length)); return { name: h.name, l: v[0], p: v[1], t: v[2] } });
        const last = order[order.length - 1], tm = /jumlah\s*\/?\s*total/i.exec(text.slice(last.e)), tv = tm ? toks(text.slice(last.e + tm.index + tm[0].length)) : [];
        return { rows, total: tv.length >= 2 ? { l: tv[0], p: tv[1], t: tv[2] } : null };
    }

    // ---------- PDRB menurut lapangan usaha (Tabel 13.1.1 / 13.1.2) ----------
    const PD_CODES = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M,N', 'O', 'P', 'Q', 'R,S,T,U'];
    const CODE_RE = /^(M\s*,\s*N|R\s*,\s*S\s*,\s*T\s*,\s*U|[A-U])(?=\s|$)/;
    // want: 'konstan' (ADHK) | 'berlaku' (ADHB) | undefined. Jenis tabel dikenali dari judul ("...Harga Konstan 2010..." / "...Harga Berlaku...");
    // halaman lanjutan tanpa judul mengikuti jenis halaman sebelumnya. Baris dari tabel jenis lain diabaikan.
    function parsePdrb(pages, want) {
        const got = {}, dup = [], seen = []; let years = [], pend = null, span = null, seg = null, skipped = 0;
        for (const rows of pages) for (const [ri, r] of rows.entries()) {
            if (ri < 14) { const km = /harga\s*(konstan|berlaku)|(constant|current)\s+market\s+prices/i.exec(r.text); if (km) { seg = /konstan|constant/i.test(km[0]) ? 'konstan' : 'berlaku'; pend = null; if (!seen.includes(seg)) seen.push(seg) } }
            if (want && seg && seg !== want) { if (CODE_RE.test(r.items[0].s)) skipped++; continue }
            if (!years.length && !span) { const sm = /(20\d\d)\s*[–—-]\s*(20\d\d)/.exec(r.text); if (sm && +sm[2] - +sm[1] === 4) span = [+sm[1], +sm[2]] }   // judul tabel: "..., 2021–2025"
            const yl = r.items.map(i => /^(20\d\d)\*{0,2}$/.exec(i.s)).filter(Boolean);
            if (yl.length >= 3 && yl.length === r.items.length) { if (!years.length) years = yl.map(m => +m[1]); continue }
            const first = r.items[0], m = CODE_RE.exec(first.s), nums = r.items.filter(i => /^-?\d{1,3}(?:\.\d{3})*,\d{2}$/.test(i.s)).map(i => parseNum(i.s));
            if (m && nums.length >= 5) {
                const c = m[1].replace(/\s/g, ''); if (got[c]) dup.push(c); else got[c] = nums.slice(-5); pend = null;
            } else if (m && !nums.length) pend = m[1].replace(/\s/g, '');
            else if (!m && nums.length >= 5 && pend && r.items.every(i => isNum(i.s) || /,\d\d$/.test(i.s))) { if (!got[pend]) got[pend] = nums.slice(-5); pend = null }
        }
        if (!years.length && span) years = [0, 1, 2, 3, 4].map(k => span[0] + k);
        return { years, rows: PD_CODES.map(c => got[c] || null), missing: PD_CODES.filter(c => !got[c]), dup, seen, skipped };
    }

    return { NUM, isNum, parseNum, norm, rowsFromItems, printedNo, parsePages, parseAreas, matchAreas, parseEmployment, parsePdrb, PD_CODES, PK_DEF };
})();
if (typeof module !== 'undefined') module.exports = BPSPDF;
