// bps-pdf.js - Langkah 4: isi Tabel III-1 s.d. III-5 langsung dari PDF BPS.
// User cukup menyeret PDF BPS lalu mengetik nomor halaman tiap tabel. Dimuat setelah main.js, ui.js, extra.js.
// Pembaca tabel (tanpa DOM) ada di bps-pdf-parse.js. Pustaka PDF: pdf.js (index.html).
(() => {
    const P = BPSPDF, st = msg => { $('status').textContent = msg };
    const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const fi = v => Math.round(v).toLocaleString('id');
    if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const S = { doc: null, name: '', rows: new Map(), printed: new Map(), res: null, sel: { pop: -1, laju: -1, rasio: -1 } };
    const FIELDS = { pop: 'pg-pop', pk: 'pg-pk', adhk: 'pg-adhk', adhb: 'pg-adhb' };
    try { const v = JSON.parse(localStorage.getItem('djpa_bps_pages') || '{}'); Object.entries(FIELDS).forEach(([k, id]) => { if (v[k]) $(id).value = v[k] }); if (v.mode) document.querySelector(`input[name=pgmode][value=${v.mode}]`).checked = true } catch (e) { }
    const mode = () => document.querySelector('input[name=pgmode]:checked').value;   // 'cetak' | 'pdf'
    const saveForm = () => { try { const v = { mode: mode() }; Object.entries(FIELDS).forEach(([k, id]) => v[k] = $(id).value); localStorage.setItem('djpa_bps_pages', JSON.stringify(v)) } catch (e) { } };

    // ---------- baca halaman PDF ----------
    async function rowsOf(n) {
        if (!S.rows.has(n)) { const pg = await S.doc.getPage(n), c = await pg.getTextContent(); S.rows.set(n, P.rowsFromItems(c.items)) }
        return S.rows.get(n);
    }
    async function printedOf(n) { if (!S.printed.has(n)) S.printed.set(n, P.printedNo(await rowsOf(n))); return S.printed.get(n) }
    // nomor halaman CETAK (yang tertulis di kaki halaman / daftar isi) -> nomor halaman PDF. Pencarian biner: nomor cetak naik seiring halaman PDF.
    async function pdfPageOfPrinted(target) {
        const N = S.doc.numPages, est = async i => { for (const d of [0, 1, -1, 2, -2, 3, -3, 4, -4, 5, -5]) { const j = i + d; if (j < 1 || j > N) continue; const p = await printedOf(j); if (p != null) return p - d } return null };
        let lo = 1, hi = N;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1, e = await est(mid);
            if (e === null) { lo = mid + 1; continue }          // halaman tanpa nomor (sampul/romawi): anggap di depan
            if (e === target) { for (const d of [0, 1, -1, 2, -2, 3, -3, 4, -4]) { const j = mid + d; if (j >= 1 && j <= N && await printedOf(j) === target) return j } return mid }
            if (e < target) lo = mid + 1; else hi = mid - 1;
        }
        throw new Error(`Halaman cetak ${target} tidak ditemukan di PDF. Periksa nomornya, atau pilih "Nomor halaman PDF".`);
    }
    async function pagesFor(spec) {
        const want = P.parsePages(spec), N = S.doc.numPages;
        if (!want.length) return [];
        if (mode() === 'pdf') { want.forEach(p => { if (p < 1 || p > N) throw new Error(`Halaman PDF ${p} di luar rentang 1-${N}.`) }); return want }
        const out = []; let prev = null;
        for (const t of want) {
            let j = null;
            if (prev) { const g = prev.j + (t - prev.t); if (g >= 1 && g <= N && await printedOf(g) === t) j = g }   // halaman berurutan: cukup cek tetangga
            if (j === null) j = await pdfPageOfPrinted(t);
            out.push(j); prev = { t, j };
        }
        return out;
    }
    const pagesRows = async arr => Promise.all(arr.map(rowsOf));

    // ---------- baca semua tabel yang diisi ----------
    async function readAll() {
        const names = [...new Set(ROWS.map(kecOf))].filter(x => x !== '-').sort(), R = { names, notes: [] };
        const spec = k => $(FIELDS[k]).value.trim(), note = (k, kind, title, msg, sub) => R.notes.push({ k, kind, title, msg, sub });
        if (spec('pop')) {
            const pg = await pagesFor(spec('pop')), B = P.parseAreas(await pagesRows(pg)); R.pop = {};
            for (const [k, label] of [['pop', 'Jumlah penduduk'], ['laju', 'Laju pertumbuhan'], ['rasio', 'Rasio jenis kelamin']]) {
                if (!B[k]) { note(k, 'warn', label, 'Blok tabel tidak ditemukan di halaman yang diberikan.', 'Pastikan rentang halaman mencakup judul "' + (k === 'pop' ? 'Jumlah Penduduk (jiwa)' : k === 'laju' ? 'Laju Pertumbuhan Penduduk per tahun (%)' : 'Rasio Jenis Kelamin') + '".'); continue }
                const m = P.matchAreas(B[k], names); R.pop[k] = { labels: B[k].labels, ...m };
                if (S.sel[k] < 0 || S.sel[k] >= B[k].labels.length) S.sel[k] = Math.max(0, B[k].labels.length - 1);
            }
        }
        if (spec('pk')) { const E = P.parseEmployment(await pagesRows(await pagesFor(spec('pk')))); R.pk = E }
        for (const k of ['adhk', 'adhb']) if (spec(k)) R[k] = P.parsePdrb(await pagesRows(await pagesFor(spec(k))), k === 'adhk' ? 'konstan' : 'berlaku');
        if (!R.pop && !R.pk && !R.adhk && !R.adhb) throw new Error('Isi minimal satu nomor halaman tabel.');
        return R;
    }

    // ---------- terapkan ke tabel aplikasi + laporan ----------
    function applyAndReport(R) {
        const out = [], add = (kind, title, msg, sub) => out.push({ kind, title, msg, sub });
        let filled = 0;
        if (R.pop) {
            const L = { pop: 'Jumlah penduduk', laju: 'Laju pertumbuhan penduduk', rasio: 'Rasio jenis kelamin' };
            for (const k of ['pop', 'laju', 'rasio']) {
                const d = R.pop[k]; if (!d) { add('warn', `Tabel III-${k === 'pop' ? 1 : 2} · ${L[k]}`, 'Blok tidak ditemukan di halaman yang diberikan.'); continue }
                const c = S.sel[k]; let n = 0;
                R.names.forEach(nm => { const v = d.values[nm]; if (!v || v[c] === null || v[c] === undefined) return; if (k === 'pop') popSet(nm, Math.round(v[c])); else t2Set(k, nm, v[c]); n++ });
                filled += n;
                const lab = d.labels[c] || `kolom ${c + 1}`;
                let kind = d.miss.length ? 'warn' : 'ok', sub = d.miss.length ? 'Tidak ditemukan: ' + d.miss.join(', ') : '';
                if (k === 'pop' && d.total && d.total[c] != null) {
                    const sum = R.names.reduce((s, nm) => s + ((d.values[nm] || [])[c] || 0), 0);
                    sub += (sub ? ' · ' : '') + (Math.round(sum) === Math.round(d.total[c]) ? `Jumlah ${fi(sum)} = total provinsi di PDF ✓` : `Jumlah ${fi(sum)} ≠ total provinsi di PDF ${fi(d.total[c])} (selisih ${fi(sum - d.total[c])})`);
                    if (Math.round(sum) !== Math.round(d.total[c]) && !d.miss.length) kind = 'warn';
                }
                add(kind, `Tabel III-${k === 'pop' ? 1 : 2} · ${L[k]} (${lab})`, `${n} dari ${R.names.length} kabupaten/kota terisi.`, sub);
            }
        }
        if (R.pk) {
            if (R.pk.err) add('bad', 'Tabel III-3 · Status pekerjaan utama', R.pk.err);
            else {
                R.pk.rows.forEach((r, i) => { if (r.l != null) pkSet(i, 'l', r.l); if (r.p != null) pkSet(i, 'p', r.p) }); filled += R.pk.rows.length * 2;
                const badRow = R.pk.rows.filter(r => r.t != null && Math.abs(r.l + r.p - r.t) > 1), sl = R.pk.rows.reduce((s, r) => s + r.l, 0), sp = R.pk.rows.reduce((s, r) => s + r.p, 0), T = R.pk.total;
                let sub = badRow.length ? `Laki-laki + perempuan ≠ jumlah pada: ${badRow.map(r => r.name).join(', ')}` : 'Laki-laki + perempuan = jumlah di tiap baris ✓';
                let kind = badRow.length ? 'warn' : 'ok';
                if (T) { const okT = Math.round(sl) === Math.round(T.l) && Math.round(sp) === Math.round(T.p); sub += okT ? ' · Total sesuai baris Jumlah di PDF ✓' : ` · Total (${fi(sl)} / ${fi(sp)}) ≠ baris Jumlah di PDF (${fi(T.l)} / ${fi(T.p)})`; if (!okT) kind = 'warn' }
                add(kind, 'Tabel III-3 · Status pekerjaan utama', `${R.pk.rows.length} dari ${P.PK_DEF.length} status terisi.`, sub);
            }
        }
        for (const [k, no, nama, set] of [['adhk', 4, 'PDRB atas dasar harga konstan (ADHK)', pdSet], ['adhb', 5, 'PDRB atas dasar harga berlaku (ADHB)', pd2Set]]) {
            const d = R[k]; if (!d) continue;
            let years = d.years, sub = '', kind = 'ok';
            const want = k === 'adhk' ? 'konstan' : 'berlaku', other = k === 'adhk' ? 'ADHB (harga berlaku)' : 'ADHK (harga konstan)';
            if (d.seen.length && !d.seen.includes(want)) { add('bad', `Tabel III-${no} · ${nama}`, `Halaman yang diberikan berisi tabel ${other}, bukan ${k.toUpperCase()}. Tidak ada data yang diisi.`, 'Periksa nomor halaman tabel yang benar di daftar isi.'); continue }
            if (d.skipped) { kind = 'warn'; sub = `Rentang halaman juga memuat tabel ${other}; bagian itu diabaikan.` }
            if (!years.length) { add('bad', `Tabel III-${no} · ${nama}`, 'Tahun kolom tidak terbaca dari judul tabel. Pastikan halaman pertama tabel ikut dimasukkan.'); continue }
            if (years.length === PD_TAHUN.length && years.some((y, i) => y !== PD_TAHUN[i]) && !R._yearsSet) { PD_TAHUN.splice(0, PD_TAHUN.length, ...years); R._yearsSet = true; sub += (sub ? ' · ' : '') + `Tahun kolom disesuaikan ke ${years[0]}–${years[years.length - 1]}.` }
            let n = 0;
            PD_USAHA.forEach((u, i) => { const row = d.rows[i]; if (!row) return; PD_TAHUN.forEach((y, yi) => { const c = years.indexOf(y); if (c >= 0 && row[c] != null) { set(i, yi, row[c]); n++ } }) });
            filled += n;
            const miss = PD_TAHUN.filter(y => !years.includes(y));
            if (miss.length) { kind = 'warn'; sub += (sub ? ' · ' : '') + `Tahun ${miss.join(', ')} tidak ada di tabel ini (tahun di PDF: ${years.join(', ')}).` }
            if (d.missing.length) { kind = 'warn'; sub += (sub ? ' · ' : '') + `Lapangan usaha tidak ditemukan: ${d.missing.join(', ')}.` }
            if (d.dup.length) { kind = 'warn'; sub += (sub ? ' · ' : '') + `Kode ${[...new Set(d.dup)].join(', ')} muncul lebih dari sekali; yang pertama dipakai. Cek apakah rentang halaman memuat tabel lain.` }
            add(kind, `Tabel III-${no} · ${nama}`, `${PD_USAHA.length - d.missing.length} dari ${PD_USAHA.length} lapangan usaha, ${years[0]}–${years[years.length - 1]} (${n} nilai).`, sub);
        }
        const on = document.querySelector('.tab.on'), idx = on ? on.dataset.i : null;
        render(); if (idx !== null) document.querySelector(`.tab[data-i="${idx}"]`)?.click();
        $('dl').disabled = false;
        return { out, filled };
    }

    function showReport(out, R) {
        const box = $('bps-review'); box.hidden = false;
        const sels = ['pop', 'laju', 'rasio'].filter(k => R.pop && R.pop[k] && R.pop[k].labels.length > 1).map(k =>
            `<label>${{ pop: 'Penduduk (III-1, III-2)', laju: 'Laju pertumbuhan (III-2)', rasio: 'Rasio jenis kelamin (III-2)' }[k]}<select data-sel="${k}">${R.pop[k].labels.map((l, i) => `<option value="${i}" ${i === S.sel[k] ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select></label>`).join('');
        box.innerHTML = out.map(o => `<div class="rp ${o.kind}"><b>${esc(o.title)}</b> · ${esc(o.msg)}${o.sub ? `<small>${esc(o.sub)}</small>` : ''}</div>`).join('') + (sels ? `<div class="rp-sel">${sels}<small class="rp-hint">Pilih kolom/tahun yang dipakai bila tabel BPS memuat beberapa tahun.</small></div>` : '');
    }

    async function run() {
        if (!S.doc) { st('Unggah file PDF BPS dulu (kotak di langkah 4).'); return }
        if (!ROWS.length) { st('Unggah file utama (Excel) dulu, baru isi tabel dari PDF BPS.'); return }
        saveForm(); const btn = $('bps-run'); btn.disabled = true;
        Busy.show('Membaca tabel dari PDF BPS', ['Mencari halaman', 'Membaca angka tabel', 'Mengisi Tabel III-1 s.d. III-5']); await tick();
        try {
            Busy.stage(1); Busy.detail('Membaca halaman PDF…');
            const R = await readAll(); S.res = R;
            Busy.stage(2); await tick();
            const { out, filled } = applyAndReport(R); showReport(out, R);
            const warn = out.filter(o => o.kind !== 'ok').length;
            st(`PDF BPS: ${filled.toLocaleString('id')} nilai terisi ke ${out.length} tabel` + (warn ? `. Peringatan: ${warn} tabel perlu dicek (lihat rincian di langkah 4).` : '. Semua pemeriksaan lolos.'));
        } catch (e) { st('Gagal membaca PDF BPS: ' + e.message) } finally { Busy.hide(); btn.disabled = false }
    }

    // ---------- unggah PDF ----------
    async function loadPdf(file) {
        const dz = $('drop3');
        if (!window.pdfjsLib) { st('Gagal membaca PDF: pustaka pdf.js belum termuat (perlu koneksi internet).'); UI.drop(dz, 'error', file.name, 'pdf.js tidak termuat.'); return }
        st('Membaca ' + file.name + '…'); Busy.show('Membuka ' + file.name, []);
        try {
            S.doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise; S.name = file.name; S.rows.clear(); S.printed.clear();
            UI.drop(dz, 'done', file.name, `${S.doc.numPages.toLocaleString('id')} halaman. Isi nomor halaman tabel di bawah, lalu klik "Baca & isi tabel".`);
            st(`${file.name}: ${S.doc.numPages.toLocaleString('id')} halaman terbuka. Isi nomor halaman tabel, lalu klik "Baca & isi tabel".`);
            $('bps-run').disabled = false;
        } catch (e) { S.doc = null; UI.drop(dz, 'error', file.name, 'Bukan PDF yang valid / tidak bisa dibuka.'); st('Gagal membaca PDF: ' + e.message) } finally { Busy.hide() }
    }
    const dz3 = $('drop3'), f3 = $('bps-pdf');
    dz3.onclick = () => f3.click(); f3.onchange = () => { f3.files[0] && loadPdf(f3.files[0]); f3.value = '' };
    dz3.addEventListener('drop', e => { const f = [...e.dataTransfer.files].find(x => /\.pdf$/i.test(x.name)); f ? loadPdf(f) : st('Gagal membaca PDF: file harus berformat .pdf.') });
    $('bps-run').onclick = run;
    document.querySelectorAll('.pgform input').forEach(i => i.addEventListener('keydown', e => { if (e.key === 'Enter') run() }));
    $('bps-review').addEventListener('change', e => {
        const t = e.target; if (!t.matches('select[data-sel]') || !S.res) return;
        S.sel[t.dataset.sel] = +t.value; const { out } = applyAndReport(S.res); showReport(out, S.res);
        st(`PDF BPS: kolom ${t.dataset.sel} diganti ke "${t.options[t.selectedIndex].text}" dan tabel diperbarui.`);
    });
    window.BPSPDFUI = { S, run, loadPdf };   // dipakai pengujian
})();
