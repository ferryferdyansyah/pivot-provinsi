// bps-image.js - Isi Tabel III-1..III-5 dari gambar tabel BPS: OCR (Tesseract.js, di browser) -> pratinjau bisa diedit -> pilih tabel tujuan -> terapkan.
// Dimuat setelah extra.js & ui.js. Fungsi tanpa DOM ada di bps-parse.js.
(() => {
    const DEST = {
        pop: { t: 'Tabel III-1 · Jumlah penduduk (jiwa)', k: 'kab', tab: '01_' },
        laju: { t: 'Tabel III-2 · Laju pertumbuhan penduduk (%)', k: 'kab', tab: '02_' },
        rasio: { t: 'Tabel III-2 · Rasio jenis kelamin', k: 'kab', tab: '02_' },
        pk: { t: 'Tabel III-3 · Status pekerjaan utama', k: 'pk', tab: '03_' },
        pd: { t: 'Tabel III-4 · PDRB per lapangan usaha', k: 'pd', tab: '04_' },
        pd2: { t: 'Tabel III-5 · PDRB per lapangan usaha (set kedua)', k: 'pd', tab: '05_' }
    };
    const S = { url: '', file: null, text: '', dest: 'pop', col: 0, mult: 1, rows: [], years: [], det: { d: [], share: [] }, names: [] };
    const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const fmtV = v => v === null || v === undefined || isNaN(v) ? '' : String(+(+v).toFixed(4));
    const st = msg => { $('status').textContent = msg };

    // ---------- OCR ----------
    let worker = null;
    const loadTess = () => window.Tesseract ? Promise.resolve() : new Promise((ok, no) => {
        const s = document.createElement('script'); s.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
        s.onload = ok; s.onerror = () => no(new Error('Tesseract.js gagal dimuat (perlu koneksi internet).')); document.head.append(s);
    });
    function prep(img, sharpen) {          // perbesar (tanda koma/titik kecil ikut terbaca) + ubah ke hitam-putih (watermark pucat hilang)
        const sc = Math.min(3, Math.max(1, 2000 / img.naturalWidth)), w = Math.round(img.naturalWidth * sc), h = Math.round(img.naturalHeight * sc);
        const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
        const cx = cv.getContext('2d', { willReadFrequently: true }); cx.imageSmoothingQuality = 'high'; cx.drawImage(img, 0, 0, w, h);
        if (!sharpen) return cv;
        const id = cx.getImageData(0, 0, w, h), d = id.data, hist = new Array(256).fill(0), N = w * h;
        for (let i = 0; i < N; i++) { const v = Math.round(.299 * d[i * 4] + .587 * d[i * 4 + 1] + .114 * d[i * 4 + 2]); d[i * 4] = v; hist[v]++ }
        let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i];
        let sB = 0, wB = 0, mx = 0, th = 128;   // ambang Otsu
        for (let i = 0; i < 256; i++) { wB += hist[i]; if (!wB) continue; const wF = N - wB; if (!wF) break; sB += i * hist[i]; const b = wB * wF * (sB / wB - (sum - sB) / wF) ** 2; if (b > mx) { mx = b; th = i } }
        for (let i = 0; i < N; i++) { const v = d[i * 4] > th ? 255 : 0; d[i * 4] = d[i * 4 + 1] = d[i * 4 + 2] = v; d[i * 4 + 3] = 255 }
        cx.putImageData(id, 0, 0); return cv;
    }
    async function runOcr() {
        if (!ROWS.length) { st('Unggah file utama (Excel) dulu, baru isi data BPS dari gambar.'); return }
        if (!S.url) return;
        Busy.show('Membaca gambar tabel BPS', ['Menyiapkan gambar', 'Membaca teks (OCR)', 'Menyusun pratinjau']);
        try {
            Busy.detail('Memuat mesin OCR… (pertama kali agak lama, ±10 MB)'); await loadTess();
            const img = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error('Gambar tidak bisa dibuka.')); i.src = S.url });
            const cv = prep(img, $('bps-sharp') ? $('bps-sharp').checked : true);
            Busy.stage(1); Busy.detail('Membaca teks…');
            if (!worker) {
                worker = await Tesseract.createWorker('eng', 1, { logger: m => { if (m.status === 'recognizing text') { Busy.progress(m.progress * 100); Busy.detail(`Membaca teks… ${Math.round(m.progress * 100)}%`) } else Busy.detail(m.status) } });
                await worker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
            }
            const { data } = await worker.recognize(cv);
            Busy.stage(2); await UI.tick();
            setText(data.text); st(`Gambar terbaca: ${S.rows.length} baris. Periksa angkanya, pilih tabel tujuan, lalu klik Terapkan.`);
        } catch (e) { st('Gagal membaca gambar: ' + e.message) } finally { Busy.hide() }
    }

    // ---------- Susun baris dari teks OCR ----------
    function setText(text) {
        S.text = text; S.dest = BPS.guess(text); S.mult = BPS.unitOf(text); S.col = 0;
        const one = BPS.parseText(text).rows.filter(r => r.nums.length === 1 && /[A-Za-z]{3}/.test(r.label)).map(r => r.nums[0].v);
        if (S.dest === 'pop' && one.length >= 5 && one.filter(x => x > 80 && x < 130).length >= one.length * 0.7) S.dest = 'rasio';   // judul tak terbaca: nilai 80-130 = rasio jenis kelamin
        build(true);
        if (S.dest === 'pop' && S.mult === 1) { const v = S.rows.map(r => r.nums[0] && r.nums[0].val).filter(x => x > 0).sort((a, b) => a - b); if (v.length >= 3 && v[v.length >> 1] < 5000) S.mult = 1000 }   // angka terlalu kecil untuk jiwa -> ribu jiwa
        review();
    }
    function build(keepDest) {
        const P = BPS.parseText(S.text), k = DEST[S.dest].k; S.years = P.years;
        S.names = [...new Set(ROWS.map(kecOf))].filter(x => x !== '-').sort();
        if (k === 'kab') S.rows = P.rows.filter(r => /[A-Za-z]{3}/.test(r.label)).map(r => ({ ...r, target: BPS.matchName(r.label, S.names) || '' }));
        else {
            const rs = P.rows.filter(r => r.nums.length >= 3);
            if (k === 'pk') S.rows = rs.map((r, i) => ({ ...r, target: i < PK_STATUS.length ? String(i) : '' }));
            else {
                const codes = PD_USAHA.map(u => u[0].replace(/\s/g, '')), by = rs.map(r => codes.indexOf(BPS.codeOf(r.label)));
                const useCode = by.filter(i => i >= 0).length >= rs.length / 2;
                S.rows = rs.map((r, i) => ({ ...r, target: String(useCode ? by[i] : i < PD_USAHA.length ? i : -1).replace('-1', '') }));
            }
        }
        S.det = BPS.detectDecimals(S.rows); if (k === 'kab') S.col = Math.min(S.col, Math.max(0, S.det.d.length - 1));
        recalc();
    }
    function recalc() { BPS.applyDec(S.rows, S.det, DEST[S.dest].k === 'pd'); if (DEST[S.dest].k === 'pk') S.rows.forEach(r => { const [l, p, t] = r.nums.map(n => n.val); r.rowWarn = r.nums.length >= 3 && Math.abs(l + p - t) > 1 ? `Laki-laki + Perempuan (${fmtV(l + p)}) ≠ Total (${fmtV(t)}). Ada digit salah baca.` : '' }) }

    // ---------- Pratinjau ----------
    const colYear = c => S.years[c] !== undefined ? PD_TAHUN.indexOf(S.years[c]) : (c < PD_TAHUN.length ? c : -1);
    function targetOpts(sel) {
        const k = DEST[S.dest].k, o = ['<option value="">— abaikan —</option>'];
        if (k === 'kab') S.names.forEach(n => o.push(`<option ${n === sel ? 'selected' : ''}>${esc(n)}</option>`));
        else if (k === 'pk') PK_STATUS.forEach((n, i) => o.push(`<option value="${i}" ${String(i) === sel ? 'selected' : ''}>${esc(n)}</option>`));
        else PD_USAHA.forEach((u, i) => o.push(`<option value="${i}" ${String(i) === sel ? 'selected' : ''}>${esc(u[0] + ' · ' + u[1].slice(0, 38))}</option>`));
        return o.join('');
    }
    function review() {
        const box = $('bps-review'), k = DEST[S.dest].k, nc = Math.max(0, ...S.rows.map(r => r.nums.length));
        const heads = k === 'pk' ? ['Laki-laki', 'Perempuan', 'Total'] : k === 'pd' ? Array.from({ length: nc }, (_, c) => S.years[c] !== undefined ? (colYear(c) < 0 ? `${S.years[c]} ⚠ tidak ada di PD_TAHUN` : S.years[c]) : 'Kol ' + (c + 1)) : Array.from({ length: nc }, (_, c) => 'Kolom ' + (c + 1));
        const matched = S.rows.filter(r => r.target !== '').length;
        box.hidden = false;
        box.innerHTML = `<div class="bps-top"><img class="bps-thumb" src="${S.url}" alt="Gambar tabel BPS">
  <div class="bps-ctl">
   <label>Letakkan di<select id="bps-dest">${Object.entries(DEST).map(([v, d]) => `<option value="${v}" ${v === S.dest ? 'selected' : ''}>${d.t}</option>`).join('')}</select></label>
   ${k === 'kab' ? `<label>Ambil dari kolom<select id="bps-col">${heads.map((h, c) => `<option value="${c}" ${c === S.col ? 'selected' : ''}>${h}</option>`).join('')}</select></label>` : ''}
   ${S.dest === 'pop' ? `<label>Satuan angka<select id="bps-mult"><option value="1" ${S.mult === 1 ? 'selected' : ''}>jiwa (×1)</option><option value="1000" ${S.mult === 1000 ? 'selected' : ''}>ribu jiwa (×1.000)</option></select></label>` : ''}
   <label class="chk"><input type="checkbox" id="bps-sharp" checked> Pertajam gambar</label><button id="bps-redo" class="ghost">Baca ulang</button>
  </div></div>
  <div class="bps-dec">Desimal tiap kolom: ${S.det.d.map((d, c) => `<label>${heads[c] || c + 1}<input type="number" min="0" max="4" class="bps-d" data-c="${c}" value="${d}"></label>`).join('')}
   <small>Dipakai untuk mengembalikan koma yang hilang (mis. 588 → 58,8). Ubah bila tidak sesuai.</small></div>
  <div class="bps-scroll"><table class="bps-tb"><tr><th>Teks terbaca</th><th>Dimasukkan ke</th>${heads.map((h, c) => `<th class="${k === 'kab' && c === S.col ? 'sel' : ''}">${h}</th>`).join('')}</tr>
  ${S.rows.map((r, i) => `<tr class="${r.target === '' ? 'ign' : ''}" ${r.rowWarn ? `title="${esc(r.rowWarn)}"` : ''}><td>${esc(r.label || '—')}${r.rowWarn ? ' <b class="w">⚠</b>' : ''}</td>
   <td><select class="bps-t" data-r="${i}">${targetOpts(r.target)}</select></td>
   ${heads.map((_, c) => { const n = r.nums[c]; return n ? `<td class="${k === 'kab' && c !== S.col ? 'dim' : ''}"><input class="bps-in ${n.warn ? 'bad' : n.fixed ? 'fix' : n.flag ? 'chk2' : ''}" data-r="${i}" data-c="${c}" value="${fmtV(n.val)}" title="${esc(n.flag || r.rowWarn || '')}"></td>` : '<td></td>' }).join('')}</tr>`).join('')}</table></div>
  <div class="bps-foot"><span id="bps-sum"></span><button id="bps-apply">Terapkan ke ${DEST[S.dest].t.split(' · ')[0]}</button></div>
  <details><summary>Teks mentah hasil OCR</summary><pre>${esc(S.text)}</pre></details>`;
        sum();
    }
    function sum() {
        const k = DEST[S.dest].k, ok = S.rows.filter(r => r.target !== '').length, bad = document.querySelectorAll('#bps-review .bps-in.bad,#bps-review .w').length, fix = document.querySelectorAll('#bps-review .bps-in.fix').length;
        let t = `${ok} dari ${S.rows.length} baris akan diisi.`;
        if (k === 'kab') { const miss = S.names.filter(n => !S.rows.some(r => r.target === n)); t += miss.length ? ` Belum terisi (${miss.length}): ${miss.slice(0, 6).join(', ')}${miss.length > 6 ? '…' : ''}.` : ' Semua kabupaten/kota terisi.' }
        if (fix) t += ` ${fix} angka dikoreksi otomatis (kuning).`; if (bad) t += ` ${bad} perlu dicek (merah).`;
        $('bps-sum').textContent = t;
    }

    // ---------- Terapkan ----------
    function apply() {
        const k = DEST[S.dest].k; let n = 0;
        S.rows.forEach(r => {
            if (r.target === '') return;
            if (k === 'kab') { const x = r.nums[S.col]; if (!x || x.val === null) return; const v = S.dest === 'pop' ? Math.round(x.val * S.mult) : x.val;
                if (S.dest === 'pop') popSet(r.target, v); else t2Set(S.dest, r.target, v); n++ }
            else if (k === 'pk') { const i = +r.target; [['l', 0], ['p', 1]].forEach(([j, c]) => { const x = r.nums[c]; if (x && x.val !== null) { pkSet(i, j, x.val); n++ } }) }
            else r.nums.forEach((x, c) => { const y = colYear(c); if (y < 0 || x.val === null) return; (S.dest === 'pd' ? pdSet : pd2Set)(+r.target, y, x.val); n++ });
        });
        const on = DEST[S.dest].tab;
        render(); [...document.querySelectorAll('.tab')].find(t => t.textContent.startsWith(on))?.click();
        st(`${n.toLocaleString('id')} nilai diterapkan ke ${DEST[S.dest].t}. Gambar masih tersedia: ubah tujuan/kolom untuk mengisi tabel lain.`);
        $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    // ---------- Event ----------
    const setFile = f => { if (!f || !/^image\//.test(f.type)) { st('File harus berupa gambar (PNG/JPG).'); return } if (S.url) URL.revokeObjectURL(S.url); S.file = f; S.url = URL.createObjectURL(f); runOcr() };
    $('bps-pick').onclick = () => $('bps-file').click();
    $('bps-file').onchange = e => { setFile(e.target.files[0]); e.target.value = '' };
    const card = $('bps');
    ['dragover', 'dragenter'].forEach(t => card.addEventListener(t, e => { e.preventDefault(); card.classList.add('over') }));
    ['dragleave', 'drop'].forEach(t => card.addEventListener(t, e => { e.preventDefault(); card.classList.remove('over') }));
    card.addEventListener('drop', e => setFile(e.dataTransfer.files[0]));
    document.addEventListener('paste', e => { const f = [...(e.clipboardData?.files || [])].find(x => /^image\//.test(x.type)); if (f) setFile(f) });   // Ctrl+V screenshot
    $('bps-review').addEventListener('change', e => {
        const t = e.target;
        if (t.id === 'bps-dest') { S.dest = t.value; S.col = 0; build(); review() }
        else if (t.id === 'bps-col') { S.col = +t.value; review() }
        else if (t.id === 'bps-mult') S.mult = +t.value;
        else if (t.matches('.bps-d')) { S.det.d[+t.dataset.c] = Math.max(0, Math.min(4, parseInt(t.value) || 0)); recalc(); review() }
        else if (t.matches('.bps-t')) { S.rows[+t.dataset.r].target = t.value; t.closest('tr').classList.toggle('ign', t.value === ''); sum() }
        else if (t.matches('.bps-in')) { const n = S.rows[+t.dataset.r].nums[+t.dataset.c], v = parseFloat(String(t.value).replace(',', '.')); n.val = isNaN(v) ? null : v; n.edited = true; n.warn = n.flag = ''; n.fixed = false; t.className = 'bps-in'; t.title = ''; if (DEST[S.dest].k === 'pk') { recalc(); review() } else sum() }
    });
    $('bps-review').addEventListener('click', e => { if (e.target.id === 'bps-apply') apply(); else if (e.target.id === 'bps-redo') runOcr() });
    window.BPSUI = { S, setText: t => { S.url = S.url || ''; setText(t) }, apply };   // dipakai pengujian
})();
