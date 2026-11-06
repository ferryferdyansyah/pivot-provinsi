// main.js
// Penghubung: menangani upload/drag-drop, memanggil parser, render, loading, dan tombol download.
const DL_LABEL = `Download Excel (${SHEETS.length} sheet)`;
const refreshPdf = () => { try { window.updatePdfImportPreview && window.updatePdfImportPreview() } catch (e) { } };   // fitur PDF dinonaktifkan; aman diabaikan
const STEPS = ['Membaca file', 'Menggabungkan per kabupaten/kota', 'Membuat tabel & grafik'];

// ===== Input ZIP shapefile (.zip berisi .dbf) =====
async function load(file) {
    const st = $('status'), dl = $('dl'), dz = $('drop');
    st.textContent = 'Membaca ' + file.name + '…'; dl.disabled = true;
    Busy.show('Memuat ' + file.name, STEPS.filter((_, i) => i !== 1));
    try {
        const zip = await JSZip.loadAsync(file);
        const e = Object.values(zip.files).find(f => !f.dir && /\.dbf$/i.test(f.name));
        if (!e) { ROWS = []; COLS = []; render(); st.textContent = 'File .dbf tidak ditemukan di dalam ZIP.'; UI.drop(dz, 'error', file.name, 'File .dbf tidak ditemukan di dalam ZIP.'); return }
        Busy.detail('Membaca tabel atribut (.dbf)…');
        const r = parseDBF(await e.async('arraybuffer')), cc = canonCols(r.cols, r.rows);
        Busy.stage(1); Busy.detail(`${r.rows.length.toLocaleString('id')} baris`); await tick();
        ROWS = r.rows; COLS = cc.cols; UI.meta.excelRows = 0; render();
        const miss = ['LUASHA', 'WADMKC'].filter(c => !COLS.includes(c));
        st.textContent = `${file.name}: ${ROWS.length.toLocaleString('id')} baris, ${COLS.length} kolom.` + (cc.renamed.length ? ` Nama kolom disesuaikan: ${cc.renamed.join(', ')}.` : '') + (miss.length ? ` Peringatan: kolom ${miss.join(', ')} tidak ditemukan.` : '');
        UI.drop(dz, 'done', file.name, `${ROWS.length.toLocaleString('id')} baris, ${COLS.length} kolom`); dl.disabled = false;
    } catch (err) {
        ROWS = []; COLS = []; render(); st.textContent = 'Gagal membaca file: ' + err.message; UI.drop(dz, 'error', file.name, 'Gagal membaca file.');
    } finally { Busy.hide() }
}

// ===== Input Excel (.xlsx) besar: dibaca streaming, digabung per kabupaten/kota -> 1 Excel Provinsi Jawa Timur =====
const KOLOM_PAKAI = ['WADMKK', 'WADMKC', 'GNAME25', 'GREKLAS', 'QNAME25', 'QREKLAS', 'ONAME25', 'NAMOBJ', 'FKWS', 'GQREKLAS', 'GQNAME', 'KSPOLA', 'VNAME', 'V_ARAHAN', 'STD_C_G', 'STD_C_Q', 'TON_C_G', 'TON_C_Q', 'LUASHA'];
async function loadXlsx(file) {
    const st = $('status'), dl = $('dl'), dz = $('drop'), t0 = Date.now();
    st.textContent = 'Membaca ' + file.name + '…'; dl.disabled = true;
    Busy.show('Memuat ' + file.name, STEPS);
    try {
        const X = await readXlsxStream(file, n => {
            const s = `${n.toLocaleString('id')} baris dibaca (${Math.round((Date.now() - t0) / 1000)} dtk)`;
            st.textContent = `Membaca ${file.name}… ${s}`; Busy.detail(s)
        }, { keep: KOLOM_PAKAI, sum: ['LUASHA', 'TON_C_G', 'TON_C_Q'], unit: 'WADMKK' });   // WADMKC (kolom wilayah tabel) diisi nama kabupaten/kota
        Busy.stage(1); Busy.detail(`${X.total.toLocaleString('id')} baris Excel → digabung per kabupaten/kota`); await tick();
        window.NAMA_KAB = 'Jawa_Timur';
        const rows = X.rows.map(a => { const o = {}; X.cols.forEach((c, i) => o[c] = a[i]); return o }), cc = canonCols(X.cols.slice(), rows);
        Busy.stage(2); Busy.detail('Menyusun tabel dan grafik…'); await tick();
        ROWS = rows; COLS = cc.cols; UI.meta.excelRows = X.total; render();
        const nk = new Set(rows.map(r => r.WADMKK)).size;
        st.textContent = `${file.name}: ${X.total.toLocaleString('id')} baris Excel → ${ROWS.length.toLocaleString('id')} baris Raw Data (digabung per kabupaten/kota), ${nk} kabupaten/kota. Unggah ZIP HAT (opsional) lalu download Excel.`;
        UI.drop(dz, 'done', file.name, `${X.total.toLocaleString('id')} baris → ${ROWS.length.toLocaleString('id')} baris, ${nk} kab/kota`);
        dl.disabled = false;
    } catch (err) {
        ROWS = []; COLS = []; render(); st.textContent = 'Gagal membaca Excel: ' + err.message; UI.drop(dz, 'error', file.name, 'Gagal membaca file.');
    } finally { Busy.hide() }
}

// ===== Drop zone utama =====
const drop = $('drop'), fi = $('file');
const pick = f => f && (/\.xlsx$/i.test(f.name) ? loadXlsx(f) : load(f));
drop.onclick = () => fi.click(); fi.onchange = () => { pick(fi.files[0]); fi.value = '' };
['dragover', 'dragenter'].forEach(t => document.querySelectorAll('.dz').forEach(d => d.addEventListener(t, e => { e.preventDefault(); d.classList.add('over') })));
['dragleave', 'drop'].forEach(t => document.querySelectorAll('.dz').forEach(d => d.addEventListener(t, e => { e.preventDefault(); d.classList.remove('over') })));
drop.addEventListener('drop', e => pick(e.dataTransfer.files[0]));
// cegah browser membuka file bila dijatuhkan di luar kotak
['dragover', 'drop'].forEach(t => window.addEventListener(t, e => e.preventDefault()));

// ===== Download =====
$('dl').onclick = async () => {
    const b = $('dl'); b.disabled = true; b.textContent = 'Menyiapkan…';
    Busy.show('Menyusun file Excel', []); Busy.detail(`${SHEETS.length} sheet, ${ROWS.length.toLocaleString('id')} baris Raw Data. Mohon tunggu…`); await tick();
    try { await downloadXLSX(); $('status').textContent = 'File Excel berhasil dibuat dan diunduh.' }
    catch (err) { $('status').textContent = 'Gagal membuat Excel: ' + err.message }
    finally { Busy.hide(); b.disabled = false; b.textContent = DL_LABEL }
};
$('dl').textContent = DL_LABEL;

// ===== Tema terang/gelap =====
const tb = $('theme'), root = document.documentElement;
const paintTheme = () => { tb.textContent = root.dataset.theme === 'light' ? '☾' : '☀' };
tb.onclick = () => { root.dataset.theme = root.dataset.theme === 'light' ? 'dark' : 'light'; try { localStorage.theme = root.dataset.theme } catch (e) { } paintTheme() };
paintTheme();
render();
