// config.js
// Pengaturan utama: daftar sheet (DEFS), warna, variabel data global, dan fungsi format angka.
// Mau ubah judul/tahun/kolom sumber? Cukup edit DEFS di bawah.

// ===== Tahun data (diubah lewat kotak "Tahun" di halaman; disimpan di browser) =====
// Teks di tabel/grafik yang tertulis 2014 dan 2026 otomatis diganti lewat yr().
let TH_LAMA = 2014, TH_BARU = 2026;
try { const t = JSON.parse(localStorage.getItem('djpa_tahun') || '{}'); if (t.lama) TH_LAMA = +t.lama; if (t.baru) TH_BARU = +t.baru } catch (e) { }
// Tahun penggunaan tanah lama PER kabupaten/kota (opsional; kosong = pakai TH_LAMA). Diisi lewat kotak di tahun.js, disimpan di browser.
let TH_KAB = {};
try { TH_KAB = JSON.parse(localStorage.getItem('djpa_tahun_kab') || '{}') || {} } catch (e) { }
let _kl = null;
function kabList() { if (!_kl || _kl.rows !== ROWS) _kl = { rows: ROWS, list: [...new Set(ROWS.map(r => r.WADMKK).filter(Boolean))].sort() }; return _kl.list }
const thLamaOf = k => { const v = +TH_KAB[k]; return v >= 1900 && v < TH_BARU ? v : TH_LAMA };
function thLamaRange() { const ys = kabList().map(thLamaOf); return ys.length ? [Math.min(...ys), Math.max(...ys)] : [TH_LAMA, TH_LAMA] }
// 2014 -> tahun lama (jika beda tiap kab/kota: "2015–2020"), 2026 -> tahun baru. "2014-2026" -> "2015–2020 s.d. 2026" bila tahun lama bervariasi.
const yr = s => {
  if (typeof s !== 'string') return s;
  const [a, b] = thLamaRange(), same = a === b, lama = same ? String(a) : a + '–' + b;
  return s.replace(/\b2014-2026\b|\b2014\b|\b2026\b/g, m => m === '2014-2026' ? (same ? a + '-' + TH_BARU : lama + ' s.d. ' + TH_BARU) : m === '2014' ? lama : TH_BARU);
};

const DEFS=[
 {f:'GNAME25',yearOld:true,tab:'06_Tabel_III-6',h:'Penggunaan Tanah Lama (Tahun 2014)',t5:'Penggunaan Tanah Lama (Tahun 2014)',ct:'LIMA PENGGUNAAN TANAH TERLUAS (HA)',type:'bar'},
 {f:'QNAME25',tab:'07_Tabel_III-7',h:'Penggunaan Tanah Baru (Tahun 2026)',t5:'Penggunaan Tanah Baru (Tahun 2026)',ct:'LIMA PENGGUNAAN TANAH TERLUAS (HA)',type:'bar'},
 {f:'ONAME25',tab:'08_Tabel_III-8',h:'Gambaran Umum Penguasaan Tanah Tahun 2026',ct:'GAMBARAN UMUM PENGUASAAN TANAH TAHUN 2026 (%)',type:'pie'},
 {f:'NAMOBJ',tab:'10_Tabel_III-10',h:'Rencana Pola Ruang pada RDTR/RTRW',t5:'Rencana Pola Ruang pada RDTR/RTRW Terbesar',ct:'LIMA RENCANA POLA RUANG',type:'bar'},
 {f:'FKWS',tab:'11_Tabel_III-11',h:'Kawasan Hutan',ct:'KAWASAN HUTAN TAHUN 2026 (%)',type:'pie'}
];
// judul kolom yang memuat tahun ikut berubah saat tahun diganti
DEFS.forEach(d => ['h', 't5'].forEach(k => { if (d[k]) { const v = d[k]; Object.defineProperty(d, k, { get: () => yr(v), configurable: true }) } }));
const PAL=['#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47','#264478','#9e480e'];
const NAVY='#1f3864';
let ROWS=[],COLS=[],MODELS=[];
const fmt=n=>n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});

// ===== Pencocokan nama kolom yang longgar =====
const CANON = ['LUASHA', 'WADMKC', 'WADMKK', 'GNAME25', 'QNAME25', 'ONAME25', 'NAMOBJ', 'FKWS', 'V_ARAHAN', 'VNAME', 'GQNAME', 'GQREKLAS', 'GREKLAS', 'QREKLAS', 'KSPOLA', 'STD_C_G', 'STD_C_Q', 'TON_C_G', 'TON_C_Q'];
// Nama lain yang dianggap sama (tulis tanpa garis bawah/spasi, huruf besar). Tambah sendiri bila perlu.
const ALIAS = { KECAMATAN: 'WADMKC', NAMAKECAMATAN: 'WADMKC', KABUPATEN: 'WADMKK', KABKOTA: 'WADMKK', KABKOT: 'WADMKK', LUASHEKTAR: 'LUASHA' };
const nkol = s => String(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
function canonCols(cols, rows, list = CANON) {
    const want = {}; list.forEach(c => want[nkol(c)] = c);
    Object.entries(ALIAS).forEach(([a, c]) => { if (list.includes(c)) want[a] = c });
    const map = {}, used = new Set(cols);
    cols.forEach(c => { const t = want[nkol(c)]; if (t && t !== c && !used.has(t)) { map[c] = t; used.add(t) } });
    const ks = Object.keys(map); if (!ks.length) return { cols, renamed: [] };
    for (const r of rows) for (const o of ks) { r[map[o]] = r[o]; delete r[o] }
    return { cols: cols.map(c => map[c] || c), renamed: ks.map(o => o + ' → ' + map[o]) };
}
