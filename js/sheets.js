// sheets.js
// DAFTAR URUTAN SEMUA SHEET (dipakai oleh tampilan web DAN file Excel).
// Mau ubah urutan / tambah sheet baru? Cukup edit array SHEETS di bawah ini.
//   kind:'raw'   -> data mentah dari .dbf
//   kind:'def'   -> sheet dari DEFS (config.js), i = nomor urut di DEFS (mulai 0)
//   kind:'extra' -> sheet dari EXTRA (extra.js), i = nomor urut di EXTRA (mulai 0)
//   kind:'new'   -> sheet baru yang isinya belum dibuat (tampil "Data Belum Ditemukan")
// Untuk 'def' dan 'extra', nama tab diambil dari properti `tab` di config.js / extra.js.
// Untuk 'new', ganti kind-nya menjadi 'extra'/'def' setelah logikanya dibuat.

const NEW_TEXT = 'Data Belum Ditemukan';
const nw = name => ({ kind: 'new', name });

const SHEETS = [
  { kind: 'raw', name: '00_Raw_Data' },
  // nw('00_Pivot_Data'),
  { kind: 'extra', i: 16 },  // 01_Tabel_III-1 (Luas wilayah + penduduk manual)
  { kind: 'extra', i: 17 },  // 02_Tabel_III-2 (Penduduk, laju, rasio)
  { kind: 'extra', i: 18 },  // 03_Tabel_III-3 (Status Pekerjaan Utama)
  { kind: 'extra', i: 19 },  // 04_Tabel_III-4 (PDRB per Lapangan Usaha)
  { kind: 'extra', i: 20 },  // 05_Tabel_III-5 (PDRB per Lapangan Usaha, set kedua)
  { kind: 'def', i: 0 },     // 06_Tabel_III-6  (Penggunaan Tanah Lama)
  { kind: 'def', i: 1 },     // 07_Tabel_III-7  (Penggunaan Tanah Baru)
  { kind: 'def', i: 2 },     // 08_Tabel_III-8  (Penguasaan Tanah)
  { kind: 'extra', i: 23 },  // 09_Tabel_III-9 (Galat Tanah Belum Terdaftar - kerangka)
  { kind: 'def', i: 3 },     // 10_Tabel_III-10 (RTRW-RDTR)
  { kind: 'def', i: 4 },     // 11_Tabel_III-11 (Kawasan Hutan)
  { kind: 'extra', i: 0 },   // 12_Tabel_IV-1   (Perubahan per Kabupaten/Kota)
  { kind: 'extra', i: 1 },   // 13_Tabel_IV-2   (Perbandingan Penggunaan)
  { kind: 'extra', i: 15 },  // 14_Tabel_IV-3  (Matriks Lama x Baru)
  { kind: 'extra', i: 2 },   // 15_Grafik_IV-1  (Perubahan Terluas)
  { kind: 'extra', i: 14 },  // 16_Tabel_IV-4  (Perubahan per RTRW)
  { kind: 'extra', i: 3 },   // 17_Tabel_IV-5   (Reklasifikasi)
  { kind: 'extra', i: 21 },  // 18_Tabel_IV-6 (Laju Perubahan Reklasifikasi)
  { kind: 'extra', i: 13 },  // 19_Tabel_IV-7  (Matriks S/T/M)
  { kind: 'extra', i: 12 },  // 20_Tabel_IV-8  (Kesesuaian terhadap RDTR)
  { kind: 'extra', i: 11 },  // 21_Tabel_IV-9  (Ketersediaan per Kabupaten/Kota)
  { kind: 'extra', i: 10 },  // 22_Tabel_IV-10 (Ketersediaan x Penggunaan Tanah)
  { kind: 'extra', i: 9 },   // 23_Tabel_IV-11
  { kind: 'extra', i: 8 },   // 24_Tabel_IV-12
  { kind: 'extra', i: 5 },
  { kind: 'extra', i: 6 },
  { kind: 'extra', i: 7 },
  { kind: 'extra', i: 4 },
  { kind: 'extra', i: 22 }   // 29_Tabel_IV-18 (Potensi Sosial Ekonomi)
];

const sheetName = s => s.kind === 'def' ? DEFS[s.i].tab : s.kind === 'extra' ? EXTRA[s.i].tab : s.name;
// nomor panel (id="p<n>") untuk sheet def/extra ke-i
const panelOf = (kind, i) => SHEETS.findIndex(s => s.kind === kind && s.i === i);
