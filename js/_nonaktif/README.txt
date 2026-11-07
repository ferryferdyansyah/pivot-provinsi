Impor data statistik (BPS) dari PDF dinonaktifkan sementara.
Untuk mengaktifkan lagi: pindahkan pdf-import.js ke folder js/, kembalikan <section id="pdf-import"> di index.html
(lihat versi sebelumnya), dan tambahkan lagi <script src="js/pdf-import.js"> serta pdf.min.js.

bps-image.js & bps-parse.js: fitur lama isi data BPS dari gambar (OCR). Digantikan js/bps-pdf.js + js/bps-pdf-parse.js (baca langsung dari PDF BPS).
