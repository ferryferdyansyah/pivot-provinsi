// tahun.js - kotak input tahun penggunaan tanah lama dan baru. Dimuat SETELAH main.js.
(() => {
    const anchor = document.getElementById('settings'); if (!anchor) return;
    const box = document.createElement('div'); box.id = 'tahunbox';
    box.innerHTML = `<label>Tahun penggunaan tanah lama<input type="number" id="th-lama" min="1900" max="2100" value="${TH_LAMA}"></label>
    <label>Tahun penggunaan tanah baru<input type="number" id="th-baru" min="1900" max="2100" value="${TH_BARU}"></label>
    <small>Dipakai untuk judul tabel/grafik dan hitungan laju per tahun. Tersimpan di browser ini.</small>`;
    anchor.append(box);
    const a = document.getElementById('th-lama'), b = document.getElementById('th-baru'), st = document.getElementById('status');
    const apply = () => {
        const l = parseInt(a.value, 10), n = parseInt(b.value, 10);
        if (!(l >= 1900 && n <= 2100 && l < n)) { st.textContent = 'Tahun lama harus lebih kecil dari tahun baru (1900-2100).'; a.value = TH_LAMA; b.value = TH_BARU; return }
        if (l === TH_LAMA && n === TH_BARU) return;
        TH_LAMA = l; TH_BARU = n;
        try { localStorage.setItem('djpa_tahun', JSON.stringify({ lama: l, baru: n })) } catch (e) { }
        const on = document.querySelector('.tab.on'), idx = on ? on.dataset.i : null;
        render();
        if (idx !== null) document.querySelector(`.tab[data-i="${idx}"]`)?.click();
        st.textContent = `Tahun diubah: lama ${l}, baru ${n}.`;
    };
    a.onchange = b.onchange = apply;

    // ===== Tahun penggunaan tanah lama PER kabupaten/kota (opsional) =====
    const kb = document.createElement('details'); kb.id = 'tahunkab'; box.after(kb);
    const rerender = msg => {
        const on = document.querySelector('.tab.on'), idx = on ? on.dataset.i : null;
        render();
        if (idx !== null) document.querySelector(`.tab[data-i="${idx}"]`)?.click();
        if (msg) st.textContent = msg;
    };
    const saveKab = () => { try { localStorage.setItem('djpa_tahun_kab', JSON.stringify(TH_KAB)) } catch (e) { } };
    function buildKab() {
        const ks = kabList(), was = kb.open;
        kb.hidden = !ks.length; if (!ks.length) return;
        kb.innerHTML = `<summary>Tahun penggunaan tanah lama per kabupaten/kota (opsional)</summary>
      <div class="tk-tools"><span>Isi semua dengan</span><input type="number" id="tk-all" min="1900" max="${TH_BARU - 1}" placeholder="${TH_LAMA}"><button id="tk-apply" type="button">Terapkan</button><button id="tk-reset" type="button">Reset (pakai default ${TH_LAMA})</button></div>
      <div class="tk-grid">${ks.map(k => `<label>${k}<input type="number" class="tk-in" data-k="${k.replace(/"/g, '&quot;')}" min="1900" max="${TH_BARU - 1}" placeholder="${TH_LAMA}" value="${TH_KAB[k] ?? ''}"></label>`).join('')}</div>
      <small>Kosong = memakai "Tahun penggunaan tanah lama" di atas (${TH_LAMA}). Dipakai untuk laju perubahan per tahun (Tabel IV-6) dan judul tabel/grafik. Tersimpan di browser ini.</small>`;
        kb.open = was;
    }
    const okYear = y => y >= 1900 && y < TH_BARU;
    kb.addEventListener('change', e => {
        const t = e.target; if (!t.matches('.tk-in')) return;
        const y = parseInt(t.value, 10);
        if (t.value !== '' && !okYear(y)) { st.textContent = `Tahun lama harus antara 1900 dan ${TH_BARU - 1}.`; t.value = TH_KAB[t.dataset.k] ?? ''; return }
        if (t.value === '') delete TH_KAB[t.dataset.k]; else TH_KAB[t.dataset.k] = y;
        saveKab(); rerender(`Tahun lama ${t.dataset.k}: ${t.value === '' ? 'default ' + TH_LAMA : y}.`);
    });
    kb.addEventListener('click', e => {
        if (e.target.id === 'tk-apply') {
            const y = parseInt(document.getElementById('tk-all').value, 10);
            if (!okYear(y)) { st.textContent = `Tahun lama harus antara 1900 dan ${TH_BARU - 1}.`; return }
            kabList().forEach(k => TH_KAB[k] = y); saveKab(); rerender(`Tahun lama semua kabupaten/kota diisi ${y}.`);
        } else if (e.target.id === 'tk-reset') {
            kabList().forEach(k => delete TH_KAB[k]); saveKab(); rerender(`Tahun lama per kabupaten/kota dikembalikan ke default (${TH_LAMA}).`);
        }
    });
    // bangun ulang daftar kabupaten/kota setiap kali halaman dirender (mis. setelah Excel diunggah)
    const baseRender = render;
    render = function () { baseRender(); buildKab() };
    buildKab();
})();