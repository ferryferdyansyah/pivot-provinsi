pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const pdfNorm = value => String(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const pdfEscape = value => String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
const pdfNumberTokens = value => value.match(/(?:-?\d{1,3}(?:\.\d{3})+(?:,\d+)?|-?\d+\.\d+|-?\d+(?:,\d+)?|(?<!\d)-(?=\s|$)|[–—−])/g) || [];
const pdfNumber = value => {
  if (/^[–—−-]$/.test(value)) return null;
  if (value.includes(',')) return Number(value.replace(/\./g, '').replace(',', '.'));
  return Number(/^-?\d{1,3}(?:\.\d{3})+$/.test(value) ? value.replace(/\./g, '') : value);
};
const pdfPagesInput = key => document.querySelector(`[data-pdf-pages="${key}"]`).value;

function parsePdfPageSpec(spec) {
  const pages = new Map();
  for (const part of spec.split(/[;,]/).map(value => value.trim()).filter(Boolean)) {
    const prefix = part.match(/^(?:cetak|tercetak|folio)\s*:\s*/i);
    const printed = Boolean(prefix), value = printed ? part.slice(prefix[0].length) : part;
    const match = value.match(/^(\d+)\s*(?:-|–|—)\s*(\d+)$/);
    if (match) {
      const first = Number(match[1]), last = Number(match[2]);
      if (last < first || last - first > 30) throw new Error(`Rentang halaman tidak valid: ${part}`);
      for (let page = first; page <= last; page++) pages.set(`${printed}:${page}`, { number: page, printed });
    } else if (/^\d+$/.test(value)) pages.set(`${printed}:${value}`, { number: Number(value), printed });
    else throw new Error(`Format halaman tidak valid: ${part}`);
  }
  if (!pages.size) throw new Error('Nomor halaman untuk semua jenis data harus diisi.');
  return [...pages.values()].sort((a, b) => a.number - b.number);
}

async function pdfLines(document, pageNumber) {
  const page = await document.getPage(pageNumber), content = await page.getTextContent(), groups = [];
  content.items.forEach(item => {
    const value = (item.str || '').trim();
    if (!value) return;
    const x = item.transform ? item.transform[4] : 0, y = item.transform ? item.transform[5] : 0;
    let group = groups.find(line => Math.abs(line.y - y) < 2);
    if (!group) groups.push(group = { y, items: [] });
    group.items.push({ x, value });
  });
  return groups.sort((a, b) => b.y - a.y).map(line => line.items.sort((a, b) => a.x - b.x).map(item => item.value).join(' ').replace(/\s+/g, ' ').trim()).filter(Boolean);
}

async function resolvePrintedPage(document, printedPage, cache) {
  const labels = await document.getPageLabels();
  if (labels) {
    const index = labels.indexOf(String(printedPage));
    if (index >= 0) return index + 1;
  }
  const start = Math.max(1, printedPage - 2), end = Math.min(document.numPages, printedPage + 100);
  for (let page = start; page <= end; page++) {
    if (!cache.has(page)) cache.set(page, await pdfLines(document, page));
    const header = cache.get(page).slice(0, 4);
    if (header.some(line => new RegExp(`^\\s*${printedPage}(?:\\s|$)`).test(line))) return page;
  }
  if (printedPage <= document.numPages) return printedPage;
  throw new Error(`Halaman tercetak ${printedPage} tidak ditemukan. Periksa nomor halaman PDF.`);
}

async function resolvePdfPage(document, selection, cache) {
  if (selection.printed) return resolvePrintedPage(document, selection.number, cache);
  if (selection.number < 1 || selection.number > document.numPages) throw new Error(`Halaman PDF ${selection.number} di luar rentang 1-${document.numPages}.`);
  return selection.number;
}

function pdfDistrictRows(lines, preferredValueCount, expectedNames = [], minimumValueCount = 2) {
  const rows = {};
  for (const name of expectedNames) {
    const words = String(name).trim().split(/\s+/).map(word => word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const pattern = new RegExp(`(?:^|[^\\p{L}])(${words.join('\\s*')})(?=\\s|$)`, 'iu');
    for (const line of lines) {
      const match = pattern.exec(line);
      if (!match) continue;
      const values = pdfNumberTokens(line.slice(match.index + match[0].length)).map(pdfNumber);
      if (values.length < minimumValueCount && values.length) continue;
      const key = pdfNorm(name), previous = rows[key];
      const normalized = values.length ? values : Array(preferredValueCount).fill(0);
      if (!previous || normalized.length === preferredValueCount || (previous.values.length !== preferredValueCount && normalized.length > previous.values.length)) rows[key] = { name, values: normalized };
    }
  }
  for (const line of lines) {
    const match = line.match(/^\s*\[\d+\]\s*(.*?)\s+(?=[\d–—−-])/);
    if (!match) continue;
    const values = pdfNumberTokens(line.slice(match[0].length)).map(pdfNumber);
    if (values.length >= minimumValueCount) {
      const key = pdfNorm(match[1]), previous = rows[key];
      if (!previous || values.length === preferredValueCount || (previous.values.length !== preferredValueCount && values.length > previous.values.length)) rows[key] = { name: match[1].trim(), values };
    }
  }
  return rows;
}

function pdfEmploymentRows(lines) {
  const text = lines.join(' ').replace(/\s+/g, ' '), definitions = [
    { name: 'Berusaha sendiri', pattern: /berusaha\s+sendiri/i },
    { name: 'Buruh tidak dibayar', pattern: /berusaha\s+dibantu\s+(?:buruh\s+)?tidak\s+tetap(?:\s*\/?\s*(?:buruh\s+)?tidak\s+dibayar)?|berusaha\s+dibantu\s+buruh\s+tidak\s+dibayar/i },
    { name: 'Buruh dibayar', pattern: /berusaha\s+dibantu\s+(?:buruh\s+)?(?:tetap|dibayar)(?:\s*\/?\s*(?:buruh\s+)?(?:tetap|dibayar))?/i },
    { name: 'Buruh/Karyawan/Pegawai', pattern: /(?:buruh\s*\/?\s*)?karyawan\s*\/?\s*pegawai|buruh\s*\/?\s*karyawan\s*\/?\s*pegawai/i },
    { name: 'Pekerja bebas', pattern: /pekerja\s+bebas/i },
    { name: 'Pekerja keluarga/tak dibayar', pattern: /pekerja\s+keluarga(?:\s*\/?\s*(?:tak|tidak)\s+dibayar)?/i }
  ];
  const matches = definitions.map(definition => ({ ...definition, match: definition.pattern.exec(text) })).filter(row => row.match).sort((a, b) => a.match.index - b.match.index);
  return matches.map((row, index) => {
    const start = row.match.index + row.match[0].length, end = matches[index + 1] ? matches[index + 1].match.index : text.length;
    const values = pdfNumberTokens(text.slice(start, end)).map(pdfNumber);
    return { name: row.name, male: values[0] ?? 0, female: values[1] ?? 0 };
  });
}

function pdfIndustryRows(lines) {
  const aliases = {
    a: ['Pertanian', 'Agriculture'],
    b: ['Pertambangan', 'Mining'],
    c: ['Industri Pengolahan', 'Manufacturing'],
    d: ['Pengadaan Listrik', 'Electricity and Gas'],
    e: ['Pengadaan Air', 'Water Supply'],
    f: ['Konstruksi', 'Construction'],
    g: ['Perdagangan Besar dan Eceran', 'Wholesale and Retail'],
    h: ['Transportasi dan Pergudangan', 'Transportation and Storage'],
    i: ['Penyediaan Akomodasi', 'Accommodation and Food'],
    j: ['Informasi dan Komunikasi', 'Information and Communication'],
    k: ['Jasa Keuangan dan Asuransi', 'Financial and Insurance'],
    l: ['Real Estat', 'Real Estate'],
    mn: ['Jasa Perusahaan', 'Business Activities'],
    o: ['Administrasi Pemerintahan', 'Public Administration'],
    p: ['Jasa Pendidikan', 'Education'],
    q: ['Jasa Kesehatan', 'Human Health'],
    rstu: ['Jasa Lainnya', 'Other Services']
  };
  const labels = PD_USAHA.map(([code]) => ({
    code: pdfNorm(code),
    names: (aliases[pdfNorm(code)] || []).map(pdfNorm).sort((a, b) => b.length - a.length)
  })).sort((a, b) => Math.max(...b.names.map(name => name.length)) - Math.max(...a.names.map(name => name.length)));
  const codePattern = /^\s*\(?\s*(R\s*,\s*S\s*,\s*T\s*,\s*U|M\s*,\s*N|[A-L]|O|P|Q)\s*\)?(?:\s+(.*))?$/i;
  const headerAt = lines.findIndex(line => /lapangan usaha|industry/i.test(line) && /2021/.test(line));
  const tableLines = headerAt < 0 ? lines : lines.slice(headerAt + 1);
  const records = [];
  let current = null, pendingCode = null, started = false;
  const startRecord = (code, text = '') => {
    if (current && current.code !== code) records.push(current);
    if (current && current.code === code) current.text += ' ' + text;
    else current = { code, text };
    pendingCode = null;
    started = true;
  };
  const findLabel = text => {
    const normalized = pdfNorm(text);
    return labels.find(label => label.names.some(name => normalized.startsWith(name)));
  };
  for (const line of tableLines) {
    const trimmed = line.replace(/https?:\/\/\S+/gi, '').trim();
    if (!trimmed) continue;
    if (/lapangan usaha|industry/i.test(trimmed) && /2021/.test(trimmed)) continue;
    if (/^\(?\s*\d+\s*\)?(?:\s+\(?\s*\d+\s*\)?)*$/.test(trimmed)) continue;
    if (/produk domestik regional bruto|gross regional domestic product|catatan\s*\/\s*note|sumber\s*\/\s*source/i.test(trimmed)) {
      if (started) break;
      continue;
    }

    const codeMatch = codePattern.exec(trimmed);
    const code = codeMatch && PD_USAHA.find(row => pdfNorm(row[0]) === pdfNorm(codeMatch[1]));
    const remainder = codeMatch ? (codeMatch[2] || '') : trimmed;
    const label = findLabel(remainder);
    if (code && label && codeMatch[2]) {
      startRecord(label.code, remainder);
      continue;
    }
    if (code && !codeMatch[2]) {
      pendingCode = pdfNorm(code[0]);
      continue;
    }
    if (label) {
      startRecord(label.code, remainder);
      continue;
    }
    if (pendingCode) {
      const pending = labels.find(item => item.code === pendingCode);
      const hasPendingName = pending && pending.names.some(name => pdfNorm(trimmed).startsWith(name));
      if (hasPendingName) startRecord(pendingCode, trimmed);
      else if (current) current.text += ' ' + trimmed;
      continue;
    }
    if (current) current.text += ' ' + trimmed;
  }
  if (current) records.push(current);

  const rows = {};
  records.forEach(record => {
    const values = pdfNumberTokens(record.text).map(pdfNumber).slice(0, 5).map(value => value ?? 0);
    rows[record.code] = [...values, ...Array(Math.max(0, 5 - values.length)).fill(0)];
  });
  return rows;
}

const pdfHasHeader = (lines, patterns) => patterns.every(pattern => pattern.test(lines.join(' ')));

async function extractPdfStatistics(file) {
  const document = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const cache = new Map();
  try {
    const pageSpecs = Object.fromEntries(['population', 'ratio', 'employment', 'adhb', 'adhk'].map(key => [key, parsePdfPageSpec(pdfPagesInput(key))]));
    const read = async pages => (await Promise.all(pages.map(async selection => {
      const physical = await resolvePdfPage(document, selection, cache);
      if (!cache.has(physical)) cache.set(physical, await pdfLines(document, physical));
      return cache.get(physical);
    }))).flat();
    const [populationLines, ratioLines, employmentLines, adhbLines, adhkLines] = await Promise.all([
      read(pageSpecs.population), read(pageSpecs.ratio), read(pageSpecs.employment), read(pageSpecs.adhb), read(pageSpecs.adhk)
    ]);
    const headingText = lines => lines.join(' ').replace(/\b(?:table|tabel)\b/gi, ' ').replace(/\b\d+(?:\.\d+)?\b/g, ' ').replace(/\s+/g, ' ');
    const hasPopulationHeader = pdfHasHeader(populationLines, [/jumlah penduduk|population/i, /laju pertumbuhan|growth rate/i, /kecamatan|district/i]);
    const hasRatioHeader = pdfHasHeader(ratioLines, [/rasio jenis kelamin|population sex ratio/i, /kecamatan|district/i]);
    const hasEmploymentHeader = pdfHasHeader(employmentLines, [/status pekerjaan utama|main employment status/i, /laki[\s-]*laki|male/i, /perempuan|female/i]);
    const isPdrb = text => /produk domestik regional bruto|gross regional domestic product/i.test(text);
    const isPercentageDistribution = text => /distribusi persentase|percentage distribution/i.test(text);
    const adhbHeading = headingText(adhbLines), adhkHeading = headingText(adhkLines);
    const hasIndustryHeader = lines => /lapangan usaha|industry/i.test(lines.join(' ')) && [2021, 2022, 2023, 2024, 2025].every(year => lines.join(' ').includes(String(year)));
    if (!hasPopulationHeader) throw new Error('Header tabel penduduk/laju tidak ditemukan pada halaman yang dipilih.');
    if (!hasRatioHeader) throw new Error('Header tabel rasio jenis kelamin tidak ditemukan pada halaman yang dipilih.');
    if (!hasEmploymentHeader) throw new Error('Header tabel status pekerjaan (laki-laki/perempuan) tidak ditemukan pada halaman yang dipilih.');
    if (!isPdrb(adhbHeading) || isPercentageDistribution(adhbHeading) || !/harga\s+berlaku|current market prices/i.test(adhbHeading)) throw new Error('Halaman PDRB ADHB tidak dikenali. Pilih halaman berjudul Produk Domestik Regional Bruto Atas Dasar Harga Berlaku (Gross Regional Domestic Product at Current Market Prices), bukan tabel distribusi/persentase.');
    if (!isPdrb(adhkHeading) || isPercentageDistribution(adhkHeading) || !/harga\s+konstan|constant market prices/i.test(adhkHeading)) throw new Error('Halaman PDRB ADHK tidak dikenali. Pilih halaman berjudul Produk Domestik Regional Bruto Atas Dasar Harga Konstan (Gross Regional Domestic Product at Constant Market Prices), bukan tabel ADHB atau distribusi/persentase.');
    if (!hasIndustryHeader(adhbLines)) throw new Error('Header lapangan usaha dan tahun 2021–2025 tidak ditemukan pada halaman PDRB ADHB.');
    if (!hasIndustryHeader(adhkLines)) throw new Error('Header lapangan usaha dan tahun 2021–2025 tidak ditemukan pada halaman PDRB ADHK.');
    const expectedNames = buildT3()?.rows.map(row => row.k) || [];
    const population = pdfDistrictRows(populationLines, 6, expectedNames), ratios = pdfDistrictRows(ratioLines, 1, expectedNames, 1);
    const districts = (expectedNames.length ? expectedNames : Object.values(population).map(row => row.name)).map(name => {
      const popValues = population[pdfNorm(name)]?.values || Array(6).fill(0), ratioValues = ratios[pdfNorm(name)]?.values || [0];
      return {
        name,
        population: (popValues.length >= 4 ? popValues[1] : popValues[0]) ?? 0,
        growth: (popValues.length >= 4 ? popValues[3] : popValues[1]) ?? 0,
        ratio: (ratioValues.length >= 4 ? ratioValues[3] : ratioValues.length >= 2 ? ratioValues[1] : ratioValues[0]) ?? 0
      };
    });
    const parsedEmployment = new Map(pdfEmploymentRows(employmentLines).map(row => [pdfNorm(row.name), row]));
    const employment = PK_STATUS.map(name => {
      const row = parsedEmployment.get(pdfNorm(name));
      return { name, male: row?.male ?? 0, female: row?.female ?? 0 };
    });
    const normalizeIndustryRows = rows => Object.fromEntries(PD_USAHA.map(([code]) => {
      const values = rows[pdfNorm(code)] || [];
      return [pdfNorm(code), [...values.slice(0, 5), ...Array(Math.max(0, 5 - values.length)).fill(0)]];
    }));
    const adhb = normalizeIndustryRows(pdfIndustryRows(adhbLines)), adhk = normalizeIndustryRows(pdfIndustryRows(adhkLines));
    return { fileName: file.name, districts, employment, adhb, adhk, years: [2021, 2022, 2023, 2024, 2025] };
  } finally {
    await document.destroy();
  }
}

const pdfFile = document.getElementById('pdf-file'), pdfStatus = document.getElementById('pdf-status'), pdfReview = document.getElementById('pdf-review');
let pendingPdfStatistics = null;
const pdfFileButton = document.querySelector('.pdf-file-button');

function pdfTable(headers, rows) {
  return `<div class="raw-scroll"><table><tr>${headers.map(header => `<th>${pdfEscape(header)}</th>`).join('')}</tr>${rows.map(row => `<tr>${row.map(value => value && value.pdfEdit ? `<td class="n"><input class="pdf-edit-input" type="number" step="any" aria-label="${pdfEscape(value.label)}" data-pdf-edit-kind="${value.kind}" data-pdf-edit-row="${value.row}" data-pdf-edit-field="${value.field}"${value.year == null ? '' : ` data-pdf-edit-year="${value.year}"`} value="${pdfEscape(value.value)}"></td>` : `<td>${pdfEscape(value ?? '-')}</td>`).join('')}</tr>`).join('')}</table></div>`;
}

const pdfEditableCell = (value, kind, row, field, label, year) => ({ pdfEdit: true, value: value ?? 0, kind, row, field, label, year });

function updatePdfImportPreview() {
  if (!pendingPdfStatistics) return;
  const table = buildT3(), activeNames = table ? table.rows.map(row => row.k) : [];
  const byName = new Map(activeNames.map(name => [pdfNorm(name), name]));
  const matched = pendingPdfStatistics.districts.filter(row => byName.has(pdfNorm(row.name)));
  const unmatched = pendingPdfStatistics.districts.filter(row => !byName.has(pdfNorm(row.name))).map(row => row.name);
  const missing = activeNames.filter(name => !pendingPdfStatistics.districts.some(row => pdfNorm(row.name) === pdfNorm(name)));
  const warnings = [];
  if (!table) warnings.push('Unggah ZIP shapefile utama terlebih dahulu.');
  if (unmatched.length) warnings.push('Kecamatan PDF yang tidak cocok: ' + unmatched.join(', '));
  if (missing.length) warnings.push('Kecamatan shapefile tanpa data PDF: ' + missing.join(', '));
  if (table && !unmatched.length && !missing.length) warnings.push(`Wilayah ${wilayah()} cocok dengan seluruh kecamatan pada PDF.`);
  const industries = kind => PD_USAHA.map(([code, name], row) => [code, name, ...pendingPdfStatistics.years.map((year, index) => pdfEditableCell(pendingPdfStatistics[kind][pdfNorm(code)][index], kind, row, 'value', `${kind.toUpperCase()} ${code} ${year}`, index))]);
  pdfReview.innerHTML = `<p><b>${pdfEscape(pendingPdfStatistics.fileName)}</b> · penduduk/pekerjaan 2025 · PDRB 2021–2025</p>
    <p>${matched.length}/${activeNames.length} kecamatan cocok · ${pendingPdfStatistics.employment.length} kategori pekerjaan · ${PD_USAHA.length} lapangan usaha per seri PDRB.</p>
    ${warnings.length ? `<p class="note pdf-review-warnings">${warnings.map(pdfEscape).join('<br>')}</p>` : '<p class="note">Semua kecamatan cocok dengan shapefile aktif.</p>'}
    <details class="pdf-review-details"><summary>Pratinjau dan edit angka sebelum diterapkan</summary>
      <h3>Penduduk per kecamatan</h3>${pdfTable(['Kecamatan', 'Penduduk', 'Pertumbuhan (%)', 'Rasio jenis kelamin'], pendingPdfStatistics.districts.map((row, index) => [row.name, pdfEditableCell(row.population, 'district', index, 'population', `Penduduk ${row.name}`), pdfEditableCell(row.growth, 'district', index, 'growth', `Pertumbuhan ${row.name}`), pdfEditableCell(row.ratio, 'district', index, 'ratio', `Rasio jenis kelamin ${row.name}`)]))}
      <h3>Status pekerjaan 2025</h3>${pdfTable(['Status', 'Laki-laki', 'Perempuan'], pendingPdfStatistics.employment.map((row, index) => [row.name, pdfEditableCell(row.male, 'employment', index, 'male', `${row.name}, laki-laki`), pdfEditableCell(row.female, 'employment', index, 'female', `${row.name}, perempuan`)]))}
      <h3>PDRB ADHB (miliar rupiah)</h3>${pdfTable(['Kode', 'Lapangan usaha', ...pendingPdfStatistics.years], industries('adhb'))}
      <h3>PDRB ADHK (miliar rupiah)</h3>${pdfTable(['Kode', 'Lapangan usaha', ...pendingPdfStatistics.years], industries('adhk'))}
    </details>
    <label class="pdf-confirm"><input type="checkbox" id="pdf-confirm"> Saya sudah memeriksa pratinjau dan setuju mengganti isian sebelumnya</label>
    <div class="bar"><button id="pdf-apply" type="button" disabled>Terapkan data PDF</button></div>`;
  pdfReview.hidden = false;
  const confirm = pdfReview.querySelector('#pdf-confirm'), apply = pdfReview.querySelector('#pdf-apply');
  const canApply = table && matched.length === pendingPdfStatistics.districts.length && missing.length === 0;
  confirm.disabled = !canApply;
  apply.disabled = !canApply;
  confirm.onchange = () => { apply.disabled = !canApply || !confirm.checked };
  apply.onclick = applyPdfStatistics;
}

window.updatePdfImportPreview = updatePdfImportPreview;
pdfReview.addEventListener('change', event => {
  const input = event.target.closest('[data-pdf-edit-kind]');
  if (!input || !pendingPdfStatistics) return;
  const value = input.value === '' ? 0 : Number(input.value);
  if (!Number.isFinite(value)) return;
  const row = Number(input.dataset.pdfEditRow), kind = input.dataset.pdfEditKind;
  if (kind === 'district' || kind === 'employment') {
    pendingPdfStatistics[kind][row][input.dataset.pdfEditField] = value;
  } else if (kind === 'adhb' || kind === 'adhk') {
    const code = pdfNorm(PD_USAHA[row][0]), year = Number(input.dataset.pdfEditYear);
    pendingPdfStatistics[kind][code][year] = value;
  }
});
pdfFile.onchange = async () => {
  const file = pdfFile.files[0];
  if (!file) return;
  pendingPdfStatistics = null;
  pdfReview.hidden = true;
  pdfStatus.classList.remove('error');
  pdfStatus.textContent = `Membaca tabel dari ${file.name}...`;
  pdfFileButton.setAttribute('aria-disabled', 'true');
  try {
    pendingPdfStatistics = await extractPdfStatistics(file);
    pdfStatus.textContent = 'PDF terbaca. Periksa kecocokan dan pratinjau sebelum menerapkan.';
    updatePdfImportPreview();
  } catch (error) {
    console.error(error);
    pdfStatus.textContent = 'Gagal membaca PDF: ' + (error.message || error);
    pdfStatus.classList.add('error');
  } finally {
    pdfFileButton.removeAttribute('aria-disabled');
    pdfFile.value = '';
  }
};

function applyPdfStatistics() {
  if (!pendingPdfStatistics || !buildT3()) return;
  const activeRows = buildT3().rows, source = new Map(pendingPdfStatistics.districts.map(row => [pdfNorm(row.name), row]));
  if (activeRows.length !== source.size || activeRows.some(row => !source.has(pdfNorm(row.k)))) return;
  let appliedDistricts = 0;
  activeRows.forEach(row => {
    const values = source.get(pdfNorm(row.k));
    if (!values) return;
    const key = popKey(row.k);
    POP[key] = values.population;
    (T2S.laju ??= {})[key] = values.growth;
    (T2S.rasio ??= {})[key] = values.ratio;
    appliedDistricts++;
  });
  pendingPdfStatistics.employment.forEach((row, index) => {
    const statusIndex = PK_STATUS.findIndex(status => pdfNorm(status) === pdfNorm(row.name));
    if (statusIndex < 0) return;
    PKS[pkKey(statusIndex, 'l')] = row.male;
    PKS[pkKey(statusIndex, 'p')] = row.female;
  });
  const storeIndustries = (storage, keyFor, rows) => PD_USAHA.forEach(([code], index) => {
    const values = rows[pdfNorm(code)];
    if (values) pendingPdfStatistics.years.forEach((year, yearIndex) => {
      const targetYear = PD_TAHUN.indexOf(year);
      if (targetYear >= 0) storage[keyFor(index, targetYear)] = values[yearIndex];
    });
  });
  storeIndustries(PDS, pdKey, pendingPdfStatistics.adhb);
  storeIndustries(PD2S, pd2Key, pendingPdfStatistics.adhk);
  try {
    localStorage.setItem('djpa_pop', JSON.stringify(POP));
    localStorage.setItem('djpa_pop2', JSON.stringify(T2S));
    localStorage.setItem('djpa_pek', JSON.stringify(PKS));
    localStorage.setItem('djpa_pdrb', JSON.stringify(PDS));
    localStorage.setItem('djpa_pdrb2', JSON.stringify(PD2S));
  } catch (error) {
    pdfStatus.textContent = 'Data tidak dapat disimpan di browser: ' + error.message;
    pdfStatus.classList.add('error');
    return;
  }
  const activeTab = document.querySelector('#tabs .tab.on')?.dataset.i;
  render();
  if (activeTab !== undefined) document.querySelector(`#tabs .tab[data-i="${activeTab}"]`)?.click();
  pdfStatus.textContent = `Data PDF diterapkan untuk ${appliedDistricts} kecamatan, ${PK_STATUS.length} status pekerjaan, serta PDRB ADHB dan ADHK 2021–2025.`;
}