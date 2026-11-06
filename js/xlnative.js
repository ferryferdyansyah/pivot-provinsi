// xlnative.js  (FILE BARU)
// 1) Helper rumus: membuat sel Excel berisi RUMUS (SUMIFS ke sheet Raw Data, SUM, persen, LARGE, dst).
// 2) Grafik ASLI Excel: ExcelJS tidak bisa membuat grafik, jadi grafik disuntikkan ke file .xlsx
//    (setelah ExcelJS selesai) memakai JSZip. Grafik membaca langsung dari sel tabel -> ikut berubah
//    kalau angka di tabel/Raw Data diubah.
// Muat file ini SETELAH config.js & sheets.js dan SEBELUM excel.js / extra.js.

// ---------------------------------------------------------------- RUMUS
// nama kolom (1,2,3..) -> huruf (A,B,..,AA)
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) } return s }
// alamat sel dari nomor baris/kolom: A1 = (1,1)
const A1 = (r, c, absR, absC) => `${absC ? '$' : ''}${colLetter(c)}${absR ? '$' : ''}${r}`;
// nama sheet Raw Data (sheet pertama berjenis 'raw' di SHEETS)
const rawName = () => sheetName(SHEETS.find(s => s.kind === 'raw'));
// rentang satu kolom Raw Data, mis. '00_Raw_Data'!$K$2:$K$5001 ; null jika kolom tidak ada
function rawRng(col) {
    const i = COLS.indexOf(col); if (i < 0 || !ROWS.length) return null;
    const L = colLetter(i + 1);
    return `'${rawName().replace(/'/g, "''")}'!$${L}$2:$${L}$${ROWS.length + 1}`;
}
// nilai sel berisi rumus + hasil sementara (supaya angka tetap tampil di aplikasi yang tidak menghitung ulang)
const FX = (formula, result) => ({ formula, result });

// ---------------------------------------------------------------- GRAFIK ASLI
let XCHARTS = [];                       // daftar grafik yang akan disuntikkan
const resetXCharts = () => { XCHARTS = [] };
// spec = { sheet, type:'bar'|'pie', title, series:[{name, nameRef, cat, catVals, val, vals, color}],
//          pointColors:[..] (pie), legend:'r'|'b'|null, from:{col,row} (0-based), ext:{w,h} (px), valFmt }
function addXChart(spec) { XCHARTS.push(spec) }

const _esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const _hex = c => String(c).replace('#', '').toUpperCase();
const _strCache = a => `<c:strCache><c:ptCount val="${a.length}"/>${a.map((v, i) => `<c:pt idx="${i}"><c:v>${_esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
const _numCache = (a, f) => `<c:numCache><c:formatCode>${_esc(f)}</c:formatCode><c:ptCount val="${a.length}"/>${a.map((v, i) => `<c:pt idx="${i}"><c:v>${+v || 0}</c:v></c:pt>`).join('')}</c:numCache>`;
const _txPr = (sz, bold) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${bold ? 1 : 0}"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;

function chartXML(sp) {
    const pie = sp.type === 'pie', hz = sp.dir === 'bar', vf = sp.valFmt || '#,##0.00';
    const title = [].concat(sp.title || '').join(' ');
    const ser = sp.series.map((s, i) => {
        const tx = s.nameRef ? `<c:tx><c:strRef><c:f>${_esc(s.nameRef)}</c:f>${_strCache([s.name])}</c:strRef></c:tx>`
            : s.name != null ? `<c:tx><c:v>${_esc(s.name)}</c:v></c:tx>` : '';
        const fill = c => `<c:spPr><a:solidFill><a:srgbClr val="${_hex(c)}"/></a:solidFill></c:spPr>`;
        const dpt = pie ? (sp.pointColors || []).map((c, k) =>
            `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${_hex(c)}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
        const dl = pie
            ? `<c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${_txPr(900, true)}<c:dLblPos val="bestFit"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>`
            : `<c:dLbls><c:numFmt formatCode="${_esc(vf)}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${_txPr(800, false)}<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
        const cat = `<c:cat><c:strRef><c:f>${_esc(s.cat)}</c:f>${s.catVals ? _strCache(s.catVals) : ''}</c:strRef></c:cat>`;
        const val = `<c:val><c:numRef><c:f>${_esc(s.val)}</c:f>${s.vals ? _numCache(s.vals, vf) : ''}</c:numRef></c:val>`;
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${pie ? '' : fill(s.color || NAVY)}${pie ? '' : '<c:invertIfNegative val="0"/>'}${dpt}${dl}${cat}${val}</c:ser>`;
    }).join('');
    const plot = pie
        ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
        : `<c:barChart><c:barDir val="${hz ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="80"/><c:axId val="111"/><c:axId val="222"/></c:barChart>`
        + `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="${hz ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${hz ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>${_txPr(900)}<c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
        + `<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${hz ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${_txPr(900)}<c:crossAx val="111"/><c:crosses val="${hz ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
    const legend = sp.legend ? `<c:legend><c:legendPos val="${sp.legend}"/><c:overlay val="0"/></c:legend>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>`
        + `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="id-ID" sz="1200" b="1"/><a:t>${_esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
        + `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
        + `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
}

function drawingXML(items) {   // items: [{from:{col,row}, ext:{w,h}, rid, n}]
    const E = 9525;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">`
        + items.map((it, i) => `<xdr:oneCellAnchor><xdr:from><xdr:col>${it.from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${it.from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${it.ext.w * E}" cy="${it.ext.h * E}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Grafik ${it.n}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${it.rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`).join('')
        + `</xdr:wsDr>`;
}

// Dipanggil SETELAH wb.xlsx.writeBuffer(). Mengembalikan Blob .xlsx yang sudah berisi grafik asli.
async function injectCharts(buf) {
    const zip = await JSZip.loadAsync(buf);
    const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    // peta nama sheet -> path file sheet (dibaca dari workbook.xml + rels, jadi tidak bergantung pada penomoran ExcelJS)
    const wbx = await zip.file('xl/workbook.xml').async('string'), rel = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const target = {}; for (const m of rel.matchAll(/<Relationship\b[^>]*>/g)) { const id = /Id="([^"]+)"/.exec(m[0]), t = /Target="([^"]+)"/.exec(m[0]); if (id && t) target[id[1]] = t[1] }
    const path = {}; for (const m of wbx.matchAll(/<sheet\b[^>]*>/g)) {
        const n = /name="([^"]*)"/.exec(m[0]), r = /r:id="([^"]+)"/.exec(m[0]); if (!n || !r) continue;
        const t = target[r[1]]; path[n[1]] = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
    }
    const maxN = re => Math.max(0, ...Object.keys(zip.files).map(f => (re.exec(f) || [0, 0])[1] | 0));
    let dn = maxN(/^xl\/drawings\/drawing(\d+)\.xml$/), cn = maxN(/^xl\/charts\/chart(\d+)\.xml$/);
    let ct = await zip.file('[Content_Types].xml').async('string'), ctAdd = '';
    const bySheet = {}; XCHARTS.forEach(c => (bySheet[c.sheet] ??= []).push(c));

    for (const [sheet, list] of Object.entries(bySheet)) {
        const sp = path[_esc(sheet)] || path[sheet]; if (!sp) { console.warn('Sheet tidak ditemukan untuk grafik:', sheet); continue }
        let sx = await zip.file(sp).async('string');
        if (/<drawing\b/.test(sx)) { console.warn('Sheet sudah punya gambar, grafik dilewati:', sheet); continue }
        const relPath = sp.replace('worksheets/', 'worksheets/_rels/') + '.rels';
        dn++; const items = [], drel = [];
        list.forEach((c, i) => {
            cn++; zip.file(`xl/charts/chart${cn}.xml`, chartXML(c));
            ctAdd += `<Override PartName="/xl/charts/chart${cn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
            items.push({ from: c.from, ext: c.ext, rid: 'rId' + (i + 1), n: cn });
            drel.push(`<Relationship Id="rId${i + 1}" Type="${RT}/chart" Target="../charts/chart${cn}.xml"/>`);
        });
        zip.file(`xl/drawings/drawing${dn}.xml`, drawingXML(items));
        zip.file(`xl/drawings/_rels/drawing${dn}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drel.join('')}</Relationships>`);
        ctAdd += `<Override PartName="/xl/drawings/drawing${dn}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
        // relasi sheet -> drawing
        const rid = 'rIdDrw' + dn, relXml = `<Relationship Id="${rid}" Type="${RT}/drawing" Target="../drawings/drawing${dn}.xml"/>`;
        const ex = zip.file(relPath);
        zip.file(relPath, ex ? (await ex.async('string')).replace('</Relationships>', relXml + '</Relationships>')
            : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relXml}</Relationships>`);
        // <drawing> harus berada sebelum legacyDrawing/tableParts/extLst
        const tag = `<drawing r:id="${rid}"/>`, at = sx.search(/<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s>\/]/);
        sx = at >= 0 ? sx.slice(0, at) + tag + sx.slice(at) : sx.replace('</worksheet>', tag + '</worksheet>');
        if (!/<worksheet\b[^>]*xmlns:r=/.test(sx)) sx = sx.replace('<worksheet', `<worksheet xmlns:r="${RT}"`);
        zip.file(sp, sx);
    }
    zip.file('[Content_Types].xml', ct.replace('</Types>', ctAdd + '</Types>'));
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------- HELPER KRITERIA (dipakai tabel ketersediaan / kesesuaian)
// teks literal untuk rumus: "Sawah" -> "Sawah" (tanda kutip digandakan)
const qs = x => '"' + String(x).replace(/"/g, '""') + '"';
// SUMIFS luas dengan kriteria tambahan: extra = [[rentang, kriteria], ...]; crit = daftar kriteria teks untuk kolom status
const _sumIfs = (fL, extra, rg, crit) => `SUMIFS(${fL}${extra.map(([r, c]) => `,${r},${c}`).join('')}${crit.map(c => `,${rg},"${c}"`).join('')})`;
// Ketersediaan (kolom VNAME): idx 0 = Tersedia, 1 = Tidak Tersedia ("tidak"/"tdk"/"belum" = tidak tersedia). null jika kolom tak ada.
function ketSumF(idx, extra = [], strict = false) {   // strict = cara hitung Tabel IV-11 (hanya awalan "tidak" yang dianggap Tidak Tersedia)
    const V = vCol(), fL = rawRng('LUASHA'), fV = V && rawRng(V); if (!fL || !fV) return null;
    const S = (...c) => _sumIfs(fL, extra, fV, c);
    if (strict) return idx === 0 ? S('*tersedia*', '<>tidak*') : S('tidak*');
    return idx === 0 ? S('*tersedia*', '<>*tidak*', '<>*tdk*', '<>*belum*')
        : `${S('*tidak*')}+${S('*tdk*', '<>*tidak*')}+${S('*belum*', '<>*tidak*', '<>*tdk*')}`;
}
// Kesesuaian (kolom KSPOLA): 'S' sesuai, 'M' mendukung, 'T' tidak sesuai
function ksSumF(code, extra = []) {
    const K = ksCol(), fL = rawRng('LUASHA'), fK = K && rawRng(K); if (!fL || !fK) return null;
    const S = (...c) => _sumIfs(fL, extra, fK, c);
    return code === 'T' ? `${S('*tidak*')}+${S('*tdk*', '<>*tidak*')}`
        : code === 'M' ? S('*mendukung*', '<>*tidak*', '<>*tdk*')
        : S('*sesuai*', '<>*tidak*', '<>*tdk*', '<>*mendukung*');
}

// Jumlah BARIS (bukan luas) per kode kesesuaian, untuk sel matriks S/T/M
function ksCntF(code, extra = []) {
    const K = ksCol(), fK = K && rawRng(K); if (!fK) return null;
    const C = (...c) => `COUNTIFS(${extra.map(([r, k]) => `${r},${k}`).concat(c.map(k => `${fK},"${k}"`)).join(',')})`;
    return code === 'T' ? `${C('*tidak*')}+${C('*tdk*', '<>*tidak*')}`
        : code === 'M' ? C('*mendukung*', '<>*tidak*', '<>*tdk*')
        : C('*sesuai*', '<>*tidak*', '<>*tdk*', '<>*mendukung*');
}
