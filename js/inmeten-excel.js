// ══════════════════════════════════════════════════════════════
// INMEET-EXCEL — maakt in de browser precies hetzelfde Excel-bestand
// als de tablet-app (ExcelHelper.java), zodat de Windows-app het
// gewoon kan lezen onder "Inmeet resultaten".
// Een .xlsx is een ZIP-bestand met XML erin. Hier wordt de ZIP zonder
// compressie ("store") opgebouwd, dat is geldig en heeft geen
// externe bibliotheek nodig.
// ══════════════════════════════════════════════════════════════

const XLSX_GROEN       = '2D5016';
const XLSX_DONKERGROEN = '1E3A0F';
const XLSX_LICHTGROEN  = 'EAF0E1';
const XLSX_GESTREEPT   = 'F5F9EE';
const XLSX_ORANJE      = 'E8820C';

// ── CRC32 (vereist door het ZIP-formaat) ──────────────────────
const _crcTabel = (function () {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = _crcTabel[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

// ── ZIP zonder compressie ─────────────────────────────────────
function maakZip(bestanden) { // bestanden: [{ naam, tekst }]
  const enc = new TextEncoder();
  const nu = new Date();
  const dosTijd = (nu.getHours() << 11) | (nu.getMinutes() << 5) | (nu.getSeconds() >> 1);
  const dosDatum = ((nu.getFullYear() - 1980) << 9) | ((nu.getMonth() + 1) << 5) | nu.getDate();

  const delen = [];
  const centraal = [];
  let offset = 0;

  bestanden.forEach(f => {
    const naamB = enc.encode(f.naam);
    const data  = enc.encode(f.tekst);
    const crc   = crc32(data);

    const lokaal = new DataView(new ArrayBuffer(30));
    lokaal.setUint32(0, 0x04034b50, true);
    lokaal.setUint16(4, 20, true);       // versie
    lokaal.setUint16(6, 0x0800, true);   // utf-8 bestandsnamen
    lokaal.setUint16(8, 0, true);        // methode 0 = store
    lokaal.setUint16(10, dosTijd, true);
    lokaal.setUint16(12, dosDatum, true);
    lokaal.setUint32(14, crc, true);
    lokaal.setUint32(18, data.length, true);
    lokaal.setUint32(22, data.length, true);
    lokaal.setUint16(26, naamB.length, true);
    lokaal.setUint16(28, 0, true);

    delen.push(new Uint8Array(lokaal.buffer), naamB, data);

    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true);
    cd.setUint16(4, 20, true);
    cd.setUint16(6, 20, true);
    cd.setUint16(8, 0x0800, true);
    cd.setUint16(10, 0, true);
    cd.setUint16(12, dosTijd, true);
    cd.setUint16(14, dosDatum, true);
    cd.setUint32(16, crc, true);
    cd.setUint32(20, data.length, true);
    cd.setUint32(24, data.length, true);
    cd.setUint16(28, naamB.length, true);
    cd.setUint16(30, 0, true);
    cd.setUint16(32, 0, true);
    cd.setUint16(34, 0, true);
    cd.setUint16(36, 0, true);
    cd.setUint32(38, 0, true);
    cd.setUint32(42, offset, true);
    centraal.push(new Uint8Array(cd.buffer), naamB);

    offset += 30 + naamB.length + data.length;
  });

  const cdGrootte = centraal.reduce((s, d) => s + d.length, 0);
  const eind = new DataView(new ArrayBuffer(22));
  eind.setUint32(0, 0x06054b50, true);
  eind.setUint16(4, 0, true);
  eind.setUint16(6, 0, true);
  eind.setUint16(8, bestanden.length, true);
  eind.setUint16(10, bestanden.length, true);
  eind.setUint32(12, cdGrootte, true);
  eind.setUint32(16, offset, true);
  eind.setUint16(20, 0, true);

  const alles = delen.concat(centraal, [new Uint8Array(eind.buffer)]);
  const totaal = alles.reduce((s, d) => s + d.length, 0);
  const uit = new Uint8Array(totaal);
  let pos = 0;
  alles.forEach(d => { uit.set(d, pos); pos += d.length; });
  return uit;
}

// ── XML helpers ───────────────────────────────────────────────
function xmlEsc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
    // tekens die niet in XML 1.0 mogen voorkomen verwijderen
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function xlsxRij(r, ht, cellen) {
  const h = ht > 0 ? ` ht="${ht}" customHeight="1"` : '';
  return `<row r="${r}"${h}>${cellen}</row>`;
}

// ── Het Excel-bestand zelf ────────────────────────────────────
// inmeting: { naam, adres, telefoon, email, locaties: [{ omschrijving, product, velden: [[naam, waarde], ...] }] }
function maakInmeetXlsx(inmeting, datum) {
  const d = datum || new Date();
  const datumLeesbaar = `${pad2(d.getDate())}-${pad2(d.getMonth() + 1)}-${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;

  const ss = []; // shared strings (ruwe tekst, wordt bij schrijven ge-escaped)
  const si = w => { let i = ss.indexOf(w); if (i < 0) { ss.push(w); i = ss.length - 1; } return i; };
  const cel = (kol, r, w, stijl) => `<c r="${kol}${r}" t="s" s="${stijl}"><v>${si(w)}</v></c>`;

  let sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="2" width="36" customWidth="1"/></cols>' +
    '<sheetData>';

  let r = 1;
  sheet += xlsxRij(r, 40, cel('A', r, 'ROZAL ZONWERING - OPMETING', 1)); r++;
  sheet += xlsxRij(r, 0, cel('A', r, 'Datum:', 3) + cel('B', r, datumLeesbaar, 4)); r++;
  sheet += xlsxRij(r, 0, cel('A', r, 'Naam:', 3) + cel('B', r, inmeting.naam, 4)); r++;
  sheet += xlsxRij(r, 0, cel('A', r, 'Adres:', 3) + cel('B', r, inmeting.adres, 4)); r++;
  if (inmeting.telefoon) { sheet += xlsxRij(r, 0, cel('A', r, 'Telefoon:', 3) + cel('B', r, inmeting.telefoon, 4)); r++; }
  if (inmeting.email)    { sheet += xlsxRij(r, 0, cel('A', r, 'E-mail:', 3) + cel('B', r, inmeting.email, 4)); r++; }
  sheet += xlsxRij(r, 0, ''); r++;

  inmeting.locaties.forEach((loc, li) => {
    sheet += xlsxRij(r, 22, cel('A', r, `LOCATIE ${li + 1}  |  ${loc.omschrijving}`, 5) + cel('B', r, ' ', 5)); r++;
    sheet += xlsxRij(r, 0, cel('A', r, 'Product:', 3) + cel('B', r, loc.product, 4)); r++;
    let gestreept = false;
    loc.velden.forEach(([k, v]) => {
      const waarde = v !== undefined && v !== null && String(v).trim() !== '' ? String(v) : '-';
      sheet += xlsxRij(r, 0, cel('A', r, k, gestreept ? 7 : 3) + cel('B', r, waarde, gestreept ? 8 : 4)); r++;
      gestreept = !gestreept;
    });
    sheet += xlsxRij(r, 8, ''); r++;
  });

  sheet += '</sheetData>' +
    '<mergeCells count="4"><mergeCell ref="A1:B1"/><mergeCell ref="B2:B2"/><mergeCell ref="B3:B3"/><mergeCell ref="B4:B4"/></mergeCells>' +
    '</worksheet>';

  const sharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${ss.length}" uniqueCount="${ss.length}">` +
    ss.map(s => `<si><t xml:space="preserve">${xmlEsc(s)}</t></si>`).join('') + '</sst>';

  return maakZip([
    { naam: '[Content_Types].xml', tekst: xlsxContentTypes() },
    { naam: '_rels/.rels', tekst: xlsxRels() },
    { naam: 'xl/_rels/workbook.xml.rels', tekst: xlsxWorkbookRels() },
    { naam: 'xl/workbook.xml', tekst: xlsxWorkbook() },
    { naam: 'xl/styles.xml', tekst: xlsxStyles() },
    { naam: 'xl/sharedStrings.xml', tekst: sharedStrings },
    { naam: 'xl/worksheets/sheet1.xml', tekst: sheet },
  ]);
}

function xlsxContentTypes() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '</Types>';
}
function xlsxRels() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';
}
function xlsxWorkbookRels() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '</Relationships>';
}
function xlsxWorkbook() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Opmeting" sheetId="1" r:id="rId1"/></sheets></workbook>';
}
function xlsxStyles() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="4">' +
    '<font><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="15"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
    '</fonts>' +
    '<fills count="7">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF${XLSX_GROEN}"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF${XLSX_DONKERGROEN}"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF${XLSX_LICHTGROEN}"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF${XLSX_ORANJE}"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF${XLSX_GESTREEPT}"/></patternFill></fill>` +
    '</fills>' +
    '<borders count="2">' +
    '<border><left/><right/><top/><bottom/></border>' +
    '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="9">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="2" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';
}
