const XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export function createXlsxWorkbook(rows, { sheetName = 'Extracted Table', firstRowHeader = true } = {}) {
  const safeRows = normalizeRows(rows);
  const name = safeSheetName(sheetName);
  const now = new Date().toISOString();
  const files = [
    ['[Content_Types].xml', contentTypesXml()],
    ['_rels/.rels', rootRelsXml()],
    ['docProps/app.xml', appPropsXml()],
    ['docProps/core.xml', corePropsXml(now)],
    ['xl/workbook.xml', workbookXml(name)],
    ['xl/_rels/workbook.xml.rels', workbookRelsXml()],
    ['xl/styles.xml', stylesXml()],
    ['xl/worksheets/sheet1.xml', sheetXml(safeRows, { firstRowHeader })],
  ];

  return zipStore(
    files.map(([path, xml]) => ({
      path,
      data: utf8(xml),
    }))
  );
}

export { XLSX_MIME };

function normalizeRows(rows) {
  const source = Array.isArray(rows) ? rows : [];
  const width = Math.max(1, ...source.map((row) => (Array.isArray(row) ? row.length : 0)));
  return source.length
    ? source.map((row) =>
        Array.from({ length: width }, (_, index) => String(row?.[index] ?? ''))
      )
    : [['']];
}

function sheetXml(rows, { firstRowHeader }) {
  const height = Math.max(1, rows.length);
  const width = Math.max(1, ...rows.map((row) => row.length));
  const ref = `A1:${columnName(width - 1)}${height}`;
  const columns = Array.from({ length: width }, (_, index) => {
    const maxLength = Math.max(
      8,
      ...rows.map((row) => String(row[index] || '').length)
    );
    const widthValue = Math.min(42, Math.max(10, maxLength + 2));
    const col = index + 1;
    return `<col min="${col}" max="${col}" width="${widthValue}" customWidth="1"/>`;
  }).join('');

  const sheetRows = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, columnIndex) => cellXml(value, rowIndex, columnIndex, firstRowHeader))
        .join('');
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join('');

  const freezeHeader = firstRowHeader && rows.length > 1
    ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
  const autoFilter = firstRowHeader && rows.length > 1 ? `<autoFilter ref="${ref}"/>` : '';

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${ref}"/>
  ${freezeHeader}
  <cols>${columns}</cols>
  <sheetData>${sheetRows}</sheetData>
  ${autoFilter}
</worksheet>`;
}

function cellXml(value, rowIndex, columnIndex, firstRowHeader) {
  const text = String(value ?? '');
  const ref = `${columnName(columnIndex)}${rowIndex + 1}`;
  const style = firstRowHeader && rowIndex === 0 ? ' s="1"' : '';
  const number = parseExcelNumber(text);
  if (number !== null) {
    return `<c r="${ref}"${style}><v>${number}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr"${style}><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
}

function parseExcelNumber(value) {
  let text = String(value || '').trim();
  if (!text || /[A-Za-z/]/.test(text)) return null;
  if (/^0\d+/.test(text)) return null;

  let negative = false;
  if (/^\(.+\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1).trim();
  }

  text = text.replace(/^[£$€]\s*/, '').replace(/,/g, '');
  if (text.endsWith('%')) text = text.slice(0, -1);
  if (!/^-?\d+(?:\.\d+)?$/.test(text)) return null;

  const number = Number(text);
  if (!Number.isFinite(number)) return null;
  return String(negative ? -Math.abs(number) : number);
}

function contentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function workbookXml(sheetName) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${escapeXmlAttribute(sheetName)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
}

function workbookRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
}

function appPropsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>PDF Workbench</Application>
</Properties>`;
}

function corePropsXml(now) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>PDF Workbench</dc:creator>
  <cp:lastModifiedBy>PDF Workbench</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(now)}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${escapeXml(now)}</dcterms:modified>
</cp:coreProperties>`;
}

function safeSheetName(name) {
  const cleaned = String(name || 'Extracted Table')
    .replace(/[\[\]:*?/\\]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return (cleaned || 'Extracted Table').slice(0, 31);
}

function columnName(index) {
  let n = index + 1;
  let name = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    n = Math.floor((n - mod) / 26);
  }
  return name;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeXmlAttribute(value) {
  return escapeXml(value).replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function utf8(text) {
  return new TextEncoder().encode(text);
}

function zipStore(files) {
  const central = [];
  const chunks = [];
  let offset = 0;
  const dosTime = 0;
  const dosDate = 0x5b21; // 2025-01-01

  for (const file of files) {
    const nameBytes = utf8(file.path);
    const data = file.data;
    const crc = crc32(data);
    const localHeader = bytesFromNumbers([
      ...u32(0x04034b50),
      ...u16(20),
      ...u16(0),
      ...u16(0),
      ...u16(dosTime),
      ...u16(dosDate),
      ...u32(crc),
      ...u32(data.length),
      ...u32(data.length),
      ...u16(nameBytes.length),
      ...u16(0),
    ]);

    chunks.push(localHeader, nameBytes, data);
    central.push({
      file,
      nameBytes,
      crc,
      offset,
    });
    offset += localHeader.length + nameBytes.length + data.length;
  }

  const centralOffset = offset;
  for (const entry of central) {
    const data = entry.file.data;
    const header = bytesFromNumbers([
      ...u32(0x02014b50),
      ...u16(20),
      ...u16(20),
      ...u16(0),
      ...u16(0),
      ...u16(dosTime),
      ...u16(dosDate),
      ...u32(entry.crc),
      ...u32(data.length),
      ...u32(data.length),
      ...u16(entry.nameBytes.length),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(entry.offset),
    ]);
    chunks.push(header, entry.nameBytes);
    offset += header.length + entry.nameBytes.length;
  }

  const centralSize = offset - centralOffset;
  chunks.push(
    bytesFromNumbers([
      ...u32(0x06054b50),
      ...u16(0),
      ...u16(0),
      ...u16(central.length),
      ...u16(central.length),
      ...u32(centralSize),
      ...u32(centralOffset),
      ...u16(0),
    ])
  );

  return concatBytes(chunks);
}

function crc32(data) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function u16(value) {
  return [value & 0xff, (value >>> 8) & 0xff];
}

function u32(value) {
  return [
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  ];
}

function bytesFromNumbers(numbers) {
  return Uint8Array.from(numbers);
}

function concatBytes(chunks) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}
