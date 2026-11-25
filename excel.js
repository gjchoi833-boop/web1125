"use strict";

const saveTripWorkbook = (() => {
  const encoder = new TextEncoder();
  const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const packageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

  function escapeXml(text) {
    return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
  }

  function xml(body) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
  }

  function cell(address, value, style = 0) {
    const attributes = `r="${address}" s="${style}"`;
    if (typeof value === "number") return `<c ${attributes}><v>${value}</v></c>`;
    return `<c ${attributes} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
  }

  function sheet({ departure, returning, context, rows, total, basis, note }) {
    const values = [
      ["국내 출장여비 계산 결과", undefined, 2, 0, 27],
      ["출장일", departure],
      ["복귀일", returning],
      [context, undefined],
      ["항목", "예상 지급액 (원)", 3, 3, 21],
      ...rows.map(([label, amount]) => [label, amount, 0, 1])
    ];
    const rowXml = values.map(([label, value, labelStyle, valueStyle, height], index) => {
      const row = index + 1;
      const size = height ? ` ht="${height}" customHeight="1"` : "";
      return `<row r="${row}"${size}>${cell("A" + row, label, labelStyle)}${value === undefined ? "" : cell("B" + row, value, valueStyle)}</row>`;
    });
    const totalRow = values.length + 1;
    const basisRow = totalRow + 2;
    const noteRow = basisRow + 1;
    rowXml.push(
      `<row r="${totalRow}" ht="21" customHeight="1">${cell("A" + totalRow, "예상 합계", 3)}<c r="B${totalRow}" s="4"><f>SUM(B6:B${totalRow - 1})</f><v>${total}</v></c></row>`,
      `<row r="${basisRow}">${cell("A" + basisRow, basis)}</row>`,
      `<row r="${noteRow}" ht="32" customHeight="1">${cell("A" + noteRow, note, 5)}</row>`
    );
    const merges = [1, 4, basisRow, noteRow].map(row => `<mergeCell ref="A${row}:B${row}"/>`);
    return xml(`<worksheet xmlns="${spreadsheet}">
      <sheetFormatPr defaultRowHeight="18"/>
      <cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="29" customWidth="1"/></cols>
      <sheetData>${rowXml.join("")}</sheetData>
      <mergeCells count="${merges.length}">${merges.join("")}</mergeCells>
    </worksheet>`);
  }

  function workbookFiles(result) {
    return {
      "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      </Types>`),
      "_rels/.rels": xml(`<Relationships xmlns="${packageRel}">
        <Relationship Id="rId1" Type="${officeRel}/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>`),
      "xl/workbook.xml": xml(`<workbook xmlns="${spreadsheet}" xmlns:r="${officeRel}">
        <sheets><sheet name="출장여비" sheetId="1" r:id="rId1"/></sheets>
      </workbook>`),
      "xl/_rels/workbook.xml.rels": xml(`<Relationships xmlns="${packageRel}">
        <Relationship Id="rId1" Type="${officeRel}/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="${officeRel}/styles" Target="styles.xml"/>
      </Relationships>`),
      "xl/worksheets/sheet1.xml": sheet(result),
      "xl/styles.xml": xml(`<styleSheet xmlns="${spreadsheet}">
        <fonts count="3">
          <font><sz val="11"/><name val="맑은 고딕"/></font>
          <font><b/><sz val="11"/><name val="맑은 고딕"/></font>
          <font><b/><sz val="16"/><color rgb="FF123D46"/><name val="맑은 고딕"/></font>
        </fonts>
        <fills count="3">
          <fill><patternFill patternType="none"/></fill>
          <fill><patternFill patternType="gray125"/></fill>
          <fill><patternFill patternType="solid"><fgColor rgb="FFE8F5F2"/><bgColor indexed="64"/></patternFill></fill>
        </fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="6">
          <xf xfId="0" numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          <xf xfId="0" numFmtId="3" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
          <xf xfId="0" numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
          <xf xfId="0" numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/>
          <xf xfId="0" numFmtId="3" fontId="1" fillId="2" borderId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
          <xf xfId="0" numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
        </cellXfs>
        <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
      </styleSheet>`)
    };
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) {
        crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
      }
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function zipWorkbook(files) {
    // XLSX is a ZIP package. These small XML parts use the standard stored method.
    const parts = [];
    const directory = [];
    let offset = 0;
    let directorySize = 0;

    for (const [path, content] of Object.entries(files)) {
      const name = encoder.encode(path);
      const bytes = encoder.encode(content);
      const checksum = crc32(bytes);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true); // ZIP 2.0
      local.setUint16(12, 0x21, true); // DOS date 1980-01-01
      local.setUint32(14, checksum, true);
      local.setUint32(18, bytes.length, true);
      local.setUint32(22, bytes.length, true);
      local.setUint16(26, name.length, true);
      parts.push(local.buffer, name, bytes);

      const entry = new DataView(new ArrayBuffer(46));
      entry.setUint32(0, 0x02014b50, true);
      entry.setUint16(4, 20, true);
      entry.setUint16(6, 20, true);
      entry.setUint16(14, 0x21, true);
      entry.setUint32(16, checksum, true);
      entry.setUint32(20, bytes.length, true);
      entry.setUint32(24, bytes.length, true);
      entry.setUint16(28, name.length, true);
      entry.setUint32(42, offset, true);
      directory.push(entry.buffer, name);
      directorySize += entry.byteLength + name.length;
      offset += local.byteLength + name.length + bytes.length;
    }

    const count = Object.keys(files).length;
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, count, true);
    end.setUint16(10, count, true);
    end.setUint32(12, directorySize, true);
    end.setUint32(16, offset, true);
    return new Blob([...parts, ...directory, end.buffer], { type: mime });
  }

  return (result) => {
    const url = URL.createObjectURL(zipWorkbook(workbookFiles(result)));
    const link = document.createElement("a");
    link.href = url;
    link.download = `출장여비_${result.departure}_${result.returning}.xlsx`;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
})();
