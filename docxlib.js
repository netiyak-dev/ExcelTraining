const d = require('docx');
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, TableOfContents, Footer, PageNumber, LevelFormat, convertInchesToTwip
} = d;

const FONT = 'TH Sarabun New';
const W = 9360; // usable width in DXA for A4 with 1" margins (approx 11906-2880=9026) -> use 9000
const TW = 9000;

const C = {
  navy: '1F3864', blue: '2E74B5', teal: '117864', grey: 'F2F2F2',
  head: 'D9E2F3', head2: 'DEEAF6', warn: 'FBE5D6', ok: 'E2EFDA'
};

function P(text, opts = {}) {
  const o = Object.assign({ size: 28, bold: false, color: '000000', spacing: { after: 80 }, align: null, italics: false }, opts);
  return new Paragraph({
    alignment: o.align || undefined,
    spacing: o.spacing,
    indent: o.indent,
    children: [new TextRun({ text, size: o.size, bold: o.bold, color: o.color, font: FONT, italics: o.italics })]
  });
}

function H(text, level) {
  const map = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  const size = { 1: 36, 2: 32, 3: 29 }[level];
  const color = { 1: C.navy, 2: C.blue, 3: C.teal }[level];
  return new Paragraph({
    heading: map[level],
    spacing: { before: level === 1 ? 320 : 240, after: 120 },
    children: [new TextRun({ text, size, bold: true, color, font: FONT })]
  });
}

function B(text, lvl = 0) {
  return new Paragraph({
    bullet: { level: lvl },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: 28, font: FONT })]
  });
}

// rich bullet: array of [text, {bold, color}]
function BR(runs, lvl = 0) {
  return new Paragraph({
    bullet: { level: lvl },
    spacing: { after: 60 },
    children: runs.map(([t, o = {}]) => new TextRun({ text: t, size: 28, font: FONT, bold: o.bold, color: o.color, italics: o.italics }))
  });
}

function cell(text, opts = {}) {
  const o = Object.assign({ w: 1000, bold: false, fill: null, size: 26, align: null, color: '000000' }, opts);
  const lines = Array.isArray(text) ? text : [text];
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: lines.map(l => new Paragraph({
      alignment: o.align || undefined,
      spacing: { after: 20 },
      children: [new TextRun({ text: l, size: o.size, bold: o.bold, font: FONT, color: o.color })]
    }))
  });
}

// rows: [[c1,c2,...], ...]; widths: array summing to TW
function T(widths, header, rows, opts = {}) {
  const hdrFill = opts.headFill || C.head;
  const trs = [];
  if (header) {
    trs.push(new TableRow({
      tableHeader: true,
      children: header.map((h, i) => cell(h, { w: widths[i], bold: true, fill: hdrFill, size: opts.size || 26 }))
    }));
  }
  rows.forEach((r, ri) => {
    trs.push(new TableRow({
      children: r.map((c, i) => {
        let val = c, fill = (ri % 2 === 1) ? C.grey : null;
        if (c && typeof c === 'object' && !Array.isArray(c)) { val = c.t; fill = c.fill || fill; }
        return cell(val, { w: widths[i], fill, size: opts.size || 26, bold: (c && c.bold) || false });
      })
    }));
  });
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: widths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' }
    },
    rows: trs
  });
}

const SP = () => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: '', size: 20, font: FONT })] });
const PB = () => new Paragraph({ children: [new PageBreak()] });

// callout box
function BOX(title, lines, fill) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [TW],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: 'A6A6A6' },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: 'A6A6A6' },
      left: { style: BorderStyle.SINGLE, size: 18, color: C.blue },
      right: { style: BorderStyle.SINGLE, size: 6, color: 'A6A6A6' },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE }
    },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: TW, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: fill || C.head2, color: 'auto' },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children: [
          new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: title, size: 27, bold: true, font: FONT, color: C.navy })] }),
          ...lines.map(l => new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: l, size: 26, font: FONT })] }))
        ]
      })]
    })]
  });
}

module.exports = { d, P, H, B, BR, cell, T, SP, PB, BOX, FONT, C, TW, Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, PageBreak, TableOfContents, Footer, PageNumber, LevelFormat, convertInchesToTwip };
