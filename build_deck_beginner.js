/**
 * สไลด์สอนระดับ Beginner — รากฐานของข้อมูลที่เชื่อถือได้ (3 ชั่วโมง)
 * ชุดการเรียนรู้ Excel & Google Sheets — อาจารย์ ดร.เนติยา การะเกตุ
 * โครงสไลด์ตรงกับ rundown ในเอกสาร Blueprint ข้อ 2.2 ทุกช่วงเวลา
 * ตัวอย่างอ้างอิงไฟล์ฝึกปฏิบัติ 7 สาขา ตัวเลขดึงจาก answer_key.json
 */
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const L = require('./deck_lib.js');
const { W, H, FONT, startSlide, checkOverlaps, report } = L;

const key = JSON.parse(fs.readFileSync(__dirname + '/answer_key.json', 'utf8'));

/* จานสีประจำระดับ Beginner — น้ำเงิน ต่างจาก Intermediate (เขียวน้ำทะเล) และ Advanced (ม่วง) */
const P = {
  ink:   '122240',
  deep:  '1B4F8A',
  blue:  '2E74B5',
  sky:   '5B9BD5',
  amber: 'E8A33D',
  coral: 'D9584F',
  white: 'FFFFFF',
  soft:  'EDF3FA',
  soft2: 'FFF6E6',
  warn:  'FCEAE8',
  ink70: '33445C',
  muted: '6B7C93',
  line:  'D3E0EE'
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Dr. Netiya Karaket';
pres.title = 'Beginner — รากฐานของข้อมูลที่เชื่อถือได้';

const shadow = () => ({ type: 'outer', color: '122240', blur: 10, offset: 2, angle: 90, opacity: 0.10 });

/* ---------- ระบบจัดสมดุลแนวตั้งสองรอบ ---------- */
const OFFSET_FILE = __dirname + '/deck_offsets_beg.json';
const OFFSETS = fs.existsSync(OFFSET_FILE) ? JSON.parse(fs.readFileSync(OFFSET_FILE, 'utf8')) : {};
const HEAD_BAND = 1.1;
let curOffset = 0, curNo = 0;

const regRaw = L.reg;
function reg(kind, x, y, w, h, text, size, opts) {
  return regRaw(kind, x, y >= HEAD_BAND ? y + curOffset : y, w, h, text, size, opts);
}
function shiftOpts(o) {
  if (o && typeof o.y === 'number' && o.y >= HEAD_BAND) { const c = Object.assign({}, o); c.y = o.y + curOffset; return c; }
  return o;
}
function wrapSlide(s) {
  const at = s.addText.bind(s), as = s.addShape.bind(s);
  s.addText = function (t, o) { return at(t, shiftOpts(o)); };
  s.addShape = function (t, o) { return as(t, shiftOpts(o)); };
  return s;
}
function beginSlide(name) {
  const info = startSlide(name);
  curNo = info.no; curOffset = OFFSETS[String(curNo)] || 0;
  return info;
}

/* ---------- ชิ้นส่วนที่ใช้ซ้ำ ---------- */
function darkSlide() { const s = pres.addSlide(); s.background = { color: P.ink }; return wrapSlide(s); }
function lightSlide() { const s = pres.addSlide(); s.background = { color: P.white }; return wrapSlide(s); }

function head(s, num, title, sub) {
  const x = 0.62, y = 0.42;
  if (num) {
    s.addShape(pres.ShapeType.ellipse, { x, y: y + 0.02, w: 0.52, h: 0.52, fill: { color: P.blue }, line: { width: 0 } });
    s.addText(String(num), { x, y: y + 0.02, w: 0.52, h: 0.52, align: 'center', valign: 'middle', fontSize: 20, bold: true, color: P.white, fontFace: FONT, margin: 0 });
    reg('num', x, y, 0.52, 0.56);
  }
  const tx = num ? x + 0.72 : x, tw = W - (num ? x + 0.72 : x) - 0.62;
  s.addText(title, { x: tx, y, w: tw, h: 0.62, fontSize: 29, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
  reg('title', tx, y, tw, 0.62, title, 29, { pad: 0.05 });
  if (sub) {
    s.addText(sub, { x: tx, y: y + 0.62, w: tw, h: 0.36, fontSize: 14, color: P.muted, fontFace: FONT, valign: 'top', margin: 0 });
    reg('sub', tx, y + 0.62, tw, 0.36, sub, 14, { pad: 0.05 });
  }
  return y + (sub ? 1.08 : 0.78);
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.09, fill: { color: fill || P.soft }, line: { width: 0 }, shadow: shadow() });
  reg('card', x, y, w, h, null, null, { soft: true });
}

function textCard(s, x, y, w, h, title, bodyText, fill, titleColor) {
  card(s, x, y, w, h, fill);
  const px = x + 0.24, pw = w - 0.48;
  s.addText(title, { x: px, y: y + 0.16, w: pw, h: 0.42, fontSize: 16, bold: true, color: titleColor || P.deep, fontFace: FONT, margin: 0, valign: 'middle' });
  reg('cardTitle', px, y + 0.16, pw, 0.42, title, 16, { pad: 0.02 });
  const bh = h - 0.72;
  s.addText(bodyText, { x: px, y: y + 0.62, w: pw, h: bh, fontSize: 14, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top', lineSpacingMultiple: 1.25 });
  reg('cardBody', px, y + 0.62, pw, bh, bodyText, 14, { pad: 0.02 });
}

function bullets(s, x, y, w, h, items, size) {
  size = size || 14;
  const runs = items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } }));
  s.addText(runs, { x, y, w, h, fontSize: size, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top', paraSpaceAfter: 6 });
  reg('bullets', x, y, w, h, items.join('\n'), size, { pad: 0.35 });
}

/* ---------- ภาพจำลองหน้าจอสเปรดชีต วาดด้วยตารางจริงของ PowerPoint ---------- */
function sheetMock(s, x, y, w, cols, rows, tone) {
  const bad = tone === 'bad';
  const hFill = bad ? 'FBDDDB' : 'DCEEDF';
  const hText = bad ? '8E2F2B' : '1E5C2A';
  const header = cols.map(c => ({
    text: String(c),
    options: { fill: { color: hFill }, color: hText, bold: true, align: 'center' }
  }));
  const body = rows.map(r => r.map(c => {
    const t = Array.isArray(c) ? c[0] : c;
    const flag = Array.isArray(c) ? (c[1] || '') : '';
    return {
      text: String(t),
      options: {
        fill: { color: flag.indexOf('x') >= 0 ? 'FBDDDB' : 'FFFFFF' },
        color: flag.indexOf('x') >= 0 ? '8E2F2B' : P.ink70,
        bold: flag.indexOf('x') >= 0,
        align: flag.indexOf('n') >= 0 ? 'right' : 'left'
      }
    };
  }));
  const rowH = 0.3;
  s.addTable([header].concat(body), {
    x, y, w, rowH, fontFace: FONT, fontSize: 11.5, valign: 'middle',
    border: { type: 'solid', color: 'C6D2E0', pt: 0.5 }, margin: 4
  });
  reg('mock', x, y, w, rowH * (rows.length + 1), null, null, { soft: true });
  return y + rowH * (rows.length + 1);
}

function mockPair(name, num, title, sub, leftLabel, rightLabel, left, right, why) {
  beginSlide(name);
  const s = lightSlide();
  const y = head(s, num, title, sub);
  const cw = (W - 1.24 - 0.4) / 2;
  const xs = [0.62, 0.62 + cw + 0.4];
  const labels = [leftLabel, rightLabel];
  const tones = ['bad', 'good'];
  const data = [left, right];
  let maxB = y;
  for (let i = 0; i < 2; i++) {
    s.addText(labels[i], {
      x: xs[i], y, w: cw, h: 0.34, fontSize: 13.5, bold: true, fontFace: FONT, margin: 0,
      color: i === 0 ? 'A83A36' : '1E5C2A'
    });
    reg('lb' + i, xs[i], y, cw, 0.34, labels[i], 13.5, { pad: 0.02 });
    const b = sheetMock(s, xs[i], y + 0.42, cw, data[i].cols, data[i].rows, tones[i]);
    if (b > maxB) { maxB = b; }
  }
  const wy = maxB + 0.3;
  card(s, 0.62, wy, W - 1.24, 1.0, P.soft2);
  s.addText(why, {
    x: 0.9, y: wy + 0.16, w: W - 1.8, h: 0.7, fontSize: 13.5, color: 'A06A0A',
    fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('why', 0.9, wy + 0.16, W - 1.8, 0.7, why, 13.5, { pad: 0.02 });
  checkOverlaps();
  return s;
}

function stat(s, x, y, w, big, label, color) {
  s.addText(big, { x, y, w, h: 0.9, fontSize: 44, bold: true, color: color || P.deep, fontFace: FONT, align: 'center', margin: 0, valign: 'middle' });
  reg('statBig', x, y, w, 0.9);
  s.addText(label, { x, y: y + 0.92, w, h: 0.66, fontSize: 13, color: P.muted, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
  reg('statLabel', x, y + 0.92, w, 0.66, label, 13, { pad: 0.1 });
}

function timeRow(s, x, y, w, time, topic, detail, accent) {
  const tw = 1.15;
  s.addShape(pres.ShapeType.roundRect, { x, y, w: tw, h: 0.46, rectRadius: 0.07, fill: { color: accent || P.blue }, line: { width: 0 } });
  s.addText(time, { x, y, w: tw, h: 0.46, fontSize: 12, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
  reg('time', x, y, tw, 0.46);
  s.addText(topic, { x: x + tw + 0.16, y, w: 4.1, h: 0.46, fontSize: 14, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
  reg('timeTopic', x + tw + 0.16, y, 4.1, 0.46, topic, 14, { pad: 0.02 });
  const dx = x + tw + 0.16 + 4.1 + 0.14, dw = x + w - dx;
  s.addText(detail, { x: dx, y, w: dw, h: 0.46, fontSize: 12.5, color: P.muted, fontFace: FONT, valign: 'middle', margin: 0 });
  reg('timeDetail', dx, y, dw, 0.46, detail, 12.5, { pad: 0.02 });
}

/* ============================================================ สไลด์ */

/* 1 ปก */
{
  beginSlide('ปก');
  const s = darkSlide();
  s.addShape(pres.ShapeType.ellipse, { x: 10.2, y: -1.6, w: 5.4, h: 5.4, fill: { color: P.deep }, line: { width: 0 } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.8, y: 4.5, w: 2.8, h: 2.8, fill: { color: P.blue }, line: { width: 0 } });
  s.addText('ระดับที่ 1', { x: 0.9, y: 1.75, w: 6, h: 0.4, fontSize: 15, color: P.sky, fontFace: FONT, margin: 0, charSpacing: 2 });
  reg('eyebrow', 0.9, 1.75, 6, 0.4);
  s.addText('Beginner', { x: 0.9, y: 2.2, w: 8.8, h: 1.0, fontSize: 44, bold: true, color: P.white, fontFace: FONT, margin: 0 });
  reg('h1', 0.9, 2.2, 8.8, 1.0);
  s.addText('รากฐานของข้อมูลที่เชื่อถือได้', { x: 0.9, y: 3.22, w: 8.8, h: 0.6, fontSize: 24, color: 'C3D9EF', fontFace: FONT, margin: 0 });
  reg('h2', 0.9, 3.22, 8.8, 0.6, 'รากฐานของข้อมูลที่เชื่อถือได้', 24, { pad: 0.05 });
  s.addText('3 ชั่วโมง รวมทำแบบฝึกหัด · ไม่ต้องมีพื้นฐานมาก่อน', { x: 0.9, y: 3.95, w: 8.8, h: 0.4, fontSize: 15, color: '8FB0D4', fontFace: FONT, margin: 0 });
  reg('h3', 0.9, 3.95, 8.8, 0.4);
  s.addText('อาจารย์ ดร.เนติยา การะเกตุ  ·  หลักสูตรวิทยาศาสตร์การเกษตร  ·  ม.มหิดล วิทยาเขตกาญจนบุรี',
    { x: 0.9, y: 5.9, w: 10, h: 0.4, fontSize: 13, color: '6C8AAD', fontFace: FONT, margin: 0 });
  reg('by', 0.9, 5.9, 10, 0.4);
  checkOverlaps();
  s.addNotes('เปิดคาบด้วยการบอกว่าคาบนี้ไม่ได้สอนให้จำสูตรเยอะที่สุด แต่สอนวินัยการจัดข้อมูล เพราะความผิดพลาดส่วนใหญ่เกิดตั้งแต่ตอนป้อนข้อมูล ไม่ใช่ตอนวิเคราะห์');
}

/* 2 แนวคิดหลัก */
{
  beginSlide('แนวคิดหลัก');
  const s = lightSlide();
  const y = head(s, null, 'แนวคิดหลักของระดับนี้');
  card(s, 0.62, y + 0.15, W - 1.24, 1.9, P.soft);
  s.addText('ข้อมูลสะอาด คือจุดเริ่มต้นของงานที่เชื่อถือได้', { x: 1.0, y: y + 0.45, w: W - 2.0, h: 0.66, fontSize: 29, bold: true, color: P.deep, fontFace: FONT, margin: 0, align: 'center' });
  reg('big', 1.0, y + 0.45, W - 2.0, 0.66);
  s.addText('เหมือนงานในห้องปฏิบัติการ ถ้าเตรียมตัวอย่างไม่สะอาด ผลวิเคราะห์ก็เชื่อถือไม่ได้', { x: 1.0, y: y + 1.18, w: W - 2.0, h: 0.5, fontSize: 16, color: P.muted, fontFace: FONT, margin: 0, align: 'center' });
  reg('bigsub', 1.0, y + 1.18, W - 2.0, 0.5);

  const cy = y + 2.3, cw = (W - 1.24 - 0.6) / 3, ch = 2.1;
  textCard(s, 0.62, cy, cw, ch, 'สิ่งที่คนส่วนใหญ่คิด', 'ใช้สเปรดชีตเป็นคือการจำสูตรให้ได้เยอะ ๆ และรู้ว่าปุ่มไหนอยู่ตรงไหน', P.soft);
  textCard(s, 0.62 + cw + 0.3, cy, cw, ch, 'สิ่งที่เป็นจริง', 'สูตรที่ใช้จริงในงานส่วนใหญ่มีไม่ถึงสิบตัว แต่ข้อมูลที่จัดโครงสร้างผิดทำให้ทุกสูตรใช้ไม่ได้', P.soft2, 'A06A0A');
  textCard(s, 0.62 + (cw + 0.3) * 2, cy, cw, ch, 'สิ่งที่คาบนี้เน้น', 'วินัยในการจัดโครงสร้างข้อมูลก่อน แล้วค่อยเรียนสูตร เพราะลำดับนี้ประหยัดเวลาในระยะยาวมากที่สุด', P.soft);
  checkOverlaps();
  s.addNotes('ใช้อุปมาห้องปฏิบัติการเพราะผู้เรียนเป็นนักศึกษาสายวิทยาศาสตร์ จะเข้าใจทันทีว่าทำไมขั้นเตรียมสำคัญกว่าที่คิด');
}

/* 3 CLO */
{
  beginSlide('CLO');
  const s = lightSlide();
  const y = head(s, null, 'เมื่อจบคาบนี้ คุณจะทำอะไรได้', 'ผลลัพธ์การเรียนรู้ทั้งหกข้อ วัดผลด้วยแบบฝึกหัดและแบบทดสอบหลังเรียน');
  const items = [
    ['1.1', 'อธิบายโครงสร้างตารางที่ถูกต้อง และชี้ได้ว่าตารางที่ให้มามีปัญหาตรงไหน'],
    ['1.2', 'ป้อนและจัดรูปแบบข้อมูลตามชนิดข้อมูล และแก้ปัญหาเมื่อโปรแกรมตีความชนิดผิด'],
    ['1.3', 'นำเข้าไฟล์ .txt และ .csv ภาษาไทยได้โดยตัวอักษรไม่เพี้ยน'],
    ['1.4', 'ทำความสะอาดข้อมูลด้วย Remove Duplicates, TRIM, Find & Replace และ Conditional Formatting'],
    ['1.5', 'ใช้ SUM, AVERAGE, COUNT, COUNTIF, SUMIF และอธิบายการอ้างอิงเซลล์แบบสัมบูรณ์ได้'],
    ['1.6', 'เรียงลำดับ กรองข้อมูล และเลือกชนิดกราฟให้เหมาะกับคำถามที่ต้องการตอบ'],
  ];
  const rh = 0.74, gap = 0.15;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2), cw = (W - 1.24 - 0.4) / 2;
    const x = 0.62 + col * (cw + 0.4), yy = y + 0.1 + row * (rh + gap);
    card(s, x, yy, cw, rh, i === 4 ? P.soft2 : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.16, y: yy + 0.17, w: 0.62, h: 0.4, rectRadius: 0.06, fill: { color: i === 4 ? P.amber : P.blue }, line: { width: 0 } });
    s.addText(it[0], { x: x + 0.16, y: yy + 0.17, w: 0.62, h: 0.4, fontSize: 13, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('cloNum', x + 0.16, yy + 0.17, 0.62, 0.4);
    s.addText(it[1], { x: x + 0.92, y: yy + 0.08, w: cw - 1.1, h: rh - 0.16, fontSize: 13, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('cloText', x + 0.92, yy + 0.08, cw - 1.1, rh - 0.16, it[1], 13, { pad: 0.02 });
  });
  s.addText('ข้อ 1.5 เรื่องการอ้างอิงแบบสัมบูรณ์ เป็นข้อชี้ขาดในแบบทดสอบจัดระดับ ผู้ที่ทำข้อนี้ไม่ได้จะถูกจัดให้เริ่มที่ระดับนี้', {
    x: 0.62, y: y + 0.1 + 3 * (rh + gap) + 0.12, w: W - 1.24, h: 0.42, fontSize: 13, italic: true, color: 'A06A0A', fontFace: FONT, margin: 0
  });
  reg('note', 0.62, y + 0.1 + 3 * (rh + gap) + 0.12, W - 1.24, 0.42);
  checkOverlaps();
  s.addNotes('บอกผู้เรียนว่าหกข้อนี้คือสิ่งที่จะถูกวัดในแบบทดสอบหลังเรียน ไม่มีอะไรนอกเหนือจากนี้ จะได้ไม่ต้องกังวล');
}

/* 4 ผังเวลา */
{
  beginSlide('ผังเวลา');
  const s = lightSlide();
  const y = head(s, null, 'ผังการเรียน 180 นาที');
  const rows = [
    ['0–10', 'แบบทดสอบก่อนเรียน', 'ทำโดยไม่เปิดเอกสาร ตอบไม่ได้ไม่เป็นไร', P.muted],
    ['10–25', 'กายวิภาคของสเปรดชีตและตารางที่ดี', 'รู้จักส่วนประกอบ และหลักหนึ่งแถวหนึ่งหน่วยสังเกต', P.blue],
    ['25–55', 'การป้อนข้อมูลและชนิดข้อมูล', 'ตัวเลข ข้อความ วันที่ และกับดักวันที่แบบไทย', P.blue],
    ['55–75', 'นำเข้าไฟล์ .txt และ .csv', 'แก้ปัญหาภาษาไทยเพี้ยนด้วยการเลือก UTF-8', P.blue],
    ['75–85', 'พักเบรก', '', P.muted],
    ['85–115', 'ทำความสะอาดข้อมูล', 'ลบซ้ำ ตัดช่องว่าง และหาค่าผิดปกติ', P.amber],
    ['115–145', 'ฟังก์ชันพื้นฐานและการอ้างอิงเซลล์', 'SUM ถึง SUMIF และเครื่องหมาย $', P.blue],
    ['145–170', 'เรียง กรอง สร้างกราฟ และแบบฝึกหัด', 'เลือกชนิดกราฟให้ตรงกับคำถาม', P.blue],
    ['170–180', 'แบบทดสอบหลังเรียนและสรุป', '', P.muted],
  ];
  rows.forEach((r, i) => timeRow(s, 0.62, y + 0.08 + i * 0.55, W - 1.24, r[0], r[1], r[2], r[3]));
  checkOverlaps();
  s.addNotes('ย้ำว่าทุกช่วงมีการลงมือทำสลับกับการฟัง ไม่มีช่วงไหนที่ต้องนั่งฟังยาวเกินสิบนาที');
}

/* 5 กายวิภาค */
{
  beginSlide('กายวิภาค');
  const s = lightSlide();
  const y = head(s, 1, 'รู้จักส่วนประกอบก่อนลงมือ', 'ศัพท์สี่คำนี้จะใช้ตลอดทั้งสามระดับ');
  const parts = [
    ['Workbook', 'ไฟล์ทั้งไฟล์', 'หนึ่งไฟล์คือหนึ่งเวิร์กบุ๊ก', P.deep],
    ['Sheet', 'แผ่นงานในไฟล์', 'หนึ่งไฟล์มีได้หลายแผ่น แท็บอยู่ด้านล่าง', P.blue],
    ['Cell', 'ช่องเดียว', 'มีที่อยู่เป็นคอลัมน์กับแถว เช่น C2', P.sky],
    ['Range', 'กลุ่มของช่อง', 'เขียนเป็นช่วง เช่น A2:D100', P.amber],
  ];
  const cw = (W - 1.24 - 0.45) / 4, ch = 2.4;
  parts.forEach((p, i) => {
    const x = 0.62 + i * (cw + 0.15);
    card(s, x, y, cw, ch, P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.16, y: y + 0.22, w: cw - 0.32, h: 0.5, rectRadius: 0.07, fill: { color: p[3] }, line: { width: 0 } });
    s.addText(p[0], { x: x + 0.16, y: y + 0.22, w: cw - 0.32, h: 0.5, fontSize: 15, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('pn', x + 0.16, y + 0.22, cw - 0.32, 0.5);
    s.addText(p[1], { x: x + 0.16, y: y + 0.84, w: cw - 0.32, h: 0.44, fontSize: 15, bold: true, color: P.deep, fontFace: FONT, align: 'center', margin: 0 });
    reg('pt', x + 0.16, y + 0.84, cw - 0.32, 0.44, p[1], 15, { pad: 0.02 });
    s.addText(p[2], { x: x + 0.18, y: y + 1.34, w: cw - 0.36, h: 0.9, fontSize: 12.5, color: P.ink70, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('pd', x + 0.18, y + 1.34, cw - 0.36, 0.9, p[2], 12.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 2.6, W - 1.24, 1.05, P.soft2);
  s.addText('ช่องชื่อ (Name Box) มุมซ้ายบนบอกว่าตอนนี้คุณอยู่เซลล์ไหน พิมพ์ที่อยู่ลงไปแล้วกด Enter จะกระโดดไปเซลล์นั้นทันที มีประโยชน์มากเมื่อข้อมูลยาวเป็นพันแถว', {
    x: 0.9, y: y + 2.76, w: W - 1.8, h: 0.72, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('nb', 0.9, y + 2.76, W - 1.8, 0.72, 'ช่องชื่อ (Name Box) มุมซ้ายบนบอกว่าตอนนี้คุณอยู่เซลล์ไหน พิมพ์ที่อยู่ลงไปแล้วกด Enter จะกระโดดไปเซลล์นั้นทันที มีประโยชน์มากเมื่อข้อมูลยาวเป็นพันแถว', 13.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('สาธิตการพิมพ์ A1000 ในช่องชื่อแล้วกด Enter ผู้เรียนหลายคนไม่เคยรู้ว่าทำได้ และประหยัดเวลาเลื่อนเมาส์มาก');
}

/* 6 ตารางที่ดี */
{
  beginSlide('ตารางที่ดี');
  const s = lightSlide();
  const y = head(s, 1, 'ตารางที่ดีหน้าตาเป็นอย่างไร', 'กฎสามข้อนี้ตัดสินว่าคุณจะวิเคราะห์ข้อมูลต่อได้หรือไม่');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.3, P.soft);
  s.addText('ตารางที่วิเคราะห์ต่อได้', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('g1', 0.86, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86, y + 0.68, cw - 0.48, 2.45, [
    'หนึ่งแถวคือหน่วยสังเกตหนึ่งหน่วย เช่น หนึ่งแปลง หนึ่งตัวอย่าง',
    'หนึ่งคอลัมน์คือตัวแปรหนึ่งตัว เก็บชนิดข้อมูลเดียวกันทั้งคอลัมน์',
    'หัวตารางมีแถวเดียว อยู่บนสุด และไม่มีเซลล์ว่างคั่น',
  ], 13.5);

  card(s, 0.62 + cw + 0.4, y, cw, 3.3, P.warn);
  s.addText('ตารางที่ดูสวยแต่ใช้ต่อไม่ได้', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('g2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86 + cw + 0.4, y + 0.68, cw - 0.48, 2.45, [
    'หัวตารางซ้อนสองแถวและรวมเซลล์ให้ดูเป็นระเบียบ',
    'รวมข้อมูลสองอย่างไว้เซลล์เดียว เช่น ชื่อพืชกับวันที่',
    'แยกข้อมูลแต่ละสัปดาห์ไว้คนละชีต',
    'มีแถวว่างคั่นระหว่างกลุ่มเพื่อให้อ่านง่าย',
  ], 13.5);

  card(s, 0.62, y + 3.5, W - 1.24, 0.85, P.soft2);
  s.addText('ทั้งสี่อย่างทางขวาทำให้ PivotTable และสูตรทำงานไม่ได้ในระดับถัดไป จึงต้องเลิกทำตั้งแต่วันนี้', {
    x: 0.9, y: y + 3.64, w: W - 1.8, h: 0.58, fontSize: 14, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('g3', 0.9, y + 3.64, W - 1.8, 0.58, 'ทั้งสี่อย่างทางขวาทำให้ PivotTable และสูตรทำงานไม่ได้ในระดับถัดไป จึงต้องเลิกทำตั้งแต่วันนี้', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('บอกตรง ๆ ว่าสิ่งที่อยู่ทางขวาคือสิ่งที่พวกเขาน่าจะเคยทำมาก่อน ไม่ใช่เรื่องผิด แต่ต้องเปลี่ยนถ้าจะวิเคราะห์ต่อ');
}

/* ภาพเปรียบเทียบ ตารางที่ใช้ต่อไม่ได้กับตารางที่ควรเป็น */
{
  const s = mockPair('ภาพตารางดี-ไม่ดี', 1,
    'เห็นภาพจริงว่าต่างกันอย่างไร',
    'ข้อมูลชุดเดียวกัน จัดสองแบบ แบบขวาคือแบบที่วิเคราะห์ต่อได้',
    'แบบที่ใช้ต่อไม่ได้', 'แบบที่ควรเป็น',
    { cols: ['ชื่อพืชและวันที่', 'ความสูง'],
      rows: [['ข้าวโพด 03/08/68', '12.4 ซม.'], ['ข้าวโพด 10/08/68', '14.1 ซม.'],
             [['', 'x'], ['เว้นแถวให้อ่านง่าย', 'x']], ['ถั่วเขียว 03/08/68', '8.2 ซม.']] },
    { cols: ['ชนิดพืช', 'วันที่', 'ความสูง_ซม'],
      rows: [['ข้าวโพด', '2025-08-03', ['12.4', 'n']], ['ข้าวโพด', '2025-08-10', ['14.1', 'n']],
             ['ถั่วเขียว', '2025-08-03', ['8.2', 'n']], ['ถั่วเขียว', '2025-08-10', ['9.6', 'n']]] },
    'ฝั่งซ้ายรวมชื่อพืชกับวันที่ไว้เซลล์เดียว ใส่หน่วยในเซลล์ และเว้นแถวคั่น ทั้งสามอย่างทำให้กรองและสรุปไม่ได้ ส่วนฝั่งขวาเก็บหน่วยไว้ในชื่อคอลัมน์แทน');
  s.addNotes('ให้เวลาดูเงียบ ๆ สิบวินาทีก่อนอธิบาย แล้วถามว่าเห็นความต่างกี่จุด ส่วนใหญ่จะเห็นสองจุดจากสามจุด');
}

/* 7 ชนิดข้อมูล */
{
  beginSlide('ชนิดข้อมูล');
  const s = lightSlide();
  const y = head(s, 2, 'ชนิดข้อมูลสามชนิดที่ต้องแยกให้ออก', 'สังเกตการจัดชิดของเซลล์ แล้วจะรู้ทันทีว่าโปรแกรมมองข้อมูลนั้นเป็นอะไร');
  const rows = [
    ['ตัวเลข', 'ชิดขวาโดยอัตโนมัติ', 'คำนวณได้ ใช้กับ SUM และ AVERAGE ได้', P.blue],
    ['ข้อความ', 'ชิดซ้ายโดยอัตโนมัติ', 'คำนวณไม่ได้ SUM จะข้ามไปเฉย ๆ ไม่แจ้งเตือน', P.coral],
    ['วันที่', 'ชิดขวา เพราะแท้จริงคือตัวเลข', 'เก็บเป็นจำนวนวันนับจากวันตั้งต้น จึงบวกลบได้', P.amber],
  ];
  const rh = 0.95, gap = 0.16;
  rows.forEach((r, i) => {
    const yy = y + 0.1 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 1 ? P.warn : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.24, w: 2.1, h: 0.48, rectRadius: 0.07, fill: { color: r[3] }, line: { width: 0 } });
    s.addText(r[0], { x: 0.86, y: yy + 0.24, w: 2.1, h: 0.48, fontSize: 15, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('dk', 0.86, yy + 0.24, 2.1, 0.48);
    s.addText(r[1], { x: 3.2, y: yy + 0.12, w: 3.5, h: rh - 0.24, fontSize: 14, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('dd', 3.2, yy + 0.12, 3.5, rh - 0.24, r[1], 14, { pad: 0.02 });
    s.addText(r[2], { x: 6.9, y: yy + 0.12, w: W - 0.62 - 6.9 - 0.26, h: rh - 0.24, fontSize: 13, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('dv', 6.9, yy + 0.12, W - 0.62 - 6.9 - 0.26, rh - 0.24, r[2], 13, { pad: 0.02 });
  });
  card(s, 0.62, y + 0.1 + 3 * (rh + gap) + 0.05, W - 1.24, 0.95, P.soft2);
  s.addText('อาการที่พบบ่อยที่สุด: คัดลอกตัวเลขจากเว็บหรือ PDF มาวาง แล้ว SUM ได้ผลเป็นศูนย์ เพราะตัวเลขนั้นถูกเก็บเป็นข้อความ', {
    x: 0.9, y: y + 0.1 + 3 * (rh + gap) + 0.2, w: W - 1.8, h: 0.66, fontSize: 14, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('dnote', 0.9, y + 0.1 + 3 * (rh + gap) + 0.2, W - 1.8, 0.66, 'อาการที่พบบ่อยที่สุด: คัดลอกตัวเลขจากเว็บหรือ PDF มาวาง แล้ว SUM ได้ผลเป็นศูนย์ เพราะตัวเลขนั้นถูกเก็บเป็นข้อความ', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('สาธิตสดโดยพิมพ์เลขปกติหนึ่งช่อง แล้วพิมพ์เลขที่มีช่องว่างนำหน้าอีกช่อง ให้เห็นว่าชิดคนละด้าน แล้วลอง SUM ทั้งสอง');
}

/* ภาพตัวเลขที่ถูกเก็บเป็นข้อความ */
{
  const s = mockPair('ภาพตัวเลขเป็นข้อความ', 1,
    'ตัวเลขที่บวกไม่ได้หน้าตาเป็นอย่างไร',
    'สังเกตการชิดซ้ายชิดขวา นี่คือวิธีตรวจที่เร็วที่สุดและไม่ต้องใช้เครื่องมือใด',
    'ตัวเลขที่เก็บเป็นข้อความ จึงชิดซ้าย', 'ตัวเลขจริง จึงชิดขวา',
    { cols: ['แปลง', 'ผลผลิต (กก.)'],
      rows: [['AG-001', ['1,250', 'x']], ['AG-002', ['980', 'x']], ['AG-003', ['1,180', 'x']],
             [['รวม', 'x'], ['0', 'x']]] },
    { cols: ['แปลง', 'ผลผลิต (กก.)'],
      rows: [['AG-001', ['1250', 'n']], ['AG-002', ['980', 'n']], ['AG-003', ['1180', 'n']],
             ['รวม', ['3410', 'n']]] },
    'SUM ของฝั่งซ้ายได้ศูนย์ เพราะโปรแกรมมองว่าเป็นข้อความไม่ใช่ตัวเลข และไม่มีข้อความผิดพลาดใดขึ้นเตือน ผลรวมที่หายไปจึงไหลเข้าไปอยู่ในรายงานได้');
  s.addNotes('ย้ำว่าเครื่องหมายคั่นหลักพันที่พิมพ์เองคือสาเหตุที่พบบ่อยที่สุด ให้ใช้การจัดรูปแบบตัวเลขแทนการพิมพ์จุลภาค');
}

/* 8 กับดักวันที่ */
{
  beginSlide('กับดักวันที่');
  const s = lightSlide();
  const y = head(s, 2, 'กับดักวันที่แบบไทย', 'ปัญหาที่เจอในไฟล์ฝึกปฏิบัติทุกสาขา');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.15, P.warn);
  s.addText('ปัญหา', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('t1', 0.86, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86, y + 0.68, cw - 0.48, 2.3, [
    'วันที่แบบ พ.ศ. ที่พิมพ์เอง โปรแกรมมองเป็นข้อความ',
    'เรียงลำดับแล้วผิด เพราะเรียงแบบข้อความ',
    'แยกเดือนหรือปีออกมาไม่ได้',
    'คำนวณจำนวนวันระหว่างสองวันที่ไม่ได้'
  ], 13.5);

  card(s, 0.62 + cw + 0.4, y, cw, 3.15, P.soft);
  s.addText('วิธีสังเกตและแก้', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('t2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86 + cw + 0.4, y + 0.68, cw - 0.48, 2.3, [
    'ถ้าวันที่ชิดซ้าย แปลว่าเป็นข้อความ ต้องแก้',
    'ถ้าวันที่กลายเป็นเลขห้าหลัก แปลว่าเป็นวันที่จริง แค่จัดรูปแบบใหม่',
    'ใช้ Text to Columns แปลงข้อความเป็นวันที่ทั้งคอลัมน์',
    'ตกลงรูปแบบวันที่กับทีมให้ตรงกันตั้งแต่ต้น'
  ], 13.5);

  card(s, 0.62, y + 3.35, W - 1.24, 0.85, P.soft2);
  s.addText('เลขห้าหลักที่โผล่มาแทนวันที่ไม่ใช่ความผิดพลาด แต่คือวิธีที่โปรแกรมเก็บวันที่จริง ๆ คือนับเป็นจำนวนวันจากวันตั้งต้น', {
    x: 0.9, y: y + 3.5, w: W - 1.8, h: 0.58, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('t3', 0.9, y + 3.5, W - 1.8, 0.58, 'เลขห้าหลักที่โผล่มาแทนวันที่ไม่ใช่ความผิดพลาด แต่คือวิธีที่โปรแกรมเก็บวันที่จริง ๆ คือนับเป็นจำนวนวันจากวันตั้งต้น', 13.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนเปิดไฟล์สาขาตนเองแล้วหาคอลัมน์วันที่ จะเห็นว่ามีทั้งแบบชิดซ้ายและชิดขวาปนกันอยู่จริง');
}

/* 9 WS1.1 */
{
  beginSlide('WS1.1');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 1.1 · สร้างตารางบันทึกผล', 'ใช้เวลา 15 นาที ทำคนเดียว');
  card(s, 0.62, y, W - 1.24, 2.2, P.soft);
  s.addText('โจทย์', { x: 0.9, y: y + 0.18, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w1', 0.9, y + 0.18, W - 1.8, 0.36);
  bullets(s, 0.9, y + 0.6, W - 1.8, 1.45, [
    'ออกแบบและป้อนตารางบันทึกการเจริญเติบโตของพืช 10 ต้น × 4 สัปดาห์',
    'ตั้งชื่อคอลัมน์เอง และกำหนดว่าคอลัมน์ไหนเก็บชนิดข้อมูลอะไร',
    'ใส่วันที่บันทึกให้เป็นวันที่จริง ไม่ใช่ข้อความ',
  ], 13.5);
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.4, cw, 1.55, 'เกณฑ์ผ่าน',
    'โครงสร้างถูกต้องตามกฎสามข้อ วันที่เป็นชนิดวันที่จริง และไม่มีการรวมเซลล์กลางตาราง', P.soft, P.deep);
  textCard(s, 0.62 + cw + 0.4, y + 2.4, cw, 1.55, 'คำถามที่ผู้สอนจะถาม',
    'ถ้าสัปดาห์หน้าต้องเพิ่มข้อมูลอีก 10 ต้น ตารางของคุณรองรับได้ทันทีหรือต้องรื้อใหม่', P.soft2, 'A06A0A');
  checkOverlaps();
  s.addNotes('คำถามในการ์ดขวาสำคัญมาก เพราะเป็นตัวชี้ว่าผู้เรียนออกแบบตารางแบบขยายได้หรือแบบตายตัว ให้เดินถามทีละคน');
}

/* 10 นำเข้าไฟล์ */
{
  beginSlide('นำเข้าไฟล์');
  const s = lightSlide();
  const y = head(s, 3, 'เปิดไฟล์ .csv ภาษาไทยไม่ให้เพี้ยน', 'ปัญหาที่ทำให้เสียเวลามากที่สุดของงานข้อมูลไทย');
  card(s, 0.62, y, W - 1.24, 1.4, P.warn);
  s.addText('อาการ', { x: 0.9, y: y + 0.16, w: W - 1.8, h: 0.34, fontSize: 15, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('i1', 0.9, y + 0.16, W - 1.8, 0.34);
  s.addText('ดับเบิลคลิกเปิดไฟล์ .csv แล้วภาษาไทยกลายเป็นอักขระอ่านไม่ออก หลายคนคิดว่าไฟล์เสียแล้วขอไฟล์ใหม่ ทั้งที่ไฟล์ไม่ได้เสียเลย', {
    x: 0.9, y: y + 0.54, w: W - 1.8, h: 0.72, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('i2', 0.9, y + 0.54, W - 1.8, 0.72, 'ดับเบิลคลิกเปิดไฟล์ .csv แล้วภาษาไทยกลายเป็นอักขระอ่านไม่ออก หลายคนคิดว่าไฟล์เสียแล้วขอไฟล์ใหม่ ทั้งที่ไฟล์ไม่ได้เสียเลย', 13.5, { pad: 0.02 });

  const steps = [
    ['ห้ามดับเบิลคลิก', 'เปิดโปรแกรมขึ้นมาก่อน แล้วค่อยนำเข้าไฟล์'],
    ['ไปที่เมนู Data', 'เลือก From Text/CSV ไม่ใช่ File แล้ว Open'],
    ['เลือกการเข้ารหัส UTF-8', 'ดูตัวอย่างในหน้าต่างว่าอ่านออกแล้วจึงกดต่อ'],
    ['ตรวจตัวคั่นให้ถูก', 'ส่วนใหญ่เป็นจุลภาค แต่บางไฟล์ใช้แท็บหรือเซมิโคลอน'],
  ];
  const bw = (W - 1.24 - 0.45) / 4;
  steps.forEach((st, i) => {
    const x = 0.62 + i * (bw + 0.15);
    card(s, x, y + 1.6, bw, 1.85, P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: x + (bw - 0.54) / 2, y: y + 1.78, w: 0.54, h: 0.54, fill: { color: P.blue }, line: { width: 0 } });
    s.addText(String(i + 1), { x: x + (bw - 0.54) / 2, y: y + 1.78, w: 0.54, h: 0.54, fontSize: 18, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sn', x + (bw - 0.54) / 2, y + 1.78, 0.54, 0.54);
    s.addText(st[0], { x: x + 0.14, y: y + 2.42, w: bw - 0.28, h: 0.42, fontSize: 13.5, bold: true, color: P.deep, fontFace: FONT, align: 'center', margin: 0 });
    reg('sname', x + 0.14, y + 2.42, bw - 0.28, 0.42, st[0], 13.5, { pad: 0.02 });
    s.addText(st[1], { x: x + 0.16, y: y + 2.86, w: bw - 0.32, h: 0.52, fontSize: 12, color: P.ink70, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('sdesc', x + 0.16, y + 2.86, bw - 0.32, 0.52, st[1], 12, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('เตรียมไฟล์ .csv ภาษาไทยไว้ให้ผู้เรียนลองเปิดผิดวิธีก่อนหนึ่งครั้ง ให้เห็นอาการจริง แล้วค่อยสอนวิธีที่ถูก จะจำได้แม่นกว่า');
}

/* 11 เบรก */
{
  beginSlide('เบรก');
  const s = darkSlide();
  s.addShape(pres.ShapeType.ellipse, { x: -1.2, y: 4.4, w: 4.6, h: 4.6, fill: { color: P.deep }, line: { width: 0 } });
  s.addText('พักเบรก 10 นาที', { x: 0.9, y: 2.6, w: 11.5, h: 0.9, fontSize: 38, bold: true, color: P.white, fontFace: FONT, margin: 0, align: 'center' });
  reg('brk', 0.9, 2.6, 11.5, 0.9);
  s.addText('กลับมาแล้วเราจะเริ่มทำความสะอาดข้อมูล ซึ่งเป็นหัวใจของคาบนี้', {
    x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 17, color: 'C3D9EF', fontFace: FONT, margin: 0, align: 'center'
  });
  reg('brk2', 0.9, 3.6, 11.5, 0.5, 'กลับมาแล้วเราจะเริ่มทำความสะอาดข้อมูล ซึ่งเป็นหัวใจของคาบนี้', 17, { pad: 0.05 });
  checkOverlaps();
  s.addNotes('ใช้ช่วงเบรกช่วยคนที่ยังนำเข้าไฟล์ไม่สำเร็จ เพราะถ้าเปิดไฟล์ไม่ได้จะทำช่วงหลังไม่ได้เลย');
}

/* 12 เครื่องมือทำความสะอาด */
{
  beginSlide('ทำความสะอาด');
  const s = lightSlide();
  const y = head(s, 4, 'สี่เครื่องมือทำความสะอาดข้อมูล', 'ใช้ตามลำดับนี้ทุกครั้ง แล้วจะไม่พลาด');
  const rows = [
    ['1. Remove Duplicates', 'ลบแถวที่ซ้ำกันทั้งแถว', 'ทำก่อนเสมอ เพราะแถวซ้ำทำให้ผลนับและผลรวมเกินจริง', P.blue],
    ['2. TRIM', 'ตัดช่องว่างหน้าและหลังข้อความ', 'ช่องว่างที่มองไม่เห็นทำให้ชื่อเดียวกันกลายเป็นคนละชื่อ', P.blue],
    ['3. Find & Replace', 'แก้คำที่สะกดไม่ตรงกันให้เป็นแบบเดียว', 'ทำหลัง TRIM เพราะบางคำต่างกันแค่ช่องว่าง', P.blue],
    ['4. Conditional Formatting', 'ระบายสีค่าที่ผิดปกติให้เห็นด้วยตา', 'ใช้หาค่าที่เป็นไปไม่ได้ เช่น ค่าติดลบหรือศูนย์', P.amber],
  ];
  const rh = 0.85, gap = 0.14;
  rows.forEach((r, i) => {
    const yy = y + 0.08 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 3 ? P.soft2 : P.soft);
    s.addText(r[0], { x: 0.88, y: yy + 0.1, w: 3.3, h: rh - 0.2, fontSize: 14, bold: true, color: r[3], fontFace: FONT, valign: 'middle', margin: 0 });
    reg('ck', 0.88, yy + 0.1, 3.3, rh - 0.2, r[0], 14, { pad: 0.02 });
    s.addText(r[1], { x: 4.32, y: yy + 0.1, w: 3.5, h: rh - 0.2, fontSize: 13, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('cd', 4.32, yy + 0.1, 3.5, rh - 0.2, r[1], 13, { pad: 0.02 });
    s.addText(r[2], { x: 7.98, y: yy + 0.1, w: W - 0.62 - 7.98 - 0.26, h: rh - 0.2, fontSize: 12.5, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('cv', 7.98, yy + 0.1, W - 0.62 - 7.98 - 0.26, rh - 0.2, r[2], 12.5, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('ย้ำเรื่องลำดับ ถ้าทำ Find & Replace ก่อน TRIM จะพลาดคำที่ต่างกันแค่ช่องว่าง เป็นบทเรียนเรื่องลำดับงานที่จะเจออีกในระดับถัดไป');
}

/* 13 ตัวอย่างจริง AG */
{
  beginSlide('ตัวอย่าง AG');
  const s = lightSlide();
  const y = head(s, 4, 'ช่องว่างที่มองไม่เห็นทำให้กลุ่มแตก', 'ตัวอย่างจริงจากไฟล์ฝึกปฏิบัติสาขาวิทยาศาสตร์การเกษตร');
  card(s, 0.62, y, W - 1.24, 2.45, P.soft);
  s.addText('ไฟล์แปลงทดลองข้าวโพดเก็บข้อมูลจาก 6 ตำบล แต่เมื่อนับด้วย COUNTIF กลับได้', {
    x: 0.9, y: y + 0.22, w: W - 1.8, h: 0.4, fontSize: 15, bold: true, color: P.deep, fontFace: FONT, margin: 0
  });
  reg('e1', 0.9, y + 0.22, W - 1.8, 0.4);
  const sw = (W - 1.8 - 1.0) / 3;
  stat(s, 0.9, y + 0.7, sw, '10', 'ชื่อตำบลที่ปรากฏในข้อมูลดิบ', P.coral);
  stat(s, 0.9 + sw + 0.5, y + 0.7, sw, '6', 'ตำบลจริงหลังตัดช่องว่างและจุด', P.blue);
  stat(s, 0.9 + (sw + 0.5) * 2, y + 0.7, sw, '9', 'แถวซ้ำทั้งแถวที่ต้องลบก่อนนับ', P.amber);

  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.6, cw, 1.65, 'สาเหตุ',
    'ผู้บันทึกสี่คนกรอกชื่อตำบลต่างกัน บางคนมีช่องว่างท้าย บางคนใส่จุด ตาเรามองว่าเหมือนกัน แต่โปรแกรมมองว่าคนละคำ', P.warn, 'A83A36');
  textCard(s, 0.62 + cw + 0.4, y + 2.6, cw, 1.65, 'ทางแก้',
    'ใช้ TRIM ตัดช่องว่างทั้งคอลัมน์ก่อน แล้วค่อยนับ ถ้าจำเป็นให้ใช้ Find & Replace เก็บกรณีที่เหลือ', P.soft, P.deep);
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนเปิดไฟล์สาขาตนเองแล้วใช้ COUNTIF นับชื่อกลุ่มดู จะเจอปรากฏการณ์เดียวกันในทุกไฟล์ เป็นช่วงที่ผู้เรียนตื่นเต้นที่สุด');
}

/* 14 ค่าที่ดูแปลก */
{
  beginSlide('ค่าที่ดูแปลก');
  const s = lightSlide();
  const y = head(s, 4, 'ค่าที่ดูแปลก ต้องคิดก่อนลบ', 'บางค่าที่ดูผิดคือข้อมูลที่ถูกต้องตามธรรมเนียมของสาขา');
  const rows = [
    ['ค่า -999', 'ไม่ใช่ค่าที่วัดได้จริง แต่เป็นรหัสแทนช่วงที่เครื่องมือไม่ทำงาน ต้องเปลี่ยนเป็นค่าว่างก่อนคำนวณ', P.coral],
    ['ค่า 0 ที่เป็นไปไม่ได้', 'เช่น ความชื้นเมล็ดเท่ากับศูนย์ ซึ่งเป็นไปไม่ได้ในทางกายภาพ ต้องตรวจย้อนกลับว่าเกิดจากอะไร', P.amber],
    ['จำนวนที่กรอกเป็นคำ', 'เช่น หลายตัว หรือ ไม่แน่ชัด เป็นข้อมูลจริงที่ผู้บันทึกตั้งใจเขียน ห้ามลบทิ้งเงียบ ๆ ให้แยกออกมานับต่างหาก', P.blue],
    ['จำนวนติดลบ', 'ในข้อมูลการขายคือรายการคืนสินค้า ซึ่งถูกต้องและต้องนับ ไม่ใช่ความผิดพลาด', P.blue],
  ];
  const rh = 0.85, gap = 0.15;
  rows.forEach((r, i) => {
    const yy = y + 0.08 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i < 2 ? P.warn : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.19, w: 0.16, h: 0.47, rectRadius: 0.05, fill: { color: r[2] }, line: { width: 0 } });
    reg('dot', 0.86, yy + 0.19, 0.16, 0.47);
    s.addText(r[0], { x: 1.2, y: yy + 0.1, w: 3.3, h: rh - 0.2, fontSize: 14, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('vk', 1.2, yy + 0.1, 3.3, rh - 0.2, r[0], 14, { pad: 0.02 });
    s.addText(r[1], { x: 4.64, y: yy + 0.1, w: W - 0.62 - 4.64 - 0.26, h: rh - 0.2, fontSize: 12.5, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('vv', 4.64, yy + 0.1, W - 0.62 - 4.64 - 0.26, rh - 0.2, r[1], 12.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 0.08 + 4 * (rh + gap) + 0.02, W - 1.24, 0.8, P.soft2);
  s.addText('กฎเหล็ก: ทุกครั้งที่ลบหรือแก้ข้อมูล ให้บันทึกไว้ว่าทำอะไรและเพราะอะไร ในชีตชื่อ บันทึกการตัดสินใจ', {
    x: 0.9, y: y + 0.08 + 4 * (rh + gap) + 0.14, w: W - 1.8, h: 0.56, fontSize: 14, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('rule', 0.9, y + 0.08 + 4 * (rh + gap) + 0.14, W - 1.8, 0.56, 'กฎเหล็ก: ทุกครั้งที่ลบหรือแก้ข้อมูล ให้บันทึกไว้ว่าทำอะไรและเพราะอะไร ในชีตชื่อ บันทึกการตัดสินใจ', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('นี่คือสไลด์ที่แยกคาบนี้ออกจากคอร์ส Excel ทั่วไป ให้เวลาอภิปรายสักห้านาที ถามว่าในสาขาของแต่ละคนมีค่าแบบไหนที่ดูแปลกแต่ถูก');
}

/* 15 WS1.2 */
{
  beginSlide('WS1.2');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 1.2 · กู้ชีพไฟล์ข้อมูลที่พัง', 'ใช้เวลา 25 นาที ทำกับไฟล์ของสาขาตนเอง');
  card(s, 0.62, y, W - 1.24, 2.3, P.soft);
  s.addText('สิ่งที่ต้องทำ', { x: 0.9, y: y + 0.2, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w1', 0.9, y + 0.2, W - 1.8, 0.36);
  bullets(s, 0.9, y + 0.62, W - 1.8, 1.55, [
    'เปิดไฟล์ฝึกปฏิบัติของสาขาตนเอง อ่านชีตอ่านก่อนเริ่มให้จบก่อน',
    'สำรวจให้ครบว่าไฟล์มีปัญหาอะไรบ้าง ก่อนลงมือแก้แม้แต่จุดเดียว',
    'ใช้เครื่องมือทั้งสี่ตามลำดับ แล้วบันทึกทุกการตัดสินใจพร้อมเหตุผล',
  ], 13.5);
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.5, cw, 1.5, 'เกณฑ์ผ่าน',
    'ข้อมูลอ่านภาษาไทยได้ ลบแถวซ้ำครบ ชื่อกลุ่มเป็นมาตรฐานเดียวกัน และมีชีตบันทึกการตัดสินใจ', P.soft, P.deep);
  textCard(s, 0.62 + cw + 0.4, y + 2.5, cw, 1.5, 'สิ่งที่ผู้สอนจะดู',
    'ผู้ที่รีบแก้ทีละจุดโดยไม่สำรวจภาพรวมก่อน มักแก้ผิดลำดับแล้วต้องย้อนกลับมาทำใหม่', P.soft2, 'A06A0A');
  checkOverlaps();
  s.addNotes('ห้ามบอกว่าไฟล์มีปัญหาอะไรบ้าง ปล่อยให้หาเอง ชีตอ่านก่อนเริ่มบอกไว้แล้วว่ามีกี่ประเภท ให้ใช้ตัวเลขนั้นเป็นเป้า');
}

/* 16 ฟังก์ชันพื้นฐาน */
{
  beginSlide('ฟังก์ชันพื้นฐาน');
  const s = lightSlide();
  const y = head(s, 5, 'ห้าฟังก์ชันที่ใช้ได้ครอบคลุมงานส่วนใหญ่', 'ไม่ต้องจำเยอะ ขอให้ใช้ห้าตัวนี้ให้คล่องก่อน');
  const rows = [
    ['SUM', 'บวกค่าทั้งช่วง', 'ผลรวมผลผลิตทั้งหมด', P.blue],
    ['AVERAGE', 'หาค่าเฉลี่ยของช่วง', 'ค่าเฉลี่ยของค่าที่วัดได้', P.blue],
    ['COUNT', 'นับเฉพาะช่องที่เป็นตัวเลข', 'มีข้อมูลที่ใช้คำนวณได้กี่แถว', P.blue],
    ['COUNTIF', 'นับเฉพาะช่องที่ตรงเงื่อนไข', 'แต่ละตำบลมีกี่แปลง', P.amber],
    ['SUMIF', 'บวกเฉพาะแถวที่ตรงเงื่อนไข', 'ผลผลิตรวมของแต่ละตำบล', P.amber],
  ];
  const rh = 0.72, gap = 0.13;
  rows.forEach((r, i) => {
    const yy = y + 0.05 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i >= 3 ? P.soft2 : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.13, w: 2.1, h: 0.46, rectRadius: 0.07, fill: { color: r[3] }, line: { width: 0 } });
    s.addText(r[0], { x: 0.86, y: yy + 0.13, w: 2.1, h: 0.46, fontSize: 13.5, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('fk', 0.86, yy + 0.13, 2.1, 0.46);
    s.addText(r[1], { x: 3.2, y: yy + 0.08, w: 4.2, h: rh - 0.16, fontSize: 13.5, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('fd', 3.2, yy + 0.08, 4.2, rh - 0.16, r[1], 13.5, { pad: 0.02 });
    s.addText(r[2], { x: 7.6, y: yy + 0.08, w: W - 0.62 - 7.6 - 0.26, h: rh - 0.16, fontSize: 12.5, color: P.muted, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('fv', 7.6, yy + 0.08, W - 0.62 - 7.6 - 0.26, rh - 0.16, r[2], 12.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 0.05 + 5 * (rh + gap) + 0.04, W - 1.24, 0.8, P.soft);
  s.addText('COUNTIF และ SUMIF คือสองตัวที่เปลี่ยนงานจากการนั่งนับด้วยตาเป็นการให้โปรแกรมนับให้', {
    x: 0.9, y: y + 0.05 + 5 * (rh + gap) + 0.16, w: W - 1.8, h: 0.56, fontSize: 14, bold: true, color: P.deep, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('fnote', 0.9, y + 0.05 + 5 * (rh + gap) + 0.16, W - 1.8, 0.56, 'COUNTIF และ SUMIF คือสองตัวที่เปลี่ยนงานจากการนั่งนับด้วยตาเป็นการให้โปรแกรมนับให้', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('อย่าสอนไวยากรณ์ของสูตรก่อน ให้เริ่มจากคำถามที่อยากรู้ก่อนเสมอ แล้วค่อยบอกว่าใช้ฟังก์ชันไหนตอบ');
}

/* 17 การอ้างอิงแบบสัมบูรณ์ */
{
  beginSlide('$ สัมบูรณ์');
  const s = lightSlide();
  const y = head(s, 5, 'เครื่องหมาย $ คือหัวใจของการลากสูตร', 'ข้อนี้คือข้อชี้ขาดในแบบทดสอบจัดระดับ ต้องเข้าใจให้ได้ก่อนออกจากห้อง');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 2.6, P.soft);
  s.addText('อ้างอิงแบบสัมพัทธ์', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('r1', 0.86, y + 0.2, cw - 0.48, 0.4);
  s.addText('เขียนว่า B2 เฉย ๆ เมื่อลากสูตรลง จะเลื่อนตามเป็น B3 B4 ไปเรื่อย ๆ', {
    x: 0.86, y: y + 0.66, w: cw - 0.48, h: 0.72, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('r1b', 0.86, y + 0.66, cw - 0.48, 0.72, 'เขียนว่า B2 เฉย ๆ เมื่อลากสูตรลง จะเลื่อนตามเป็น B3 B4 ไปเรื่อย ๆ', 13.5, { pad: 0.02 });
  s.addText('เหมาะกับ: ข้อมูลที่อยู่แถวเดียวกับสูตร เช่น ราคาของแต่ละแถว', {
    x: 0.86, y: y + 1.46, w: cw - 0.48, h: 0.6, fontSize: 12.5, color: P.muted, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('r1c', 0.86, y + 1.46, cw - 0.48, 0.6, 'เหมาะกับ: ข้อมูลที่อยู่แถวเดียวกับสูตร เช่น ราคาของแต่ละแถว', 12.5, { pad: 0.02 });

  card(s, 0.62 + cw + 0.4, y, cw, 2.6, P.soft2);
  s.addText('อ้างอิงแบบสัมบูรณ์', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0 });
  reg('r2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  s.addText('เขียนว่า $E$1 เมื่อลากสูตรไปไหนก็ยังชี้ที่ E1 เสมอ กด F4 เพื่อใส่ $ ให้อัตโนมัติ', {
    x: 0.86 + cw + 0.4, y: y + 0.66, w: cw - 0.48, h: 0.72, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('r2b', 0.86 + cw + 0.4, y + 0.66, cw - 0.48, 0.72, 'เขียนว่า $E$1 เมื่อลากสูตรไปไหนก็ยังชี้ที่ E1 เสมอ กด F4 เพื่อใส่ $ ให้อัตโนมัติ', 13.5, { pad: 0.02 });
  s.addText('เหมาะกับ: ค่าคงที่ที่ใช้ร่วมกันทุกแถว เช่น อัตราแลกเปลี่ยนหรือราคาต่อหน่วย', {
    x: 0.86 + cw + 0.4, y: y + 1.46, w: cw - 0.48, h: 0.6, fontSize: 12.5, color: P.muted, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('r2c', 0.86 + cw + 0.4, y + 1.46, cw - 0.48, 0.6, 'เหมาะกับ: ค่าคงที่ที่ใช้ร่วมกันทุกแถว เช่น อัตราแลกเปลี่ยนหรือราคาต่อหน่วย', 12.5, { pad: 0.02 });

  card(s, 0.62, y + 2.8, W - 1.24, 1.15, P.warn);
  s.addText('อาการเมื่อลืมใส่ $', { x: 0.9, y: y + 2.94, w: W - 1.8, h: 0.34, fontSize: 14.5, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('rw1', 0.9, y + 2.94, W - 1.8, 0.34);
  s.addText('แถวแรกถูก แถวต่อ ๆ มาผิดหมด และไม่มี error ขึ้นให้เห็น เพราะสูตรยังคำนวณได้ตามปกติ เพียงแต่ชี้ไปที่ช่องว่างแทน', {
    x: 0.9, y: y + 3.3, w: W - 1.8, h: 0.56, fontSize: 13, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('rw2', 0.9, y + 3.3, W - 1.8, 0.56, 'แถวแรกถูก แถวต่อ ๆ มาผิดหมด และไม่มี error ขึ้นให้เห็น เพราะสูตรยังคำนวณได้ตามปกติ เพียงแต่ชี้ไปที่ช่องว่างแทน', 13, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('สาธิตทั้งสองแบบข้าง ๆ กัน แล้วลากลงพร้อมกัน ให้เห็นว่าอันที่ไม่ใส่ $ ค่อย ๆ เพี้ยนลงไปเรื่อย ๆ โดยไม่มีสัญญาณเตือน');
}

/* 18 อ่าน error */
{
  beginSlide('อ่าน error');
  const s = lightSlide();
  const y = head(s, 5, 'อ่านข้อความ error ให้เป็น', 'error ไม่ใช่เรื่องน่ากลัว มันกำลังบอกว่าปัญหาอยู่ตรงไหน');
  const rows = [
    ['#VALUE!', 'ชนิดข้อมูลไม่ตรงกัน', 'มักเกิดจากเอาข้อความไปคำนวณ ให้ตรวจว่าคอลัมน์นั้นเป็นตัวเลขจริงหรือไม่'],
    ['#DIV/0!', 'หารด้วยศูนย์', 'ตัวหารเป็นศูนย์หรือเป็นช่องว่าง ให้ตรวจว่าข้อมูลครบหรือยัง'],
    ['#REF!', 'อ้างอิงเซลล์ที่ถูกลบไปแล้ว', 'เกิดหลังลบแถวหรือคอลัมน์ ต้องแก้สูตรให้ชี้ที่ใหม่'],
    ['#N/A', 'หาค่าที่ต้องการไม่พบ', 'พบมากในระดับถัดไปเมื่อใช้ฟังก์ชันค้นหา แปลว่าค่านั้นไม่มีในตารางอ้างอิง'],
  ];
  const rh = 0.82, gap = 0.15;
  rows.forEach((r, i) => {
    const yy = y + 0.1 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.17, w: 1.9, h: 0.48, rectRadius: 0.07, fill: { color: P.coral }, line: { width: 0 } });
    s.addText(r[0], { x: 0.86, y: yy + 0.17, w: 1.9, h: 0.48, fontSize: 13.5, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('ek', 0.86, yy + 0.17, 1.9, 0.48);
    s.addText(r[1], { x: 3.0, y: yy + 0.1, w: 3.3, h: rh - 0.2, fontSize: 13.5, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('ed', 3.0, yy + 0.1, 3.3, rh - 0.2, r[1], 13.5, { pad: 0.02 });
    s.addText(r[2], { x: 6.44, y: yy + 0.1, w: W - 0.62 - 6.44 - 0.26, h: rh - 0.2, fontSize: 12.5, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('ev', 6.44, yy + 0.1, W - 0.62 - 6.44 - 0.26, rh - 0.2, r[2], 12.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 0.1 + 4 * (rh + gap) + 0.04, W - 1.24, 0.8, P.soft2);
  s.addText('ที่น่ากลัวกว่า error คือสูตรที่ไม่ขึ้น error แต่ให้ตัวเลขผิด ซึ่งเป็นหัวข้อหลักของระดับ Advanced', {
    x: 0.9, y: y + 0.1 + 4 * (rh + gap) + 0.16, w: W - 1.8, h: 0.56, fontSize: 14, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('enote', 0.9, y + 0.1 + 4 * (rh + gap) + 0.16, W - 1.8, 0.56, 'ที่น่ากลัวกว่า error คือสูตรที่ไม่ขึ้น error แต่ให้ตัวเลขผิด ซึ่งเป็นหัวข้อหลักของระดับ Advanced', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('สร้าง error ทั้งสี่แบบให้ดูสด ๆ ใช้เวลาไม่เกินสามนาที ผู้เรียนจะเลิกกลัว error ทันทีเมื่อรู้ว่าแต่ละอันหมายถึงอะไร');
}

/* 19 WS1.3 */
{
  beginSlide('WS1.3');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 1.3 · COUNTIF และ SUMIF', 'สรุปข้อมูลรายกลุ่ม ใช้เวลา 20 นาที ทำคนเดียว');
  card(s, 0.62, y, W - 1.24, 2.2, P.soft);
  s.addText('สิ่งที่ต้องทำ', { x: 0.9, y: y + 0.2, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w1', 0.9, y + 0.2, W - 1.8, 0.36);
  bullets(s, 0.9, y + 0.62, W - 1.8, 1.45, [
    'ใช้ COUNTIF นับจำนวนรายการของแต่ละกลุ่มในไฟล์สาขาตนเอง',
    'ใช้ SUMIF รวมค่าของแต่ละกลุ่ม โดยใส่ $ ให้สูตรลากได้',
    'ตรวจว่าผลรวมของทุกกลุ่มเท่ากับผลรวมทั้งหมดหรือไม่',
  ], 13.5);
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.4, cw, 1.55, 'เกณฑ์ผ่าน',
    'ผลลัพธ์ตรงเฉลย และสูตรใช้การอ้างอิงแบบสัมบูรณ์ในตำแหน่งที่ถูกต้อง', P.soft, P.deep);
  textCard(s, 0.62 + cw + 0.4, y + 2.4, cw, 1.55, 'เคล็ดลับตรวจงานตัวเอง',
    'ถ้าผลรวมของทุกกลุ่มไม่เท่ากับผลรวมทั้งหมด แปลว่ามีกลุ่มตกหล่น มักเป็นชื่อที่สะกดต่างกัน', P.soft2, 'A06A0A');
  checkOverlaps();
  s.addNotes('เทคนิคตรวจผลรวมย้อนกลับในการ์ดขวาเป็นทักษะที่จะใช้ไปตลอด ให้ย้ำว่าเป็นวิธีตรวจงานตัวเองที่เร็วที่สุด');
}

/* 20 กราฟ */
{
  beginSlide('กราฟ');
  const s = lightSlide();
  const y = head(s, 6, 'เลือกกราฟจากคำถาม ไม่ใช่จากความสวย', 'ถามตัวเองก่อนว่าอยากให้คนดูเห็นอะไร');
  const rows = [
    ['อยากเปรียบเทียบระหว่างกลุ่ม', 'กราฟแท่ง', 'ผลผลิตของแต่ละตำบล', P.blue],
    ['อยากดูการเปลี่ยนแปลงตามเวลา', 'กราฟเส้น', 'ค่าฝุ่นรายเดือนตลอดปี', P.blue],
    ['อยากดูสัดส่วนของทั้งหมด', 'กราฟวงกลม ไม่เกิน 5 กลุ่ม', 'สัดส่วนชนิดพืชที่ปลูกในพื้นที่', P.amber],
    ['อยากดูความสัมพันธ์ของสองตัวแปร', 'กราฟกระจาย', 'ปริมาณปุ๋ยกับผลผลิต', P.blue],
  ];
  const rh = 0.8, gap = 0.15;
  rows.forEach((r, i) => {
    const yy = y + 0.08 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 2 ? P.soft2 : P.soft);
    s.addText(r[0], { x: 0.88, y: yy + 0.1, w: 4.2, h: rh - 0.2, fontSize: 13.5, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('gq', 0.88, yy + 0.1, 4.2, rh - 0.2, r[0], 13.5, { pad: 0.02 });
    s.addShape(pres.ShapeType.roundRect, { x: 5.2, y: yy + 0.16, w: 3.0, h: 0.48, rectRadius: 0.07, fill: { color: r[3] }, line: { width: 0 } });
    s.addText(r[1], { x: 5.2, y: yy + 0.16, w: 3.0, h: 0.48, fontSize: 13, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('gt', 5.2, yy + 0.16, 3.0, 0.48);
    s.addText(r[2], { x: 8.44, y: yy + 0.1, w: W - 0.62 - 8.44 - 0.26, h: rh - 0.2, fontSize: 12.5, color: P.muted, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('ge', 8.44, yy + 0.1, W - 0.62 - 8.44 - 0.26, rh - 0.2, r[2], 12.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 0.08 + 4 * (rh + gap) + 0.04, W - 1.24, 0.85, P.warn);
  s.addText('กราฟทุกชิ้นต้องมีชื่อแกนและหน่วยเสมอ กราฟที่ไม่มีหน่วยคือกราฟที่อ่านไม่ได้ ไม่ว่าจะสวยแค่ไหน', {
    x: 0.9, y: y + 0.08 + 4 * (rh + gap) + 0.16, w: W - 1.8, h: 0.6, fontSize: 14, bold: true, color: 'A83A36', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('gnote', 0.9, y + 0.08 + 4 * (rh + gap) + 0.16, W - 1.8, 0.6, 'กราฟทุกชิ้นต้องมีชื่อแกนและหน่วยเสมอ กราฟที่ไม่มีหน่วยคือกราฟที่อ่านไม่ได้ ไม่ว่าจะสวยแค่ไหน', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ยกตัวอย่างกราฟวงกลมที่เอาค่าฝุ่นสิบสองเดือนมาทำ แล้วถามว่ามันแปลว่าอะไร ผู้เรียนจะเห็นเองว่าไม่มีความหมาย');
}

/* 21 WS1.4 */
{
  beginSlide('WS1.4');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 1.4 · เล่าเรื่องด้วยกราฟ', 'ชิ้นงานประเมินหลักของระดับนี้ ใช้เวลา 20 นาที ทำเป็นคู่');
  card(s, 0.62, y, W - 1.24, 1.75, P.soft);
  s.addText('โจทย์', { x: 0.9, y: y + 0.18, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('f1', 0.9, y + 0.18, W - 1.8, 0.36);
  s.addText('จากไฟล์ที่ทำความสะอาดแล้ว สร้างกราฟ 2 ชิ้นที่ตอบคำถามของสาขาท่านตามที่ระบุในชีตอ่านก่อนเริ่ม พร้อมเขียนข้อสรุป 2 บรรทัดใต้กราฟ', {
    x: 0.9, y: y + 0.6, w: W - 1.8, h: 0.95, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('f2', 0.9, y + 0.6, W - 1.8, 0.95, 'จากไฟล์ที่ทำความสะอาดแล้ว สร้างกราฟ 2 ชิ้นที่ตอบคำถามของสาขาท่านตามที่ระบุในชีตอ่านก่อนเริ่ม พร้อมเขียนข้อสรุป 2 บรรทัดใต้กราฟ', 13.5, { pad: 0.02 });

  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 1.95, cw, 2.0, 'สิ่งที่ทำให้ได้คะแนนดี',
    'ชนิดกราฟตรงกับคำถาม มีชื่อแกนและหน่วยครบ และข้อสรุปอ้างตัวเลขจากกราฟจริง ไม่ใช่ความรู้สึก', P.soft, P.deep);
  textCard(s, 0.62 + cw + 0.4, y + 1.95, cw, 2.0, 'สิ่งที่ทำให้เสียคะแนน',
    'ข้อสรุปที่บรรยายซ้ำสิ่งที่กราฟแสดงอยู่แล้ว เช่น เขียนว่าตำบล ก มีค่าสูงที่สุด โดยไม่บอกว่าแปลว่าอะไร', P.warn, 'A83A36');
  checkOverlaps();
  s.addNotes('แจกเกณฑ์ให้ดูก่อนเริ่มทำ และย้ำว่าข้อสรุปสองบรรทัดมีน้ำหนักคะแนนเท่ากับกราฟทั้งสองชิ้นรวมกัน');
}

/* 22 สรุป */
{
  beginSlide('สรุป');
  const s = lightSlide();
  const y = head(s, null, 'สามอย่างที่อยากให้จำ');
  const items = [
    ['1', 'ดูการจัดชิดของเซลล์ก่อนเสมอ', 'ชิดซ้ายคือข้อความ ชิดขวาคือตัวเลข สัญญาณเล็ก ๆ นี้บอกได้ทันทีว่าทำไมสูตรถึงไม่ทำงาน', P.blue],
    ['2', 'สำรวจให้ครบก่อนลงมือแก้', 'การรีบแก้ทีละจุดโดยไม่เห็นภาพรวม มักทำให้แก้ผิดลำดับแล้วต้องย้อนกลับมาทำใหม่', P.amber],
    ['3', 'บันทึกเหตุผลของทุกการตัดสินใจ', 'สิ่งที่ทำให้งานวิเคราะห์เชื่อถือได้ ไม่ใช่ตัวเลขที่สวย แต่คือการที่คนอื่นตรวจสอบย้อนกลับได้', P.deep],
  ];
  const ch = 1.32, gap = 0.22;
  items.forEach((it, i) => {
    const yy = y + 0.15 + i * (ch + gap);
    card(s, 0.62, yy, W - 1.24, ch, P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: 0.9, y: yy + 0.36, w: 0.6, h: 0.6, fill: { color: it[3] }, line: { width: 0 } });
    s.addText(it[0], { x: 0.9, y: yy + 0.36, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sn2', 0.9, yy + 0.36, 0.6, 0.6);
    s.addText(it[1], { x: 1.72, y: yy + 0.18, w: W - 0.62 - 1.72 - 0.3, h: 0.44, fontSize: 16.5, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('sk2', 1.72, yy + 0.18, W - 0.62 - 1.72 - 0.3, 0.44, it[1], 16.5, { pad: 0.02 });
    s.addText(it[2], { x: 1.72, y: yy + 0.64, w: W - 0.62 - 1.72 - 0.3, h: 0.56, fontSize: 13, color: P.ink70, fontFace: FONT, valign: 'top', margin: 0 });
    reg('sv2', 1.72, yy + 0.64, W - 0.62 - 1.72 - 0.3, 0.56, it[2], 13, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนจดสามข้อนี้ลงในชีตบันทึกการตัดสินใจของตัวเอง แล้วจึงทำแบบทดสอบหลังเรียน');
}

/* 23 ปิด */
{
  beginSlide('ปิด');
  const s = darkSlide();
  s.addShape(pres.ShapeType.ellipse, { x: 10.8, y: 4.2, w: 3.8, h: 3.8, fill: { color: P.deep }, line: { width: 0 } });
  s.addText('แบบทดสอบหลังเรียน', { x: 0.9, y: 2.05, w: 9.5, h: 0.8, fontSize: 34, bold: true, color: P.white, fontFace: FONT, margin: 0 });
  reg('e1', 0.9, 2.05, 9.5, 0.8);
  s.addText('ใช้เวลา 5 นาที ทำโดยไม่เปิดเอกสาร เพื่อให้เห็นพัฒนาการของตัวเองได้จริง', {
    x: 0.9, y: 2.95, w: 9.5, h: 0.5, fontSize: 16, color: 'C3D9EF', fontFace: FONT, margin: 0
  });
  reg('e2', 0.9, 2.95, 9.5, 0.5, 'ใช้เวลา 5 นาที ทำโดยไม่เปิดเอกสาร เพื่อให้เห็นพัฒนาการของตัวเองได้จริง', 16, { pad: 0.05 });
  s.addText('ระดับถัดไปรออยู่', { x: 0.9, y: 3.85, w: 9.5, h: 0.5, fontSize: 19, bold: true, color: P.sky, fontFace: FONT, margin: 0 });
  reg('e3', 0.9, 3.85, 9.5, 0.5);
  s.addText('ระดับ Intermediate จะสอนการเชื่อมตารางและ PivotTable ซึ่งใช้ข้อมูลชุดเดิมที่คุณทำความสะอาดไว้แล้ววันนี้  ·  ส่งแบบฝึกหัดผ่าน Google Classroom ภายในหนึ่งสัปดาห์', {
    x: 0.9, y: 4.4, w: 9.8, h: 0.75, fontSize: 13, color: '8FB0D4', fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('e4', 0.9, 4.4, 9.8, 0.75, 'ระดับ Intermediate จะสอนการเชื่อมตารางและ PivotTable ซึ่งใช้ข้อมูลชุดเดิมที่คุณทำความสะอาดไว้แล้ววันนี้  ·  ส่งแบบฝึกหัดผ่าน Google Classroom ภายในหนึ่งสัปดาห์', 13, { pad: 0.05 });
  checkOverlaps();
  s.addNotes('ปิดด้วยการบอกว่าไฟล์ที่ทำความสะอาดวันนี้จะถูกใช้ต่อในระดับถัดไป งานวันนี้จึงไม่สูญเปล่า เป็นแรงจูงใจให้ทำให้ดี');
}

/* ============================================================ */
const out = process.argv[2];
pres.writeFile({ fileName: out }).then(() => {
  console.log('เขียนไฟล์แล้ว:', out);
  const n = report();
  if (process.argv[3] === '--measure') {
    const off = {};
    L.extents.forEach(e => {
      const prev = OFFSETS[String(e.no)] || 0;
      off[String(e.no)] = Math.round((prev + Math.max(0, Math.min(1.1, e.unused * 0.5))) * 100) / 100;
    });
    fs.writeFileSync(OFFSET_FILE, JSON.stringify(off, null, 2));
    console.log('บันทึกระยะเลื่อนลงไฟล์ deck_offsets_beg.json แล้ว');
  }
  if (n > 0) process.exitCode = 1;
});
