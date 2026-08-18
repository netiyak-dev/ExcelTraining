/**
 * สไลด์สอนระดับ Intermediate — จากตารางดิบสู่ข้อสรุป (3 ชั่วโมง)
 * ชุดการเรียนรู้ Excel & Google Sheets — อาจารย์ ดร.เนติยา การะเกตุ
 * ตัวอย่างทั้งหมดอ้างอิงไฟล์ฝึกปฏิบัติ 7 สาขา ตัวเลขดึงจาก answer_key.json
 */
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const L = require('./deck_lib.js');
const { W, H, FONT, startSlide, checkOverlaps, report } = L;

const key = JSON.parse(fs.readFileSync(__dirname + '/answer_key.json', 'utf8'));

/* จานสีประจำระดับ Intermediate — เขียวน้ำทะเล ต่างจาก Advanced ที่เป็นม่วง
   เพื่อให้ผู้เรียนแยกออกทันทีว่ากำลังดูสไลด์ระดับไหน */
const P = {
  ink:   '0B2E2A',
  deep:  '0E5C52',
  teal:  '2A9D8F',
  amber: 'E9A23B',
  coral: 'D2544F',
  white: 'FFFFFF',
  soft:  'EAF5F2',
  soft2: 'FDF3E3',
  warn:  'FCE9E8',
  ink70: '2E4A45',
  muted: '5C7873',
  line:  'CDE3DE'
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Dr. Netiya Karaket';
pres.title = 'Intermediate — จากตารางดิบสู่ข้อสรุป';

const shadow = () => ({ type: 'outer', color: '0B2E2A', blur: 10, offset: 2, angle: 90, opacity: 0.10 });

/* ---------- ระบบจัดสมดุลแนวตั้ง (สองรอบ เหมือนชุด Advanced) ---------- */
const OFFSET_FILE = __dirname + '/deck_offsets_int.json';
const OFFSETS = fs.existsSync(OFFSET_FILE) ? JSON.parse(fs.readFileSync(OFFSET_FILE, 'utf8')) : {};
const HEAD_BAND = 1.1;
let curOffset = 0, curNo = 0;

const regRaw = L.reg;
function reg(kind, x, y, w, h, text, size, opts) {
  return regRaw(kind, x, y >= HEAD_BAND ? y + curOffset : y, w, h, text, size, opts);
}
function shiftOpts(o) {
  if (o && typeof o.y === 'number' && o.y >= HEAD_BAND) {
    const c = Object.assign({}, o); c.y = o.y + curOffset; return c;
  }
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
  curNo = info.no;
  curOffset = OFFSETS[String(curNo)] || 0;
  return info;
}

/* ---------- ชิ้นส่วนที่ใช้ซ้ำ ---------- */
function darkSlide() { const s = pres.addSlide(); s.background = { color: P.ink }; return wrapSlide(s); }
function lightSlide() { const s = pres.addSlide(); s.background = { color: P.white }; return wrapSlide(s); }

function head(s, num, title, sub) {
  const x = 0.62, y = 0.42;
  if (num) {
    s.addShape(pres.ShapeType.ellipse, { x, y: y + 0.02, w: 0.52, h: 0.52, fill: { color: P.teal }, line: { width: 0 } });
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
  s.addShape(pres.ShapeType.roundRect, { x, y, w: tw, h: 0.46, rectRadius: 0.07, fill: { color: accent || P.teal }, line: { width: 0 } });
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
  s.addShape(pres.ShapeType.ellipse, { x: 11.8, y: 4.5, w: 2.8, h: 2.8, fill: { color: P.teal }, line: { width: 0 } });
  s.addText('ระดับที่ 2', { x: 0.9, y: 1.75, w: 6, h: 0.4, fontSize: 15, color: P.teal, fontFace: FONT, margin: 0, charSpacing: 2 });
  reg('eyebrow', 0.9, 1.75, 6, 0.4);
  s.addText('Intermediate', { x: 0.9, y: 2.2, w: 8.8, h: 1.0, fontSize: 44, bold: true, color: P.white, fontFace: FONT, margin: 0 });
  reg('h1', 0.9, 2.2, 8.8, 1.0);
  s.addText('จากตารางดิบสู่ข้อสรุป', { x: 0.9, y: 3.22, w: 8.8, h: 0.6, fontSize: 24, color: 'B9DBD5', fontFace: FONT, margin: 0 });
  reg('h2', 0.9, 3.22, 8.8, 0.6, 'จากตารางดิบสู่ข้อสรุป', 24, { pad: 0.05 });
  s.addText('3 ชั่วโมง รวมทำแบบฝึกหัด · ใช้ชุดข้อมูลจริงของสาขาตนเอง', { x: 0.9, y: 3.95, w: 8.8, h: 0.4, fontSize: 15, color: '7FAEA6', fontFace: FONT, margin: 0 });
  reg('h3', 0.9, 3.95, 8.8, 0.4);
  s.addText('อาจารย์ ดร.เนติยา การะเกตุ  ·  หลักสูตรวิทยาศาสตร์การเกษตร  ·  ม.มหิดล วิทยาเขตกาญจนบุรี',
    { x: 0.9, y: 5.9, w: 10, h: 0.4, fontSize: 13, color: '5D8880', fontFace: FONT, margin: 0 });
  reg('by', 0.9, 5.9, 10, 0.4);
  checkOverlaps();
  s.addNotes('เปิดคาบด้วยการถามว่า ใครเคยทำรายงานเสร็จแล้วพบทีหลังว่าตัวเลขผิดบ้าง แล้วโยงว่าระดับนี้คือการเปลี่ยนบทบาทจากผู้บันทึกข้อมูลเป็นผู้วิเคราะห์');
}

/* 2 แนวคิดหลัก */
{
  beginSlide('แนวคิดหลัก');
  const s = lightSlide();
  const y = head(s, null, 'แนวคิดหลักของระดับนี้');
  card(s, 0.62, y + 0.15, W - 1.24, 1.9, P.soft);
  s.addText('ข้อมูลไม่ได้พูดเอง เราต้องตั้งคำถามให้ถูก', { x: 1.0, y: y + 0.45, w: W - 2.0, h: 0.66, fontSize: 30, bold: true, color: P.deep, fontFace: FONT, margin: 0, align: 'center' });
  reg('big', 1.0, y + 0.45, W - 2.0, 0.66);
  s.addText('เครื่องมือช่วยสรุปได้ แต่ไม่ได้บอกว่าอะไรสำคัญ', { x: 1.0, y: y + 1.18, w: W - 2.0, h: 0.5, fontSize: 17, color: P.muted, fontFace: FONT, margin: 0, align: 'center' });
  reg('bigsub', 1.0, y + 1.18, W - 2.0, 0.5);

  const cy = y + 2.3, cw = (W - 1.24 - 0.6) / 3, ch = 2.1;
  textCard(s, 0.62, cy, cw, ch, 'ระดับที่แล้วทำอะไร', 'ทำให้ข้อมูลสะอาดและคำนวณค่าพื้นฐานได้ถูกต้อง เป็นการเตรียมวัตถุดิบ', P.soft);
  textCard(s, 0.62 + cw + 0.3, cy, cw, ch, 'ระดับนี้ทำอะไร', 'เชื่อมข้อมูลหลายตารางเข้าด้วยกัน สรุปหลายมิติ แล้วเปลี่ยนตัวเลขเป็นข้อสรุปที่ใช้ตัดสินใจได้', P.soft2, 'A06A0A');
  textCard(s, 0.62 + (cw + 0.3) * 2, cy, cw, ch, 'สิ่งที่ยังต้องทำเอง', 'การตั้งคำถามว่าอะไรควรรู้ และการตีความว่าตัวเลขที่ได้แปลว่าอะไร ยังเป็นหน้าที่ของคน', P.soft);
  checkOverlaps();
  s.addNotes('ย้ำว่าเครื่องมือในคาบนี้ทรงพลังมาก จึงยิ่งอันตรายถ้าใช้กับข้อมูลที่ยังไม่สะอาด เพราะจะได้ข้อสรุปที่ผิดแบบดูน่าเชื่อถือ');
}

/* 3 CLO */
{
  beginSlide('CLO');
  const s = lightSlide();
  const y = head(s, null, 'เมื่อจบคาบนี้ คุณจะทำอะไรได้', 'ผลลัพธ์การเรียนรู้ทั้งหกข้อ วัดผลด้วยแบบฝึกหัดและแบบทดสอบหลังเรียน');
  const items = [
    ['2.1', 'แปลงช่วงข้อมูลเป็นตารางแบบมีโครงสร้าง และใช้ชื่อช่วง เพื่อให้สูตรขยายตัวเองได้'],
    ['2.2', 'ดึงข้อมูลข้ามตารางด้วย VLOOKUP, XLOOKUP หรือ INDEX+MATCH และจัดการข้อผิดพลาดด้วย IFERROR'],
    ['2.3', 'สร้างและปรับแต่ง PivotTable สรุปข้อมูลหลายมิติ พร้อมใช้ Slicer และ Pivot Chart'],
    ['2.4', 'ใช้ฟังก์ชันเงื่อนไขหลายชั้น แก้โจทย์ที่มีเงื่อนไขมากกว่าหนึ่งข้อ'],
    ['2.5', 'ตั้งกฎ Data Validation และป้องกันชีต เพื่อลดความผิดพลาดของผู้กรอกข้อมูลคนอื่น'],
    ['2.6', 'ตีความผลสรุป และเขียนข้อสรุปเชิงวิเคราะห์ที่มีหลักฐานสนับสนุน'],
  ];
  const rh = 0.74, gap = 0.15;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2), cw = (W - 1.24 - 0.4) / 2;
    const x = 0.62 + col * (cw + 0.4), yy = y + 0.1 + row * (rh + gap);
    card(s, x, yy, cw, rh, i === 5 ? P.soft2 : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.16, y: yy + 0.17, w: 0.62, h: 0.4, rectRadius: 0.06, fill: { color: i === 5 ? P.amber : P.teal }, line: { width: 0 } });
    s.addText(it[0], { x: x + 0.16, y: yy + 0.17, w: 0.62, h: 0.4, fontSize: 13, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('cloNum', x + 0.16, yy + 0.17, 0.62, 0.4);
    s.addText(it[1], { x: x + 0.92, y: yy + 0.08, w: cw - 1.1, h: rh - 0.16, fontSize: 13, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('cloText', x + 0.92, yy + 0.08, cw - 1.1, rh - 0.16, it[1], 13, { pad: 0.02 });
  });
  s.addText('ข้อ 2.6 เป็นข้อเดียวที่อยู่ระดับ Analyze และเป็นสิ่งที่ทำให้งานของคุณต่างจากการกดปุ่มให้โปรแกรมสรุป', {
    x: 0.62, y: y + 0.1 + 3 * (rh + gap) + 0.12, w: W - 1.24, h: 0.42, fontSize: 13, italic: true, color: 'A06A0A', fontFace: FONT, margin: 0
  });
  reg('note', 0.62, y + 0.1 + 3 * (rh + gap) + 0.12, W - 1.24, 0.42);
  checkOverlaps();
  s.addNotes('บอกว่าข้อ 2.1 ถึง 2.5 คือเครื่องมือ ส่วน 2.6 คือเหตุผลที่เราเรียนเครื่องมือเหล่านั้น');
}

/* 4 ผังเวลา */
{
  beginSlide('ผังเวลา');
  const s = lightSlide();
  const y = head(s, null, 'ผังการเรียน 180 นาที');
  const rows = [
    ['0–10', 'แบบทดสอบก่อนเรียน', 'ทำโดยไม่เปิดเอกสาร', P.muted],
    ['10–35', 'โครงสร้างข้อมูลระดับมืออาชีพ', 'ตารางแบบมีโครงสร้างและการแยกข้อมูลดิบออกจากรายงาน', P.teal],
    ['35–70', 'ตระกูลฟังก์ชันค้นหา', 'VLOOKUP, XLOOKUP, INDEX+MATCH และ IFERROR', P.teal],
    ['70–80', 'พักเบรก', '', P.muted],
    ['80–120', 'PivotTable', 'หัวใจของการสรุปข้อมูลหลายมิติ', P.amber],
    ['120–150', 'ฟังก์ชันเงื่อนไขหลายชั้น', 'SUMIFS, COUNTIFS, AVERAGEIFS และฟังก์ชันข้อความกับวันที่', P.teal],
    ['150–170', 'ป้องกันความผิดพลาดและแบบฝึกหัดรวบยอด', 'Data Validation และการเขียนข้อสรุป', P.teal],
    ['170–180', 'แบบทดสอบหลังเรียนและสรุป', '', P.muted],
  ];
  rows.forEach((r, i) => timeRow(s, 0.62, y + 0.1 + i * 0.6, W - 1.24, r[0], r[1], r[2], r[3]));
  checkOverlaps();
  s.addNotes('ชี้ว่าช่วง 80 ถึง 120 นาทีคือช่วงที่มีน้ำหนักที่สุด และเป็นทักษะที่ผู้เรียนจะได้ใช้บ่อยที่สุดหลังจบคาบ');
}

/* 5 ตารางแบบมีโครงสร้าง */
{
  beginSlide('Ctrl+T');
  const s = lightSlide();
  const y = head(s, 1, 'เปลี่ยนช่วงข้อมูลให้เป็นตารางก่อนเสมอ', 'กด Ctrl+T หนึ่งครั้ง แล้วปัญหาที่ตามมาทีหลังหายไปครึ่งหนึ่ง');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.15, P.warn);
  s.addText('ถ้าไม่แปลงเป็นตาราง', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('t1', 0.86, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86, y + 0.68, cw - 0.48, 2.3, [
    'เพิ่มข้อมูลใหม่แล้วสูตรไม่ครอบคลุม ต้องแก้ช่วงเอง',
    'PivotTable ไม่เห็นข้อมูลที่เพิ่มเข้ามา',
    'สูตรอ่านยาก เต็มไปด้วย A2:A5000',
    'ลืมขยายช่วงเพียงครั้งเดียว รายงานก็ผิดทั้งฉบับ',
  ], 13.5);

  card(s, 0.62 + cw + 0.4, y, cw, 3.15, P.soft);
  s.addText('เมื่อแปลงเป็นตารางแล้ว', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('t2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86 + cw + 0.4, y + 0.68, cw - 0.48, 2.3, [
    'ช่วงขยายตัวเองเมื่อมีข้อมูลใหม่',
    'PivotTable กด Refresh แล้วครอบคลุมทันที',
    'สูตรอ่านเป็นภาษาคน เช่น ผลผลิต แทน G2:G500',
    'ตัวกรองและแถบหัวตารางติดมาให้เลย',
  ], 13.5);

  card(s, 0.62, y + 3.35, W - 1.24, 0.8, P.soft2);
  s.addText('ถ้าจำได้อย่างเดียวจากโมดูลนี้ ให้จำว่า เปิดไฟล์ปุ๊บ กด Ctrl+T ก่อนทำอย่างอื่น', {
    x: 0.9, y: y + 3.48, w: W - 1.8, h: 0.54, fontSize: 15.5, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('t3', 0.9, y + 3.48, W - 1.8, 0.54, 'ถ้าจำได้อย่างเดียวจากโมดูลนี้ ให้จำว่า เปิดไฟล์ปุ๊บ กด Ctrl+T ก่อนทำอย่างอื่น', 15.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('สาธิตสดโดยเพิ่มแถวใหม่ท้ายตาราง ให้เห็นว่าสูตรและ PivotTable ตามมาเองจริง ๆ ใช้เวลาไม่เกินสองนาที');
}

/* ภาพเปรียบเทียบ ช่วงธรรมดากับตารางแบบมีโครงสร้าง */
{
  const s = mockPair('ภาพช่วงกับตาราง', 1,
    'เกิดอะไรขึ้นเมื่อมีแถวใหม่เข้ามา',
    'ข้อมูลเหมือนกันทุกอย่าง ต่างกันแค่ฝั่งขวาแปลงเป็นตารางแล้ว',
    'ช่วงธรรมดา ผลรวมไม่ขยับ', 'ตารางแบบมีโครงสร้าง ขยายเอง',
    { cols: ['แปลง', 'ผลผลิต', 'สูตรที่ใช้'],
      rows: [['AG-01', ['1250', 'n'], ''], ['AG-02', ['1180', 'n'], ''],
             [['AG-03', 'x'], ['1340', 'xn'], ['แถวใหม่', 'x']],
             ['รวม', ['2430', 'xn'], 'SUM(B2:B3)']] },
    { cols: ['แปลง', 'ผลผลิต', 'สูตรที่ใช้'],
      rows: [['AG-01', ['1250', 'n'], ''], ['AG-02', ['1180', 'n'], ''],
             ['AG-03', ['1340', 'n'], 'แถวใหม่'],
             ['รวม', ['3770', 'n'], 'SUM(ผลผลิต[ผลผลิต])']] },
    'ช่วง B2:B3 ถูกตรึงไว้ตอนเขียนสูตร แถวที่เพิ่มเข้ามาจึงตกหล่นเงียบ ๆ นี่คือสาเหตุอันดับหนึ่งของรายงานที่ผิดโดยไม่มีใครรู้ตัว');
  s.addNotes('ถามว่าเคยเจอไหมที่ยอดในรายงานไม่ตรงกับที่คำนวณเอง ส่วนใหญ่จะเคย และมักมาจากสาเหตุนี้');
}

/* 6 แยกข้อมูลดิบออกจากรายงาน */
{
  beginSlide('แยกดิบ/รายงาน');
  const s = lightSlide();
  const y = head(s, 1, 'แยกชีตข้อมูลดิบออกจากชีตรายงาน', 'กฎข้อเดียวที่ทำให้ไฟล์อยู่รอดได้เป็นปี');
  const bw = 3.5, gapx = (W - 1.24 - bw * 3) / 2;
  const steps = [
    ['ชีตข้อมูลดิบ', 'วางข้อมูลที่ได้มาตามเดิม ห้ามแก้ ห้ามใส่สูตร ห้ามจัดสีเล่น', P.teal],
    ['ชีตคำนวณ', 'ทำความสะอาดและเพิ่มคอลัมน์ช่วยที่นี่ ทุกอย่างเป็นสูตร', P.deep],
    ['ชีตรายงาน', 'PivotTable กราฟ และข้อสรุปสำหรับคนอ่าน', P.amber],
  ];
  steps.forEach((st, i) => {
    const x = 0.62 + i * (bw + gapx);
    card(s, x, y + 0.3, bw, 2.1, P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: x + (bw - 0.58) / 2, y: y + 0.5, w: 0.58, h: 0.58, fill: { color: st[2] }, line: { width: 0 } });
    s.addText(String(i + 1), { x: x + (bw - 0.58) / 2, y: y + 0.5, w: 0.58, h: 0.58, fontSize: 19, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sn', x + (bw - 0.58) / 2, y + 0.5, 0.58, 0.58);
    s.addText(st[0], { x: x + 0.16, y: y + 1.2, w: bw - 0.32, h: 0.4, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, align: 'center', margin: 0 });
    reg('sname', x + 0.16, y + 1.2, bw - 0.32, 0.4, st[0], 16, { pad: 0.02 });
    s.addText(st[1], { x: x + 0.18, y: y + 1.62, w: bw - 0.36, h: 0.68, fontSize: 13, color: P.ink70, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('sdesc', x + 0.18, y + 1.62, bw - 0.36, 0.68, st[1], 13, { pad: 0.02 });
    if (i < 2) s.addText('→', { x: x + bw + 0.02, y: y + 1.1, w: gapx - 0.04, h: 0.5, fontSize: 22, color: P.line, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
  });
  card(s, 0.62, y + 2.65, W - 1.24, 1.3, P.warn);
  s.addText('เหตุผลที่ต้องแยก', { x: 0.9, y: y + 2.8, w: W - 1.8, h: 0.36, fontSize: 15, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('rw1', 0.9, y + 2.8, W - 1.8, 0.36);
  s.addText('เมื่อข้อมูลชุดใหม่มาถึง คุณแค่วางทับชีตข้อมูลดิบแล้วกด Refresh งานเสร็จ แต่ถ้าคำนวณปนอยู่ในชีตเดียวกับข้อมูล คุณจะต้องทำใหม่ทั้งหมดทุกครั้ง และจะไม่มีทางรู้ว่าตัวเลขไหนเป็นของเดิม', {
    x: 0.9, y: y + 3.18, w: W - 1.8, h: 0.66, fontSize: 13, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('rw2', 0.9, y + 3.18, W - 1.8, 0.66, 'เมื่อข้อมูลชุดใหม่มาถึง คุณแค่วางทับชีตข้อมูลดิบแล้วกด Refresh งานเสร็จ แต่ถ้าคำนวณปนอยู่ในชีตเดียวกับข้อมูล คุณจะต้องทำใหม่ทั้งหมดทุกครั้ง และจะไม่มีทางรู้ว่าตัวเลขไหนเป็นของเดิม', 13, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('เชื่อมโยงว่าข้อนี้คือข้อชี้ขาดของแบบทดสอบจัดระดับ Block B ผู้ที่ตอบข้อนี้ได้คือผู้ที่พร้อมขึ้นระดับ Advanced');
}

/* 7 ล้างก่อน Pivot */
{
  beginSlide('ล้างก่อน Pivot');
  const s = lightSlide();
  const y = head(s, 1, 'ทำไมต้องล้างข้อมูลก่อนสรุป', 'ตัวอย่างจริงจากไฟล์ฝึกปฏิบัติสาขาบริหารธุรกิจ');
  card(s, 0.62, y, W - 1.24, 2.5, P.soft);
  s.addText('ไฟล์ยอดขายร้านกาแฟมีข้อมูล 4 สาขา แต่เมื่อสร้าง PivotTable กลับได้', {
    x: 0.9, y: y + 0.22, w: W - 1.8, h: 0.4, fontSize: 15, bold: true, color: P.deep, fontFace: FONT, margin: 0
  });
  reg('e1', 0.9, y + 0.22, W - 1.8, 0.4);
  const sw = (W - 1.8 - 1.0) / 3;
  stat(s, 0.9, y + 0.72, sw, '7', 'กลุ่มสาขาใน PivotTable ถ้าใช้ข้อมูลดิบ', P.coral);
  stat(s, 0.9 + sw + 0.5, y + 0.72, sw, '4', 'สาขาจริงหลังทำชื่อให้เป็นมาตรฐานเดียวกัน', P.teal);
  stat(s, 0.9 + (sw + 0.5) * 2, y + 0.72, sw, '10', 'แถวซ้ำทั้งแถวที่ต้องลบก่อนนับ', P.amber);

  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.7, cw, 1.62, 'สาเหตุ',
    'สาขาศาลายา ถูกกรอกเป็น ศาลายา บ้าง มีช่องว่างท้ายบ้าง PivotTable จึงมองว่าเป็นคนละสาขา', P.warn, 'A83A36');
  textCard(s, 0.62 + cw + 0.4, y + 2.7, cw, 1.62, 'บทเรียน',
    'PivotTable เชื่อข้อมูลที่ป้อนเข้าไปเสมอ มันไม่รู้ว่าสองชื่อนี้คือสาขาเดียวกัน การล้างข้อมูลจึงต้องมาก่อน', P.soft, P.deep);
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนเปิดไฟล์ BA แล้วสร้าง PivotTable นับจำนวนใบเสร็จรายสาขาทันที จะเห็นเจ็ดกลุ่มด้วยตาตัวเอง ทรงพลังกว่าการบอกเฉย ๆ มาก');
}

/* 8 ตระกูลฟังก์ชันค้นหา */
{
  beginSlide('lookup family');
  const s = lightSlide();
  const y = head(s, 2, 'ตระกูลฟังก์ชันค้นหา เลือกอันไหนดี', 'ทั้งสามทำงานเดียวกัน แต่มีข้อจำกัดต่างกัน');
  const rows = [
    ['VLOOKUP', 'มีในทุกเวอร์ชัน เขียนง่าย', 'ค้นย้อนซ้ายไม่ได้ และพังเมื่อมีการแทรกคอลัมน์', P.amber],
    ['XLOOKUP', 'ค้นได้ทุกทิศ ระบุค่าเมื่อไม่พบได้ในตัว', 'ไม่มีใน Excel 2016 และ 2019', P.teal],
    ['INDEX + MATCH', 'ทำงานได้ทุกเวอร์ชัน ทนต่อการแทรกคอลัมน์', 'เขียนยาวกว่า ต้องเข้าใจสองฟังก์ชัน', P.deep],
  ];
  const rh = 1.15, gap = 0.18;
  rows.forEach((r, i) => {
    const yy = y + 0.1 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.34, w: 2.5, h: 0.48, rectRadius: 0.07, fill: { color: r[3] }, line: { width: 0 } });
    s.addText(r[0], { x: 0.86, y: yy + 0.34, w: 2.5, h: 0.48, fontSize: 14, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('lk', 0.86, yy + 0.34, 2.5, 0.48);
    s.addText('ข้อดี', { x: 3.6, y: yy + 0.16, w: 4.0, h: 0.3, fontSize: 12, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
    reg('lgt', 3.6, yy + 0.16, 4.0, 0.3);
    s.addText(r[1], { x: 3.6, y: yy + 0.48, w: 4.0, h: 0.56, fontSize: 13, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top' });
    reg('lg', 3.6, yy + 0.48, 4.0, 0.56, r[1], 13, { pad: 0.02 });
    s.addText('ข้อจำกัด', { x: 7.9, y: yy + 0.16, w: W - 0.62 - 7.9 - 0.26, h: 0.3, fontSize: 12, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
    reg('lbt', 7.9, yy + 0.16, W - 0.62 - 7.9 - 0.26, 0.3);
    s.addText(r[2], { x: 7.9, y: yy + 0.48, w: W - 0.62 - 7.9 - 0.26, h: 0.56, fontSize: 13, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top' });
    reg('lb', 7.9, yy + 0.48, W - 0.62 - 7.9 - 0.26, 0.56, r[2], 13, { pad: 0.02 });
  });
  s.addText('คำแนะนำ: ถ้าไฟล์จะส่งให้คนอื่นใช้ด้วย ให้เลือก INDEX+MATCH เพราะไม่มีทางพังเพราะเวอร์ชัน', {
    x: 0.62, y: y + 0.1 + 3 * (rh + gap) + 0.08, w: W - 1.24, h: 0.44, fontSize: 13.5, italic: true, color: P.deep, fontFace: FONT, margin: 0
  });
  reg('lknote', 0.62, y + 0.1 + 3 * (rh + gap) + 0.08, W - 1.24, 0.44);
  checkOverlaps();
  s.addNotes('ถามผู้เรียนว่าใครใช้ Excel 2019 หรือเก่ากว่าบ้าง จะได้รู้ว่าต้องเน้น INDEX+MATCH มากแค่ไหนในคาบนี้');
}

/* 9 กับดัก VLOOKUP */
{
  beginSlide('กับดัก VLOOKUP');
  const s = lightSlide();
  const y = head(s, 2, 'สองกับดักของ VLOOKUP ที่ผิดแบบเงียบ ๆ');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.4, P.warn);
  s.addText('กับดักที่ 1 · ค้นย้อนซ้ายไม่ได้', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 16, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('v1', 0.86, y + 0.2, cw - 0.48, 0.4);
  s.addText('VLOOKUP ค้นจากคอลัมน์ซ้ายสุดของช่วงเสมอ และคืนค่าจากคอลัมน์ทางขวาเท่านั้น ถ้าสิ่งที่อยากได้อยู่ทางซ้ายของคอลัมน์ค้นหา จะทำไม่ได้เลย', {
    x: 0.86, y: y + 0.66, w: cw - 0.48, h: 1.0, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('v1b', 0.86, y + 0.66, cw - 0.48, 1.0, 'VLOOKUP ค้นจากคอลัมน์ซ้ายสุดของช่วงเสมอ และคืนค่าจากคอลัมน์ทางขวาเท่านั้น ถ้าสิ่งที่อยากได้อยู่ทางซ้ายของคอลัมน์ค้นหา จะทำไม่ได้เลย', 13.5, { pad: 0.02 });
  card(s, 0.86, y + 1.78, cw - 0.48, 1.42, P.white);
  s.addText('ทางแก้', { x: 1.06, y: y + 1.9, w: cw - 0.88, h: 0.3, fontSize: 12.5, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
  reg('v1c', 1.06, y + 1.9, cw - 0.88, 0.3);
  s.addText('ใช้ INDEX+MATCH ซึ่งแยกการค้นหาออกจากการคืนค่า จึงไม่สนใจว่าคอลัมน์ไหนอยู่ซ้ายหรือขวา', {
    x: 1.06, y: y + 2.22, w: cw - 0.88, h: 0.85, fontSize: 13, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('v1d', 1.06, y + 2.22, cw - 0.88, 0.85, 'ใช้ INDEX+MATCH ซึ่งแยกการค้นหาออกจากการคืนค่า จึงไม่สนใจว่าคอลัมน์ไหนอยู่ซ้ายหรือขวา', 13, { pad: 0.02 });

  card(s, 0.62 + cw + 0.4, y, cw, 3.4, P.warn);
  s.addText('กับดักที่ 2 · ลืมใส่ FALSE', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 16, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('v2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  s.addText('อาร์กิวเมนต์ตัวสุดท้ายมีค่าเริ่มต้นเป็น TRUE ซึ่งแปลว่าค้นแบบใกล้เคียง ถ้าข้อมูลไม่ได้เรียงลำดับไว้ก่อน จะคืนค่าของแถวที่ผิดโดยไม่ขึ้น error ใด ๆ', {
    x: 0.86 + cw + 0.4, y: y + 0.66, w: cw - 0.48, h: 1.0, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('v2b', 0.86 + cw + 0.4, y + 0.66, cw - 0.48, 1.0, 'อาร์กิวเมนต์ตัวสุดท้ายมีค่าเริ่มต้นเป็น TRUE ซึ่งแปลว่าค้นแบบใกล้เคียง ถ้าข้อมูลไม่ได้เรียงลำดับไว้ก่อน จะคืนค่าของแถวที่ผิดโดยไม่ขึ้น error ใด ๆ', 13.5, { pad: 0.02 });
  card(s, 0.86 + cw + 0.4, y + 1.78, cw - 0.48, 1.42, P.white);
  s.addText('ทางแก้', { x: 1.06 + cw + 0.4, y: y + 1.9, w: cw - 0.88, h: 0.3, fontSize: 12.5, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
  reg('v2c', 1.06 + cw + 0.4, y + 1.9, cw - 0.88, 0.3);
  s.addText('ใส่ FALSE ทุกครั้งจนติดเป็นนิสัย นี่คือข้อผิดพลาดที่ตรวจจับยากที่สุดข้อหนึ่งในงานจริง', {
    x: 1.06 + cw + 0.4, y: y + 2.22, w: cw - 0.88, h: 0.85, fontSize: 13, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('v2d', 1.06 + cw + 0.4, y + 2.22, cw - 0.88, 0.85, 'ใส่ FALSE ทุกครั้งจนติดเป็นนิสัย นี่คือข้อผิดพลาดที่ตรวจจับยากที่สุดข้อหนึ่งในงานจริง', 13, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('กับดักที่สองสำคัญกว่าที่คิด เพราะผลลัพธ์ดูปกติทุกประการ ให้สาธิตด้วยข้อมูลที่ไม่ได้เรียงลำดับ แล้วชี้ว่าค่าที่ได้เป็นของแถวอื่น');
}

/* 10 IFERROR */
{
  beginSlide('IFERROR');
  const s = lightSlide();
  const y = head(s, 2, 'จัดการค่าที่หาไม่พบด้วย IFERROR', 'แต่ต้องใช้ให้ถูกวิธี ไม่ใช่ใช้กลบปัญหา');
  card(s, 0.62, y, W - 1.24, 1.65, P.soft);
  s.addText('ตัวอย่างจริงจากไฟล์สาขาบริหารธุรกิจ', { x: 0.9, y: y + 0.18, w: W - 1.8, h: 0.36, fontSize: 15, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('if1', 0.9, y + 0.18, W - 1.8, 0.36);
  s.addText('ไฟล์มีรหัสสินค้า BEV-099 ปนอยู่ 9 แถว ซึ่งไม่มีในตารางสินค้า เมื่อใช้ VLOOKUP ดึงราคาต้นทุนจึงได้ #N/A ทำให้ผลรวมกำไรทั้งคอลัมน์กลายเป็น #N/A ตามไปด้วย', {
    x: 0.9, y: y + 0.6, w: W - 1.8, h: 0.85, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('if2', 0.9, y + 0.6, W - 1.8, 0.85, 'ไฟล์มีรหัสสินค้า BEV-099 ปนอยู่ 9 แถว ซึ่งไม่มีในตารางสินค้า เมื่อใช้ VLOOKUP ดึงราคาต้นทุนจึงได้ #N/A ทำให้ผลรวมกำไรทั้งคอลัมน์กลายเป็น #N/A ตามไปด้วย', 13.5, { pad: 0.02 });

  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 1.85, cw, 2.05, 'วิธีที่ผิด',
    'ครอบ IFERROR แล้วให้คืนค่า 0 หรือช่องว่าง จบเรื่อง วิธีนี้ทำให้ตัวเลขดูสะอาดแต่ซ่อนปัญหาไว้ และคุณจะไม่มีวันรู้ว่ามีสินค้าที่ไม่มีในระบบอยู่กี่รายการ',
    P.warn, 'A83A36');
  textCard(s, 0.62 + cw + 0.4, y + 1.85, cw, 2.05, 'วิธีที่ถูก',
    'ให้ IFERROR คืนข้อความที่บอกสาเหตุ เช่น ไม่พบรหัสสินค้า แล้วนับจำนวนแถวที่เป็นแบบนั้น รายงานควบคู่ไปกับผลวิเคราะห์เสมอ',
    P.soft, P.deep);
  checkOverlaps();
  s.addNotes('ย้ำหลักการว่า IFERROR มีไว้ทำให้ข้อผิดพลาดอ่านรู้เรื่อง ไม่ใช่ทำให้ข้อผิดพลาดหายไปจากสายตา');
}

/* 11 WS2.2 */
{
  beginSlide('WS2.2');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 2.2 · เชื่อมตารางและคำนวณกำไร', 'ใช้เวลา 20 นาที ทำกับไฟล์ของสาขาตนเอง');
  card(s, 0.62, y, W - 1.24, 2.35, P.soft);
  s.addText('สิ่งที่ต้องทำ', { x: 0.9, y: y + 0.2, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w1', 0.9, y + 0.2, W - 1.8, 0.36);
  bullets(s, 0.9, y + 0.62, W - 1.8, 1.6, [
    'แปลงข้อมูลดิบและตารางอ้างอิงให้เป็นตารางด้วย Ctrl+T',
    'ใช้ฟังก์ชันค้นหาดึงข้อมูลจากตารางอ้างอิงมาต่อกับข้อมูลหลัก',
    'ครอบด้วย IFERROR ที่คืนข้อความบอกสาเหตุ แล้วนับว่ามีกี่แถวที่หาไม่พบ',
    'สรุปด้วย SUMIFS ตามเงื่อนไขที่โจทย์ของสาขาท่านกำหนด',
  ], 13.5);
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.55, cw, 1.35, 'เกณฑ์ผ่าน',
    'ไม่มี #N/A ค้างในไฟล์ และรายงานจำนวนแถวที่หาไม่พบไว้ชัดเจน', P.soft2, 'A06A0A');
  textCard(s, 0.62 + cw + 0.4, y + 2.55, cw, 1.35, 'สิ่งที่ผู้สอนจะดู',
    'ผู้ที่กลบ #N/A ด้วยศูนย์เฉย ๆ จะได้ผลลัพธ์สวยแต่ไม่ผ่านเกณฑ์', P.warn, 'A83A36');
  checkOverlaps();
  s.addNotes('เดินดูว่าใครใช้ IFERROR คืนค่า 0 ให้ถามกลับว่า แล้วจะรู้ได้อย่างไรว่ามีปัญหา แทนที่จะบอกว่าผิด');
}

/* 12 เบรก */
{
  beginSlide('เบรก');
  const s = darkSlide();
  s.addShape(pres.ShapeType.ellipse, { x: -1.2, y: 4.4, w: 4.6, h: 4.6, fill: { color: P.deep }, line: { width: 0 } });
  s.addText('พักเบรก 10 นาที', { x: 0.9, y: 2.6, w: 11.5, h: 0.9, fontSize: 38, bold: true, color: P.white, fontFace: FONT, margin: 0, align: 'center' });
  reg('brk', 0.9, 2.6, 11.5, 0.9);
  s.addText('กลับมาแล้วเราจะเข้าเครื่องมือที่ทรงพลังที่สุดของคาบนี้ นั่นคือ PivotTable', {
    x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 17, color: 'B9DBD5', fontFace: FONT, margin: 0, align: 'center'
  });
  reg('brk2', 0.9, 3.6, 11.5, 0.5, 'กลับมาแล้วเราจะเข้าเครื่องมือที่ทรงพลังที่สุดของคาบนี้ นั่นคือ PivotTable', 17, { pad: 0.05 });
  checkOverlaps();
  s.addNotes('ใช้ช่วงเบรกช่วยคนที่ยังติดแบบฝึกหัด 2.2 เพราะถ้าเชื่อมตารางไม่ได้ จะทำ PivotTable ต่อไม่ได้เลย');
}

/* 13 PivotTable 4 พื้นที่ */
{
  beginSlide('Pivot 4 พื้นที่');
  const s = lightSlide();
  const y = head(s, 3, 'PivotTable คือการถามคำถามกับข้อมูล', 'สี่พื้นที่ที่ต้องเข้าใจ แล้วที่เหลือคือการลองลากดู');
  const parts = [
    ['Rows', 'อยากแยกดูตามอะไร', 'เช่น รายจังหวัด รายสาขา รายชนิด', P.teal],
    ['Columns', 'อยากเปรียบเทียบตามอะไรอีกมิติ', 'เช่น รายเดือน รายปี รายกะ', P.deep],
    ['Values', 'อยากรู้ตัวเลขอะไร และสรุปแบบไหน', 'ผลรวม ค่าเฉลี่ย หรือจำนวนนับ', P.amber],
    ['Filters', 'อยากดูเฉพาะส่วนไหน', 'เช่น เฉพาะปีนี้ เฉพาะสาขาที่สนใจ', P.muted],
  ];
  const cw = (W - 1.24 - 0.45) / 4, ch = 2.55;
  parts.forEach((p, i) => {
    const x = 0.62 + i * (cw + 0.15);
    card(s, x, y, cw, ch, P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.16, y: y + 0.22, w: cw - 0.32, h: 0.5, rectRadius: 0.07, fill: { color: p[3] }, line: { width: 0 } });
    s.addText(p[0], { x: x + 0.16, y: y + 0.22, w: cw - 0.32, h: 0.5, fontSize: 15, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('pv', x + 0.16, y + 0.22, cw - 0.32, 0.5);
    s.addText(p[1], { x: x + 0.16, y: y + 0.84, w: cw - 0.32, h: 0.85, fontSize: 13.5, bold: true, color: P.deep, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('pvq', x + 0.16, y + 0.84, cw - 0.32, 0.85, p[1], 13.5, { pad: 0.02 });
    s.addText(p[2], { x: x + 0.18, y: y + 1.72, w: cw - 0.36, h: 0.72, fontSize: 12.5, color: P.muted, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('pve', x + 0.18, y + 1.72, cw - 0.36, 0.72, p[2], 12.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 2.78, W - 1.24, 1.05, P.soft2);
  s.addText('เคล็ดลับ: ตั้งคำถามเป็นประโยคก่อน แล้วค่อยลาก เช่น อยากรู้ค่าเฉลี่ย PM2.5 ของแต่ละสถานี แยกตามเดือน', {
    x: 0.9, y: y + 2.95, w: W - 1.8, h: 0.72, fontSize: 14.5, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('pvnote', 0.9, y + 2.95, W - 1.8, 0.72, 'เคล็ดลับ: ตั้งคำถามเป็นประโยคก่อน แล้วค่อยลาก เช่น อยากรู้ค่าเฉลี่ย PM2.5 ของแต่ละสถานี แยกตามเดือน', 14.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ผู้เรียนส่วนใหญ่ลากมั่วเพราะไม่ได้ตั้งคำถามก่อน บังคับให้พูดคำถามออกมาดัง ๆ ก่อนแตะเมาส์ จะเร็วขึ้นมาก');
}

/* 14 อ่าน Pivot ให้เป็น */
{
  beginSlide('อ่าน Pivot');
  const s = lightSlide();
  const y = head(s, 3, 'ค่าเฉลี่ยตัวเดียวหลอกได้', 'ดูจำนวนตัวอย่างควบคู่เสมอ นี่คือนิสัยที่แยกนักวิเคราะห์ออกจากคนกดปุ่ม');
  card(s, 0.62, y, W - 1.24, 1.55, P.warn);
  s.addText('สถานการณ์ที่เจอบ่อยที่สุดในรายงานนักศึกษา', { x: 0.9, y: y + 0.18, w: W - 1.8, h: 0.36, fontSize: 15, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('a1', 0.9, y + 0.18, W - 1.8, 0.36);
  s.addText('PivotTable บอกว่าจังหวัดหนึ่งมีค่าเฉลี่ยผลผลิตสูงที่สุด นักศึกษาเขียนลงรายงานทันที แต่พอเปลี่ยนวิธีสรุปเป็น Count กลับพบว่ามีข้อมูลแค่ 2 แถว', {
    x: 0.9, y: y + 0.6, w: W - 1.8, h: 0.8, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('a2', 0.9, y + 0.6, W - 1.8, 0.8, 'PivotTable บอกว่าจังหวัดหนึ่งมีค่าเฉลี่ยผลผลิตสูงที่สุด นักศึกษาเขียนลงรายงานทันที แต่พอเปลี่ยนวิธีสรุปเป็น Count กลับพบว่ามีข้อมูลแค่ 2 แถว', 13.5, { pad: 0.02 });

  const rows = [
    ['ทำอย่างไร', 'ลากตัวแปรเดิมเข้าช่อง Values สองครั้ง ตั้งอันหนึ่งเป็น Average อีกอันเป็น Count'],
    ['อ่านอย่างไร', 'ค่าเฉลี่ยจากข้อมูลน้อยกว่า 5 แถว ให้ถือว่ายังสรุปไม่ได้ ต้องรายงาน n กำกับเสมอ'],
    ['เขียนรายงานอย่างไร', 'เขียนว่า ค่าเฉลี่ย 4,540 กก. จาก 38 แปลง ไม่ใช่เขียนแค่ ค่าเฉลี่ย 4,540 กก.'],
  ];
  const rh = 0.78, gap = 0.14;
  rows.forEach((r, i) => {
    const yy = y + 1.78 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 2 ? P.soft2 : P.soft);
    s.addText(r[0], { x: 0.88, y: yy + 0.1, w: 3.0, h: rh - 0.2, fontSize: 14, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('ak', 0.88, yy + 0.1, 3.0, rh - 0.2, r[0], 14, { pad: 0.02 });
    s.addText(r[1], { x: 4.02, y: yy + 0.1, w: W - 0.62 - 4.02 - 0.26, h: rh - 0.2, fontSize: 13, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('av', 4.02, yy + 0.1, W - 0.62 - 4.02 - 0.26, rh - 0.2, r[1], 13, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('ตัวเลข 4,540 กิโลกรัมมาจากไฟล์ AG จริง ให้ผู้เรียนเปิดตรวจได้ ถ้ามีเวลาให้ลองสร้าง Pivot ที่มีทั้ง Average และ Count พร้อมกัน');
}

/* ภาพค่าเฉลี่ยที่หลอกได้ */
{
  const s = mockPair('ภาพค่าเฉลี่ยหลอก', 3,
    'ตารางสองอันนี้มาจากข้อมูลชุดเดียวกัน',
    'ต่างกันแค่ฝั่งขวาลากฟิลด์เดิมลงพื้นที่ค่าอีกครั้ง แล้วเปลี่ยนเป็นจำนวนนับ',
    'สรุปด้วยค่าเฉลี่ยอย่างเดียว', 'สรุปพร้อมจำนวนตัวอย่าง',
    { cols: ['ตำบล', 'ผลผลิตเฉลี่ย'],
      rows: [['ลุ่มสุ่ม', ['1205', 'n']], ['หนองโรง', ['1265', 'n']],
             [['วังด้ง', 'x'], ['1890', 'xn']], ['แก่งเสี้ยน', ['1150', 'n']]] },
    { cols: ['ตำบล', 'ผลผลิตเฉลี่ย', 'จำนวนแปลง'],
      rows: [['ลุ่มสุ่ม', ['1205', 'n'], ['84', 'n']], ['หนองโรง', ['1265', 'n'], ['77', 'n']],
             ['วังด้ง', ['1890', 'n'], ['2', 'n']], ['แก่งเสี้ยน', ['1150', 'n'], ['63', 'n']]] },
    'ตารางซ้ายชวนให้สรุปว่าวังด้งเก่งที่สุด แต่ค่าเฉลี่ยนั้นมาจากสองแปลง จึงเป็นความบังเอิญ การแสดงจำนวนตัวอย่างคู่กับค่าเฉลี่ยเสมอ คือนิสัยที่แยกนักวิเคราะห์ออกจากคนกดปุ่ม');
  s.addNotes('ให้ผู้เรียนลองทำสดในไฟล์ของตนเอง ลากฟิลด์เดียวกันลงพื้นที่ค่าสองครั้งแล้วเปลี่ยนวิธีสรุปอันที่สอง');
}

/* 15 Slicer Refresh */
{
  beginSlide('Slicer');
  const s = lightSlide();
  const y = head(s, 3, 'ทำให้รายงานใช้งานต่อได้เอง', 'สามอย่างที่เปลี่ยน PivotTable จากตารางนิ่ง เป็นรายงานที่คนอื่นเปิดใช้เองได้');
  const items = [
    ['Slicer', 'ปุ่มกรองที่คนอื่นกดเองได้ ไม่ต้องเข้าใจ PivotTable ก็ใช้เป็น เหมาะมากเวลาส่งรายงานให้ผู้บริหารหรืออาจารย์ที่ปรึกษา', P.teal],
    ['Pivot Chart', 'กราฟที่ผูกกับ Pivot เปลี่ยนตัวกรองแล้วกราฟเปลี่ยนตาม ไม่ต้องสร้างกราฟใหม่ทุกครั้ง', P.deep],
    ['Refresh', 'เมื่อข้อมูลดิบเปลี่ยน ต้องกด Refresh เอง PivotTable ไม่อัปเดตอัตโนมัติ นี่คือสาเหตุอันดับหนึ่งที่รายงานแสดงตัวเลขเก่า', P.coral],
  ];
  const ch = 1.32, gap = 0.2;
  items.forEach((it, i) => {
    const yy = y + 0.15 + i * (ch + gap);
    card(s, 0.62, yy, W - 1.24, ch, i === 2 ? P.warn : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.88, y: yy + 0.4, w: 2.3, h: 0.52, rectRadius: 0.07, fill: { color: it[2] }, line: { width: 0 } });
    s.addText(it[0], { x: 0.88, y: yy + 0.4, w: 2.3, h: 0.52, fontSize: 14.5, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sl', 0.88, yy + 0.4, 2.3, 0.52);
    s.addText(it[1], { x: 3.42, y: yy + 0.16, w: W - 0.62 - 3.42 - 0.26, h: ch - 0.32, fontSize: 13.5, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('slv', 3.42, yy + 0.16, W - 0.62 - 3.42 - 0.26, ch - 0.32, it[1], 13.5, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('เรื่อง Refresh ต้องย้ำหนัก ๆ เพราะเป็นข้อผิดพลาดที่เกิดกับทุกคนอย่างน้อยหนึ่งครั้ง และตรวจจับยากเพราะตัวเลขเก่าก็ดูสมเหตุสมผล');
}

/* 16 WS2.3 */
{
  beginSlide('WS2.3');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 2.3 · สรุปด้วย PivotTable', 'ใช้เวลา 25 นาที ทำเป็นคู่');
  card(s, 0.62, y, W - 1.24, 2.45, P.soft);
  s.addText('สิ่งที่ต้องทำ', { x: 0.9, y: y + 0.2, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w31', 0.9, y + 0.2, W - 1.8, 0.36);
  bullets(s, 0.9, y + 0.62, W - 1.8, 1.7, [
    'สร้าง PivotTable ตอบโจทย์ระดับ Intermediate ในชีตอ่านก่อนเริ่มของไฟล์สาขาท่าน',
    'ต้องแสดงทั้งค่าที่สรุปได้ และจำนวนตัวอย่างของแต่ละกลุ่มควบคู่กัน',
    'เพิ่ม Slicer อย่างน้อยหนึ่งตัว ให้คนอื่นกรองดูเองได้',
    'สร้าง Pivot Chart ที่เลือกชนิดกราฟให้ตรงกับคำถาม',
  ], 13.5);
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.65, cw, 1.45, 'เกณฑ์ผ่าน',
    'Pivot ถูกโครงสร้าง Slicer ใช้งานได้จริง และมีจำนวนตัวอย่างแสดงคู่กับค่าสรุป', P.soft2, 'A06A0A');
  textCard(s, 0.62 + cw + 0.4, y + 2.65, cw, 1.45, 'ข้อควรระวัง',
    'ถ้ายังไม่ได้ล้างข้อมูลจากแบบฝึกหัดก่อนหน้า กลุ่มใน Pivot จะแตกเกินจริง', P.warn, 'A83A36');
  checkOverlaps();
  s.addNotes('ถ้าคู่ไหนเสร็จเร็ว ให้ลองเปลี่ยนวิธีสรุปจากผลรวมเป็นร้อยละของทั้งหมด แล้วถามว่าข้อสรุปเปลี่ยนไปอย่างไร');
}

/* 17 SUMIFS ตระกูล */
{
  beginSlide('SUMIFS');
  const s = lightSlide();
  const y = head(s, 4, 'สรุปแบบมีเงื่อนไขหลายข้อ', 'เมื่อ PivotTable ยังไม่พอ หรือเมื่อต้องการตัวเลขไปใช้ในสูตรอื่นต่อ');
  const rows = [
    ['SUMIFS', 'รวมค่าเมื่อเข้าเงื่อนไขทุกข้อ', 'ช่วงที่จะบวก มาก่อน แล้วตามด้วยคู่ของช่วงเงื่อนไขกับเงื่อนไข', P.teal],
    ['COUNTIFS', 'นับจำนวนแถวที่เข้าเงื่อนไขทุกข้อ', 'ไม่มีช่วงที่จะบวก มีแต่คู่ของเงื่อนไข', P.deep],
    ['AVERAGEIFS', 'เฉลี่ยเฉพาะแถวที่เข้าเงื่อนไขทุกข้อ', 'ลำดับเหมือน SUMIFS คือช่วงที่จะเฉลี่ยมาก่อน', P.amber],
  ];
  const rh = 1.0, gap = 0.16;
  rows.forEach((r, i) => {
    const yy = y + 0.1 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.26, w: 2.6, h: 0.48, rectRadius: 0.07, fill: { color: r[3] }, line: { width: 0 } });
    s.addText(r[0], { x: 0.86, y: yy + 0.26, w: 2.6, h: 0.48, fontSize: 14, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sk', 0.86, yy + 0.26, 2.6, 0.48);
    s.addText(r[1], { x: 3.7, y: yy + 0.12, w: 3.5, h: rh - 0.24, fontSize: 13.5, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('sd', 3.7, yy + 0.12, 3.5, rh - 0.24, r[1], 13.5, { pad: 0.02 });
    s.addText(r[2], { x: 7.4, y: yy + 0.12, w: W - 0.62 - 7.4 - 0.26, h: rh - 0.24, fontSize: 13, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('sv', 7.4, yy + 0.12, W - 0.62 - 7.4 - 0.26, rh - 0.24, r[2], 13, { pad: 0.02 });
  });
  card(s, 0.62, y + 0.1 + 3 * (rh + gap) + 0.05, W - 1.24, 0.92, P.warn);
  s.addText('กับดักที่พบบ่อย: AVERAGEIF วางช่วงเงื่อนไขไว้ก่อน แต่ AVERAGEIFS วางช่วงที่จะเฉลี่ยไว้ก่อน สลับแล้วได้ผลผิดโดยไม่มี error', {
    x: 0.9, y: y + 0.1 + 3 * (rh + gap) + 0.2, w: W - 1.8, h: 0.62, fontSize: 13.5, bold: true, color: 'A83A36', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('swarn', 0.9, y + 0.1 + 3 * (rh + gap) + 0.2, W - 1.8, 0.62, 'กับดักที่พบบ่อย: AVERAGEIF วางช่วงเงื่อนไขไว้ก่อน แต่ AVERAGEIFS วางช่วงที่จะเฉลี่ยไว้ก่อน สลับแล้วได้ผลผิดโดยไม่มี error', 13.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ข้อนี้อยู่ในแบบทดสอบจัดระดับด้วย เป็นข้อที่ผู้เรียนพลาดบ่อยเพราะชื่อฟังก์ชันคล้ายกันมาก');
}

/* 18 ตัวอย่างจริง AC */
{
  beginSlide('ตัวอย่าง AC');
  const s = lightSlide();
  const y = head(s, 4, 'ใช้ SUMIFS หาข้อผิดพลาดที่คนหาไม่เจอ', 'ตัวอย่างจริงจากไฟล์ฝึกปฏิบัติสาขาการบัญชี');
  card(s, 0.62, y, W - 1.24, 1.5, P.soft);
  s.addText('โจทย์: หาว่าใบสำคัญเลขที่ใดบ้างที่ยอดเดบิตไม่เท่ากับยอดเครดิต จากรายการ 389 แถว', {
    x: 0.9, y: y + 0.2, w: W - 1.8, h: 0.4, fontSize: 15, bold: true, color: P.deep, fontFace: FONT, margin: 0
  });
  reg('c1', 0.9, y + 0.2, W - 1.8, 0.4);
  s.addText('วิธีทำ: สร้างรายการเลขที่ใบสำคัญที่ไม่ซ้ำ แล้วใช้ SUMIFS รวมเดบิตและเครดิตของแต่ละใบ จากนั้นหาผลต่างและกรองเฉพาะที่ไม่เป็นศูนย์', {
    x: 0.9, y: y + 0.64, w: W - 1.8, h: 0.7, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('c2', 0.9, y + 0.64, W - 1.8, 0.7, 'วิธีทำ: สร้างรายการเลขที่ใบสำคัญที่ไม่ซ้ำ แล้วใช้ SUMIFS รวมเดบิตและเครดิตของแต่ละใบ จากนั้นหาผลต่างและกรองเฉพาะที่ไม่เป็นศูนย์', 13.5, { pad: 0.02 });

  const sw = (W - 1.24 - 0.6) / 2;
  card(s, 0.62, y + 1.7, sw, 2.2, P.warn);
  s.addText('ถ้าตรวจก่อนลบแถวซ้ำ', { x: 0.86, y: y + 1.86, w: sw - 0.48, h: 0.36, fontSize: 14, bold: true, color: 'A83A36', fontFace: FONT, margin: 0 });
  reg('c3', 0.86, y + 1.86, sw - 0.48, 0.36);
  stat(s, 0.86, y + 2.28, sw - 0.48, '13 ใบ', 'เสียเวลาไล่ตรวจใบที่จริง ๆ ไม่ได้ผิด', P.coral);

  card(s, 0.62 + sw + 0.6, y + 1.7, sw, 2.2, P.soft);
  s.addText('ถ้าลบแถวซ้ำก่อนแล้วค่อยตรวจ', { x: 0.86 + sw + 0.6, y: y + 1.86, w: sw - 0.48, h: 0.36, fontSize: 14, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('c4', 0.86 + sw + 0.6, y + 1.86, sw - 0.48, 0.36);
  stat(s, 0.86 + sw + 0.6, y + 2.28, sw - 0.48, '6 ใบ', 'จำนวนที่ผิดจริง ตรงกับความเป็นจริง', P.teal);

  s.addText('ลำดับการทำงานเปลี่ยนคำตอบ นี่คือบทเรียนที่สำคัญกว่าสูตรใด ๆ ในคาบนี้', {
    x: 0.62, y: y + 4.02, w: W - 1.24, h: 0.44, fontSize: 14.5, bold: true, italic: true, color: 'A06A0A', fontFace: FONT, margin: 0, align: 'center'
  });
  reg('c5', 0.62, y + 4.02, W - 1.24, 0.44);
  checkOverlaps();
  s.addNotes('แบ่งห้องเป็นสองกลุ่ม ให้กลุ่มหนึ่งตรวจดุลก่อน อีกกลุ่มลบซ้ำก่อน แล้วเปรียบเทียบคำตอบกัน บทเรียนจะเข้าเองโดยไม่ต้องบรรยาย');
}

/* ตารางสรุปด้วยสูตร เทียบกับ PivotTable */
{
  const s = mockPair('ภาพตารางสรุปด้วยสูตร', 4,
    'สูตรให้ผลเท่ากับ PivotTable',
    'ต่างกันที่ตารางสูตรเห็นวิธีคำนวณเบื้องหลัง และอัปเดตเองโดยไม่ต้องกดรีเฟรช',
    'PivotTable ที่ต้องกดรีเฟรชเอง', 'ตารางสรุปด้วย SUMIFS และ COUNTIFS',
    { cols: ['ตำบล', 'ผลรวม', 'จำนวนแปลง'],
      rows: [['ลุ่มสุ่ม', ['192879', 'n'], ['43', 'n']],
             ['หนองโรง', ['237749', 'n'], ['56', 'n']],
             ['วังด้ง', ['207174', 'n'], ['48', 'n']],
             [['แก้ข้อมูลดิบแล้ว', 'x'], ['ยังเป็นเลขเก่า', 'x'], ['ต้องกดรีเฟรช', 'x']]] },
    { cols: ['ตำบล', 'ผลรวม', 'จำนวนแปลง'],
      rows: [['ลุ่มสุ่ม', ['192879', 'n'], ['43', 'n']],
             ['หนองโรง', ['237749', 'n'], ['56', 'n']],
             ['วังด้ง', ['207174', 'n'], ['48', 'n']],
             ['แก้ข้อมูลดิบแล้ว', 'เปลี่ยนทันที', 'ไม่ต้องรีเฟรช']] },
    'ในไฟล์เฉลยของทุกสาขามีชีตตารางสรุปแบบ Pivot ที่คำนวณด้วยสองวิธีคู่กัน คือ SUMIFS และ SUMPRODUCT ถ้าสองค่าไม่ตรงกัน แปลว่ามีข้อความปนอยู่ในคอลัมน์ตัวเลข');
  s.addNotes('ให้ผู้เรียนเปิดชีตตารางสรุปแบบ Pivot ในไฟล์เฉลยของสาขาตนเอง แล้วทำ PivotTable จริงมาเทียบว่าตรงกันทุกช่องหรือไม่');
}

/* 19 ฟังก์ชันข้อความและวันที่ */
{
  beginSlide('ข้อความ/วันที่');
  const s = lightSlide();
  const y = head(s, 4, 'ฟังก์ชันข้อความและวันที่ที่ใช้บ่อย', 'ส่วนใหญ่ใช้เพื่อเตรียมข้อมูลให้พร้อมสรุป ไม่ใช่เพื่อความสวยงาม');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.3, P.soft);
  s.addText('ฟังก์ชันข้อความ', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('x1', 0.86, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86, y + 0.68, cw - 0.48, 2.45, [
    'TRIM ตัดช่องว่างหัวท้าย ใช้ก่อนจับคู่ข้อมูลเสมอ',
    'LEFT, RIGHT, MID แยกรหัสที่มีความหมายซ่อนอยู่',
    'TEXTJOIN รวมหลายคอลัมน์เป็นรหัสเดียวไว้ใช้จับคู่',
    'UPPER หรือ LOWER ทำตัวพิมพ์ให้เหมือนกันก่อนจัดกลุ่ม',
  ], 13.5);

  card(s, 0.62 + cw + 0.4, y, cw, 3.3, P.soft);
  s.addText('ฟังก์ชันวันที่', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('x2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86 + cw + 0.4, y + 0.68, cw - 0.48, 2.45, [
    'YEAR, MONTH แยกปีและเดือนออกมาเป็นคอลัมน์ช่วย',
    'TODAY ใช้คำนวณอายุหรือจำนวนวันที่ผ่านไป',
    'DATEDIF หาผลต่างระหว่างวันที่เป็นปี เดือน หรือวัน',
    'ระวังวันที่ที่เป็นข้อความ ต้องแปลงก่อนจึงจะคำนวณได้',
  ], 13.5);

  card(s, 0.62, y + 3.5, W - 1.24, 0.85, P.soft2);
  s.addText('ในไฟล์ฝึกปฏิบัติทุกสาขามีวันที่แบบ พ.ศ. ที่เป็นข้อความปนอยู่ ต้องแปลงก่อนจึงจะแยกเดือนได้', {
    x: 0.9, y: y + 3.64, w: W - 1.8, h: 0.58, fontSize: 14, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('x3', 0.9, y + 3.64, W - 1.8, 0.58, 'ในไฟล์ฝึกปฏิบัติทุกสาขามีวันที่แบบ พ.ศ. ที่เป็นข้อความปนอยู่ ต้องแปลงก่อนจึงจะแยกเดือนได้', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('อย่าสอนฟังก์ชันเหล่านี้แบบท่องรายการ ให้โยงกับปัญหาที่ผู้เรียนเพิ่งเจอในไฟล์ของตัวเองแทน');
}

/* 20 Data Validation */
{
  beginSlide('Validation');
  const s = lightSlide();
  const y = head(s, 5, 'ป้องกันความผิดพลาดตั้งแต่ต้นทาง', 'ถูกกว่าการตามแก้ทีหลังหลายเท่า');
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y, cw, 1.85, 'Data Validation',
    'กำหนดว่าเซลล์นี้รับค่าอะไรได้บ้าง เช่น เลือกจากรายการที่กำหนด หรือรับเฉพาะตัวเลขที่ไม่ติดลบ ผู้กรอกจะกรอกผิดไม่ได้ตั้งแต่แรก', P.soft, P.deep);
  textCard(s, 0.62 + cw + 0.4, y, cw, 1.85, 'ป้องกันชีต',
    'ล็อกเซลล์ที่มีสูตรไว้ ปลดล็อกเฉพาะช่องที่ต้องกรอก คนอื่นจึงแก้สูตรของเราพังไม่ได้แม้จะไม่ตั้งใจ', P.soft, P.deep);

  card(s, 0.62, y + 2.05, W - 1.24, 1.85, P.soft2);
  s.addText('เมื่อไรควรใช้', { x: 0.9, y: y + 2.2, w: W - 1.8, h: 0.36, fontSize: 15, bold: true, color: 'A06A0A', fontFace: FONT, margin: 0 });
  reg('d1', 0.9, y + 2.2, W - 1.8, 0.36);
  bullets(s, 0.9, y + 2.62, W - 1.8, 1.15, [
    'ทุกครั้งที่ทำไฟล์ให้คนอื่นกรอก เช่น แบบฟอร์มบันทึกแปลงทดลองหรือแบบบันทึก QC',
    'ทุกครั้งที่ไฟล์จะถูกใช้ต่อหลายเดือน เพราะคนกรอกอาจเปลี่ยนคน',
    'ในงานวิจัยที่มีผู้ช่วยหลายคนบันทึกข้อมูล ซึ่งเป็นต้นเหตุของชื่อที่สะกดไม่ตรงกัน',
  ], 13.5);
  checkOverlaps();
  s.addNotes('โยงกลับไปที่ไฟล์ BA ที่มีชื่อสาขาเจ็ดแบบ แล้วถามว่าถ้าตั้ง Data Validation ตั้งแต่แรก ปัญหานี้จะเกิดไหม');
}

/* 21 Google Sheets */
{
  beginSlide('Sheets');
  const s = lightSlide();
  const y = head(s, 5, 'งานเดียวกันใน Google Sheets', 'บางอย่าง Sheets ทำได้กระชับกว่ามาก จึงควรรู้ไว้ทั้งสองทาง');
  const rows = [
    ['ดึงรายการที่ไม่ซ้ำ', 'Remove Duplicates หรือ UNIQUE ใน Excel 365', 'UNIQUE(A2:A)'],
    ['กรองข้อมูลแบบสด', 'FILTER ใน 365 หรือ Advanced Filter', 'FILTER(A2:D, C2:C>50)'],
    ['สรุปแบบภาษาคิวรี', 'ต้องใช้ Power Query', 'QUERY(A1:D, "select B, avg(C) group by B")'],
    ['ดึงข้อมูลข้ามไฟล์', 'เชื่อมไฟล์ซึ่งเปราะบางมาก', 'IMPORTRANGE("URL","Sheet1!A:D")'],
    ['คำนวณทั้งคอลัมน์', 'ลากสูตรหรือใช้ตาราง', 'ARRAYFORMULA(B2:B*C2:C)'],
  ];
  const rh = 0.72, gap = 0.13;
  rows.forEach((r, i) => {
    const yy = y + 0.05 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 2 ? P.soft2 : P.soft);
    s.addText(r[0], { x: 0.88, y: yy + 0.08, w: 3.0, h: rh - 0.16, fontSize: 13.5, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('gk', 0.88, yy + 0.08, 3.0, rh - 0.16, r[0], 13.5, { pad: 0.02 });
    s.addText(r[1], { x: 4.02, y: yy + 0.08, w: 4.1, h: rh - 0.16, fontSize: 12.5, color: P.muted, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('gx', 4.02, yy + 0.08, 4.1, rh - 0.16, r[1], 12.5, { pad: 0.02 });
    s.addText(r[2], { x: 8.3, y: yy + 0.08, w: W - 0.62 - 8.3 - 0.26, h: rh - 0.16, fontSize: 12.5, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('gs', 8.3, yy + 0.08, W - 0.62 - 8.3 - 0.26, rh - 0.16, r[2], 12.5, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('QUERY เป็นตัวที่คุ้มที่สุดที่จะสอน ถ้ามีเวลาเหลือให้สาธิตสัก 5 นาที เพราะทำงานที่ต้องใช้ Power Query ได้ในบรรทัดเดียว');
}

/* จดหมายเวียน */
{
  const s = mockPair('ภาพจดหมายเวียน', 6,
    'จดหมายเวียน ทำหนังสือรายคนจากตารางเดียว',
    'ต่อยอดโดยตรงจากตารางที่สะอาดแล้ว และเป็นงานที่คนพิมพ์มือมากที่สุดในสำนักงาน',
    'ตารางรายชื่อที่ใช้ทำจดหมายเวียนไม่ได้', 'ตารางที่ใช้ได้',
    { cols: ['A', 'B', 'C'],
      rows: [[['รายชื่อผู้เข้าอบรม ปี 2569', 'x'], ['', 'x'], ['', 'x']],
             ['', '', ''],
             [['ชื่อ', 'x'], ['นามสกุล', 'x'], ['ระดับ', 'x']],
             ['สมชาย', 'ใจดี', 'Beginner'],
             [['Word เห็นชื่อคอลัมน์', 'x'], ['เป็น F1 F2 F3', 'x'], ['', 'x']]] },
    { cols: ['ชื่อ', 'นามสกุล', 'ระดับ'],
      rows: [['สมชาย', 'ใจดี', 'Beginner'],
             ['สุภาพร', 'แสงทอง', 'Intermediate'],
             ['ธนกฤต', 'บุญมี', 'Advanced'],
             ['หัวตารางอยู่แถวที่ 1', 'ข้อมูลเริ่มแถว 2', 'ไม่มีแถวว่างคั่น']] },
    'Word อ่านแถวแรกของชีตเป็นชื่อคอลัมน์เสมอ หัวเรื่องต้องไปอยู่ชีตอื่น และวันที่กับตัวเลขต้องเติมรหัสรูปแบบใน Word เพราะ Word อ่านค่าดิบ ไม่ได้อ่านรูปแบบที่ Excel แสดง');
  s.addNotes('สาธิตสดสิบนาที ใช้ไฟล์แม่แบบจดหมายเวียนกับไฟล์รายชื่อผู้รับที่แจกไว้ เน้นสองจุด คือหัวตารางแถวแรก และรหัสรูปแบบวันที่กับตัวเลข');
}

/* 22 WS2.4 */
{
  beginSlide('WS2.4');
  const s = lightSlide();
  const y = head(s, null, 'แบบฝึกหัด 2.4 · รายงานวิเคราะห์หนึ่งหน้า', 'ชิ้นงานประเมินหลักของระดับนี้ ใช้เวลา 20 นาที');
  card(s, 0.62, y, W - 1.24, 1.7, P.soft);
  s.addText('โจทย์', { x: 0.9, y: y + 0.18, w: W - 1.8, h: 0.36, fontSize: 16, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('f1', 0.9, y + 0.18, W - 1.8, 0.36);
  s.addText('จากผลสรุปที่ทำในแบบฝึกหัดก่อนหน้า เขียนข้อสรุป 5 ถึง 8 บรรทัดตอบคำถามของสาขาท่าน โดยอ้างตัวเลขจากตารางประกอบทุกข้อความที่ยืนยัน', {
    x: 0.9, y: y + 0.6, w: W - 1.8, h: 0.9, fontSize: 13.5, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('f2', 0.9, y + 0.6, W - 1.8, 0.9, 'จากผลสรุปที่ทำในแบบฝึกหัดก่อนหน้า เขียนข้อสรุป 5 ถึง 8 บรรทัดตอบคำถามของสาขาท่าน โดยอ้างตัวเลขจากตารางประกอบทุกข้อความที่ยืนยัน', 13.5, { pad: 0.02 });

  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 1.9, cw, 2.1, 'สิ่งที่ทำให้ได้คะแนนดี',
    'ระบุจำนวนตัวอย่างกำกับทุกค่าเฉลี่ย บอกว่าตัดข้อมูลอะไรทิ้งไปบ้างและเพราะอะไร และระบุข้อจำกัดของข้อสรุปไว้ท้ายรายงาน',
    P.soft, P.deep);
  textCard(s, 0.62 + cw + 0.4, y + 1.9, cw, 2.1, 'สิ่งที่ทำให้เสียคะแนน',
    'บรรยายตัวเลขซ้ำกับที่อยู่ในตารางโดยไม่เพิ่มความหมาย และสรุปเกินหลักฐาน เช่น บอกว่าดีขึ้นทั้งที่ข้อมูลมีแค่สองเดือน',
    P.warn, 'A83A36');
  checkOverlaps();
  s.addNotes('แจก rubric ให้ดูก่อนเริ่มทำ ไม่ใช่ตอนตรวจ เพราะ rubric ที่เห็นล่วงหน้าทำหน้าที่เป็นเครื่องมือเรียนรู้');
}

/* 23 สรุป */
{
  beginSlide('สรุป');
  const s = lightSlide();
  const y = head(s, null, 'สามอย่างที่อยากให้จำ');
  const items = [
    ['1', 'กด Ctrl+T ก่อนทำอย่างอื่นเสมอ', 'ตารางแบบมีโครงสร้างทำให้สูตรและ PivotTable ขยายตัวเอง ลดข้อผิดพลาดจากการลืมขยายช่วงซึ่งเป็นสาเหตุอันดับหนึ่งของรายงานที่ผิด', P.teal],
    ['2', 'ค่าเฉลี่ยต้องมาคู่กับจำนวนตัวอย่าง', 'ค่าเฉลี่ยจากสองแถวกับจากสองร้อยแถวหน้าตาเหมือนกันในตาราง แต่มีน้ำหนักต่างกันคนละโลก', P.amber],
    ['3', 'ลำดับการทำงานเปลี่ยนคำตอบ', 'ล้างข้อมูลก่อนสรุปเสมอ ตัวอย่างใบสำคัญ 13 ใบกับ 6 ใบคือหลักฐานที่ชัดที่สุด', P.deep],
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

/* 24 ปิด */
{
  beginSlide('ปิด');
  const s = darkSlide();
  s.addShape(pres.ShapeType.ellipse, { x: 10.8, y: 4.2, w: 3.8, h: 3.8, fill: { color: P.deep }, line: { width: 0 } });
  s.addText('แบบทดสอบหลังเรียน', { x: 0.9, y: 2.05, w: 9.5, h: 0.8, fontSize: 34, bold: true, color: P.white, fontFace: FONT, margin: 0 });
  reg('e1', 0.9, 2.05, 9.5, 0.8);
  s.addText('ใช้เวลา 5 นาที ทำโดยไม่เปิดเอกสาร เพื่อให้เห็นพัฒนาการของตัวเองได้จริง', {
    x: 0.9, y: 2.95, w: 9.5, h: 0.5, fontSize: 16, color: 'B9DBD5', fontFace: FONT, margin: 0
  });
  reg('e2', 0.9, 2.95, 9.5, 0.5, 'ใช้เวลา 5 นาที ทำโดยไม่เปิดเอกสาร เพื่อให้เห็นพัฒนาการของตัวเองได้จริง', 16, { pad: 0.05 });
  s.addText('พร้อมสำหรับระดับถัดไปหรือยัง', { x: 0.9, y: 3.85, w: 9.5, h: 0.5, fontSize: 19, bold: true, color: P.teal, fontFace: FONT, margin: 0 });
  reg('e3', 0.9, 3.85, 9.5, 0.5);
  s.addText('ระดับ Advanced with AI จะสอนวิธีทำงานร่วมกับ AI และสร้างระบบอัตโนมัติ  ·  ส่งแบบฝึกหัดทั้งหมดผ่าน Google Classroom ภายในหนึ่งสัปดาห์', {
    x: 0.9, y: 4.4, w: 9.8, h: 0.7, fontSize: 13, color: '7FAEA6', fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('e4', 0.9, 4.4, 9.8, 0.7, 'ระดับ Advanced with AI จะสอนวิธีทำงานร่วมกับ AI และสร้างระบบอัตโนมัติ  ·  ส่งแบบฝึกหัดทั้งหมดผ่าน Google Classroom ภายในหนึ่งสัปดาห์', 13, { pad: 0.05 });
  checkOverlaps();
  s.addNotes('ชวนให้ลงระดับ Advanced ต่อ โดยเฉพาะผู้ที่ทำแบบฝึกหัด 2.4 ได้ดี เพราะแสดงว่าตีความข้อมูลเป็นแล้ว');
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
    console.log('บันทึกระยะเลื่อนลงไฟล์ deck_offsets_int.json แล้ว');
  }
  if (n > 0) process.exitCode = 1;
});
