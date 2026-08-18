/**
 * สไลด์สอนระดับ Advanced with AI (3 ชั่วโมง)
 * ชุดการเรียนรู้ Excel & Google Sheets — อาจารย์ ดร.เนติยา การะเกตุ
 * อ้างอิงตัวอย่างจากไฟล์ฝึกปฏิบัติ 7 สาขาที่สร้างไว้ ตัวเลขทุกตัวดึงจาก answer_key.json
 */
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const L = require('./deck_lib.js');
const { W, H, P, FONT, startSlide, checkOverlaps, report } = L;

const key = JSON.parse(fs.readFileSync(__dirname + '/answer_key.json', 'utf8'));
const meta = JSON.parse(fs.readFileSync(__dirname + '/samples_meta.json', 'utf8'));

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Dr. Netiya Karaket';
pres.title = 'Advanced with AI — ออกแบบระบบให้ข้อมูลทำงานแทนเรา';


/* ------------------------------------------------------------------
   จัดเนื้อหาให้สมดุลแนวตั้ง
   รอบแรกวัดว่าแต่ละสไลด์ใช้พื้นที่ถึงไหน แล้วบันทึกระยะเลื่อนลงไว้ในไฟล์
   รอบที่สองอ่านค่ากลับมาเลื่อนเนื้อหาลง เพื่อไม่ให้พื้นที่ว่างไปกองอยู่ด้านล่าง
   เลื่อนอย่างเดียว ไม่ยืดขนาด วงกลมจึงยังกลมและข้อความไม่มีทางล้นเพิ่ม
   ------------------------------------------------------------------ */
const OFFSET_FILE = __dirname + '/deck_offsets.json';
const OFFSETS = fs.existsSync(OFFSET_FILE) ? JSON.parse(fs.readFileSync(OFFSET_FILE, 'utf8')) : {};
const HEAD_BAND = 1.1;          // สูงกว่านี้ถือเป็นส่วนหัว ไม่เลื่อน
let curOffset = 0;
let curNo = 0;

const regRaw = L.reg;
function reg(kind, x, y, w, h, text, size, opts) {
  return regRaw(kind, x, y >= HEAD_BAND ? y + curOffset : y, w, h, text, size, opts);
}
function shiftOpts(o) {
  if (o && typeof o.y === 'number' && o.y >= HEAD_BAND) {
    const c = Object.assign({}, o);
    c.y = o.y + curOffset;
    return c;
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

const shadow = () => ({ type: 'outer', color: '2A1A4A', blur: 10, offset: 2, angle: 90, opacity: 0.10 });

/* ============================================================ ชิ้นส่วนที่ใช้ซ้ำ */
function darkSlide() {
  const s = pres.addSlide();
  s.background = { color: P.ink };
  return s;
}
function lightSlide() {
  const s = pres.addSlide();
  s.background = { color: P.white };
  return s;
}

/** หัวสไลด์: เลขโมดูลในวงกลม + ชื่อเรื่อง */
function head(s, num, title, sub) {
  const x = 0.62, y = 0.42;
  if (num) {
    s.addShape(pres.ShapeType.ellipse, {
      x, y: y + 0.02, w: 0.52, h: 0.52, fill: { color: P.teal }, line: { color: P.teal }
    });
    s.addText(String(num), {
      x, y: y + 0.02, w: 0.52, h: 0.52, align: 'center', valign: 'middle',
      fontSize: 20, bold: true, color: P.white, fontFace: FONT, margin: 0
    });
    reg('num', x, y, 0.52, 0.56);
  }
  const tx = num ? x + 0.72 : x;
  const tw = W - tx - 0.62;
  s.addText(title, {
    x: tx, y, w: tw, h: 0.62, fontSize: 28, bold: true, color: P.deep, fontFace: FONT,
    valign: 'middle', margin: 0
  });
  reg('title', tx, y, tw, 0.62, title, 28, { pad: 0.05 });
  if (sub) {
    s.addText(sub, {
      x: tx, y: y + 0.62, w: tw, h: 0.34, fontSize: 14.2, color: P.muted, fontFace: FONT,
      valign: 'top', margin: 0
    });
    reg('sub', tx, y + 0.62, tw, 0.34, sub, 13, { pad: 0.05 });
  }
  return y + (sub ? 1.06 : 0.78);
}

/** การ์ดพื้นอ่อน */
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.09,
    fill: { color: fill || P.soft }, line: { width: 0 }, shadow: shadow()
  });
  reg('card', x, y, w, h, null, null, { soft: true });
}

/** การ์ดหัวข้อ + เนื้อหา */
function textCard(s, x, y, w, h, title, bodyText, fill, titleColor) {
  card(s, x, y, w, h, fill);
  const px = x + 0.24, pw = w - 0.48;
  s.addText(title, {
    x: px, y: y + 0.16, w: pw, h: 0.42, fontSize: 16.2, bold: true,
    color: titleColor || P.deep, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('cardTitle', px, y + 0.16, pw, 0.42, title, 15, { pad: 0.02 });
  const bh = h - 0.72;
  s.addText(bodyText, {
    x: px, y: y + 0.62, w: pw, h: bh, fontSize: 14.2, color: P.ink70, fontFace: FONT,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.25
  });
  reg('cardBody', px, y + 0.62, pw, bh, bodyText, 13, { pad: 0.02 });
}

/** รายการหัวข้อย่อยแบบมีจุด */
function bullets(s, x, y, w, h, items, size) {
  size = size || 14;
  const runs = items.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < items.length - 1 }
  }));
  s.addText(runs, {
    x, y, w, h, fontSize: size, color: P.ink70, fontFace: FONT,
    margin: 0, valign: 'top', paraSpaceAfter: 6
  });
  const joined = items.join('\n');
  reg('bullets', x, y, w, h, joined, size, { pad: 0.35 });
}

/** ตัวเลขใหญ่ + คำอธิบาย */
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
  s.addText(big, {
    x, y, w, h: 0.9, fontSize: 46, bold: true, color: color || P.deep,
    fontFace: FONT, align: 'center', margin: 0, valign: 'middle'
  });
  reg('statBig', x, y, w, 0.9);
  s.addText(label, {
    x, y: y + 0.92, w, h: 0.62, fontSize: 13.7, color: P.muted, fontFace: FONT,
    align: 'center', margin: 0, valign: 'top'
  });
  reg('statLabel', x, y + 0.92, w, 0.62, label, 12.5, { pad: 0.1 });
}

/** แถบเวลาในผังการสอน */
function timeRow(s, x, y, w, time, topic, detail, accent) {
  const tw = 1.15;
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w: tw, h: 0.46, rectRadius: 0.07,
    fill: { color: accent || P.violet }, line: { width: 0 }
  });
  s.addText(time, {
    x, y, w: tw, h: 0.46, fontSize: 12.7, bold: true, color: P.white,
    fontFace: FONT, align: 'center', valign: 'middle', margin: 0
  });
  reg('time', x, y, tw, 0.46);
  s.addText(topic, {
    x: x + tw + 0.16, y, w: 3.95, h: 0.46, fontSize: 14.2, bold: true, color: P.deep,
    fontFace: FONT, valign: 'middle', margin: 0
  });
  reg('timeTopic', x + tw + 0.16, y, 3.95, 0.46, topic, 13, { pad: 0.02 });
  const dx = x + tw + 0.16 + 3.95 + 0.14;
  const dw = x + w - dx;
  s.addText(detail, {
    x: dx, y, w: dw, h: 0.46, fontSize: 12.7, color: P.muted,
    fontFace: FONT, valign: 'middle', margin: 0
  });
  reg('timeDetail', dx, y, dw, 0.46, detail, 11.5, { pad: 0.02 });
}

function footer(s, txt) {
  s.addText(txt, {
    x: 0.62, y: H - 0.52, w: W - 1.24, h: 0.3, fontSize: 11.2, color: P.line,
    fontFace: FONT, margin: 0, valign: 'middle'
  });
}

/* ============================================================ สไลด์ */

/* 1 — ปก */
{
  beginSlide('ปก');
  const s = wrapSlide(darkSlide());
  s.addShape(pres.ShapeType.ellipse, { x: 10.4, y: -1.5, w: 5.2, h: 5.2, fill: { color: P.deep }, line: { width: 0 } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.9, y: 4.6, w: 2.6, h: 2.6, fill: { color: P.violet }, line: { width: 0 } });
  s.addText('ระดับที่ 3', { x: 0.9, y: 1.75, w: 6, h: 0.4, fontSize: 16.2, color: P.teal, fontFace: FONT, margin: 0, charSpacing: 2 });
  reg('eyebrow', 0.9, 1.75, 6, 0.4);
  s.addText('Advanced with AI', { x: 0.9, y: 2.2, w: 8.6, h: 1.0, fontSize: 44, bold: true, color: P.white, fontFace: FONT, margin: 0 });
  reg('h1', 0.9, 2.2, 8.6, 1.0);
  s.addText('ออกแบบระบบให้ข้อมูลทำงานแทนเรา', { x: 0.9, y: 3.22, w: 8.6, h: 0.6, fontSize: 23, color: 'CFC4E8', fontFace: FONT, margin: 0 });
  reg('h2', 0.9, 3.22, 8.6, 0.6, 'ออกแบบระบบให้ข้อมูลทำงานแทนเรา', 23, { pad: 0.05 });
  s.addText('3 ชั่วโมง รวมทำแบบฝึกหัด · ใช้ชุดข้อมูลจริงของสาขาตนเอง', { x: 0.9, y: 3.95, w: 8.6, h: 0.4, fontSize: 15.2, color: '9C8FC0', fontFace: FONT, margin: 0 });
  reg('h3', 0.9, 3.95, 8.6, 0.4);
  s.addText('อาจารย์ ดร.เนติยา การะเกตุ  ·  หลักสูตรวิทยาศาสตร์การเกษตร  ·  ม.มหิดล วิทยาเขตกาญจนบุรี',
    { x: 0.9, y: 5.9, w: 10, h: 0.4, fontSize: 13.2, color: '7E71A5', fontFace: FONT, margin: 0 });
  reg('by', 0.9, 5.9, 10, 0.4);
  checkOverlaps();
  s.addNotes('เปิดคาบด้วยคำถาม ใครเคยให้ AI ช่วยเขียนสูตรแล้วได้สูตรที่ใช้ไม่ได้บ้าง แล้วโยงเข้าแนวคิดหลักว่า ระดับนี้ไม่ได้สอนให้พึ่ง AI แต่สอนวิธีทำงานร่วมกับ AI อย่างมีวิจารณญาณ');
}

/* 2 — แนวคิดหลัก */
{
  beginSlide('แนวคิดหลัก');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'แนวคิดหลักของระดับนี้');
  card(s, 0.62, y + 0.15, W - 1.24, 1.9, P.soft);
  s.addText('AI เป็นผู้ช่วยที่เก่ง แต่ไม่รับผิดชอบ', {
    x: 1.0, y: y + 0.45, w: W - 2.0, h: 0.66, fontSize: 30, bold: true, color: P.deep, fontFace: FONT, margin: 0, align: 'center'
  });
  reg('big', 1.0, y + 0.45, W - 2.0, 0.66);
  s.addText('ผู้ตรวจสอบคนสุดท้ายคือคุณเสมอ ไม่ใช่เครื่องมือ', {
    x: 1.0, y: y + 1.18, w: W - 2.0, h: 0.5, fontSize: 17.2, color: P.muted, fontFace: FONT, margin: 0, align: 'center'
  });
  reg('bigsub', 1.0, y + 1.18, W - 2.0, 0.5);

  const cy = y + 2.3, cw = (W - 1.24 - 0.6) / 3, ch = 2.0;
  textCard(s, 0.62, cy, cw, ch, 'สิ่งที่เปลี่ยนไป',
    'งานที่เคยใช้เวลาเป็นชั่วโมง เช่น เขียนสูตรซับซ้อนหรือร่างสคริปต์ ทำได้ในไม่กี่นาที', P.soft);
  textCard(s, 0.62 + cw + 0.3, cy, cw, ch, 'สิ่งที่ไม่เปลี่ยน',
    'ความรับผิดชอบต่อความถูกต้องของผลลัพธ์ยังอยู่ที่ผู้ใช้ทั้งหมด AI ไม่เคยรับผิดแทนใคร', P.warn);
  textCard(s, 0.62 + (cw + 0.3) * 2, cy, cw, ch, 'ทักษะที่หายากขึ้น',
    'ความสามารถในการมองออกว่าผลลัพธ์ที่ดูดีนั้นผิด เป็นทักษะที่มีค่ามากกว่าเดิมหลายเท่า', P.soft2);
  checkOverlaps();
  s.addNotes('ย้ำว่าสามการ์ดนี้คือเหตุผลที่ระดับ Advanced ไม่ได้สอนเทคนิค AI แต่สอนวิจารณญาณ ให้เวลาผู้เรียนอ่านเองสัก 30 วินาทีก่อนอธิบาย');
}

/* 3 — ผลลัพธ์การเรียนรู้ */
{
  beginSlide('CLO');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'เมื่อจบคาบนี้ คุณจะทำอะไรได้', 'ผลลัพธ์การเรียนรู้ทั้งหกข้อ วัดผลด้วยแบบฝึกหัดและแบบทดสอบหลังเรียน');
  const items = [
    ['3.1', 'อธิบายขอบเขตและข้อจำกัดของ AI ในงานสเปรดชีต และระบุงานที่ไม่ควรมอบให้ AI ตัดสินใจ'],
    ['3.2', 'เขียนคำสั่งที่ระบุบริบท โครงสร้างข้อมูล และรูปแบบผลลัพธ์ จนได้สูตรที่ใช้ได้ทันที'],
    ['3.3', 'ตรวจสอบความถูกต้องของสูตรหรือโค้ดที่ AI สร้าง ด้วยกรณีทดสอบที่รู้คำตอบล่วงหน้า'],
    ['3.4', 'สร้างระบบอัตโนมัติอย่างง่ายด้วย Apps Script หรือ Power Query โดยใช้ AI ช่วยร่างโค้ด'],
    ['3.5', 'ออกแบบ Dashboard ที่สื่อสารข้อค้นพบได้ตรงประเด็นกับผู้รับสารเป้าหมาย'],
    ['3.6', 'ตัดสินใจได้ว่าข้อมูลใดนำเข้าเครื่องมือ AI ได้หรือไม่ได้ ตามหลักความเป็นส่วนตัวของข้อมูล'],
  ];
  const rh = 0.72, gap = 0.14;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const cw = (W - 1.24 - 0.4) / 2;
    const x = 0.62 + col * (cw + 0.4);
    const yy = y + 0.1 + row * (rh + gap);
    card(s, x, yy, cw, rh, i === 2 ? P.warn : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.16, y: yy + 0.16, w: 0.62, h: 0.4, rectRadius: 0.06, fill: { color: i === 2 ? P.amber : P.violet }, line: { width: 0 } });
    s.addText(it[0], { x: x + 0.16, y: yy + 0.16, w: 0.62, h: 0.4, fontSize: 13.2, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('cloNum', x + 0.16, yy + 0.16, 0.62, 0.4);
    s.addText(it[1], { x: x + 0.92, y: yy + 0.08, w: cw - 1.1, h: rh - 0.16, fontSize: 13.2, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('cloText', x + 0.92, yy + 0.08, cw - 1.1, rh - 0.16, it[1], 12, { pad: 0.02 });
  });
  s.addText('ข้อ 3.3 คือหัวใจของระดับนี้ แบบฝึกหัดล่าสูตรผิดออกแบบมาเพื่อวัดข้อนี้โดยเฉพาะ', {
    x: 0.62, y: y + 0.1 + 3 * (rh + gap) + 0.12, w: W - 1.24, h: 0.4, fontSize: 13.7, italic: true, color: P.amber, fontFace: FONT, margin: 0
  });
  reg('note', 0.62, y + 0.1 + 3 * (rh + gap) + 0.12, W - 1.24, 0.4);
  checkOverlaps();
  s.addNotes('ชี้ให้เห็นว่าข้อ 3.3 เป็นข้อเดียวที่อยู่ระดับ Evaluate ของ Bloom และเป็นข้อที่มีน้ำหนักในข้อสอบหลังเรียนมากที่สุด');
}

/* 4 — ผังเวลา */
{
  beginSlide('ผังเวลา');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'ผังการเรียน 180 นาที');
  const rows = [
    ['0–10', 'แบบทดสอบก่อนเรียน', 'ทำโดยไม่เปิดเอกสาร เพื่อวัดพัฒนาการได้จริง', P.muted],
    ['10–30', 'กรอบคิดการทำงานร่วมกับ AI', 'AI ทำอะไรได้ดี ทำอะไรได้ไม่ดี', P.violet],
    ['30–60', 'รูปแบบคำสั่งสำหรับงานสเปรดชีต', 'สูตรสี่ส่วน และรูปแบบคำสั่งห้าแบบ', P.violet],
    ['60–90', 'การตรวจสอบผลลัพธ์อย่างเป็นระบบ', 'แบบฝึกหัดล่าสูตรผิด หัวใจของคาบนี้', P.amber],
    ['90–100', 'พักเบรก', '', P.muted],
    ['100–135', 'ระบบอัตโนมัติ', 'Apps Script และ Power Query', P.violet],
    ['135–160', 'แบบฝึกหัดรวบยอด', 'สร้างท่อข้อมูลอัตโนมัติของตนเอง', P.teal],
    ['160–175', 'จริยธรรมและการสื่อสารข้อมูล', 'PDPA และกราฟที่ทำให้เข้าใจผิด', P.violet],
    ['175–180', 'แบบทดสอบหลังเรียนและสรุป', '', P.muted],
  ];
  rows.forEach((r, i) => {
    timeRow(s, 0.62, y + 0.08 + i * 0.55, W - 1.24, r[0], r[1], r[2], r[3]);
  });
  checkOverlaps();
  s.addNotes('บอกผู้เรียนล่วงหน้าว่าช่วง 60 ถึง 90 นาทีเป็นช่วงที่สำคัญที่สุด และช่วง 135 ถึง 160 จะได้ใช้ข้อมูลของตัวเอง ให้เตรียมไฟล์มาด้วย');
}

/* 5 — โมดูล 1: AI ทำอะไรได้ ไม่ได้ */
{
  beginSlide('AI ได้/ไม่ได้');
  const s = wrapSlide(lightSlide());
  const y = head(s, 1, 'AI ทำอะไรได้ดี และทำอะไรได้ไม่ดี', 'แยกให้ออกก่อน แล้วจะใช้ AI ได้คุ้มค่าโดยไม่เสี่ยง');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.3, P.soft2);
  s.addText('ทำได้ดี', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 18.2, bold: true, color: '00776A', fontFace: FONT, margin: 0 });
  reg('okTitle', 0.86, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86, y + 0.68, cw - 0.48, 2.4, [
    'ร่างสูตรจากคำอธิบายที่เป็นภาษาคน',
    'อธิบายสูตรยาว ๆ ที่รับมาจากคนอื่นว่าทำอะไร',
    'หาจุดผิดในสูตรเมื่อเราบอกอาการที่เจอ',
    'ร่างโค้ด Apps Script โครงสร้างพื้นฐาน',
    'สร้างชุดข้อมูลจำลองไว้ทดสอบ',
  ], 13);

  card(s, 0.62 + cw + 0.4, y, cw, 3.3, P.warn);
  s.addText('ทำได้ไม่ดี', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 18.2, bold: true, color: 'A85A00', fontFace: FONT, margin: 0 });
  reg('badTitle', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86 + cw + 0.4, y + 0.68, cw - 0.48, 2.4, [
    'คำนวณตัวเลขจำนวนมากในหัว',
    'รู้ว่าข้อมูลจริงของเราหน้าตาเป็นอย่างไร',
    'รู้ว่าเราใช้โปรแกรมเวอร์ชันไหน',
    'ตัดสินใจเชิงวิชาการแทนเรา',
    'รับผิดชอบเมื่อผลลัพธ์ผิด',
  ], 13);

  card(s, 0.62, y + 3.5, W - 1.24, 0.72, P.soft);
  s.addText('หลักง่าย ๆ: ให้ AI ทำงานที่เป็นภาษาและตรรกะ แล้วให้สเปรดชีตเป็นผู้คิดเลข', {
    x: 0.86, y: y + 3.62, w: W - 1.72, h: 0.48, fontSize: 15.7, bold: true, color: P.deep, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('rule', 0.86, y + 3.62, W - 1.72, 0.48, 'หลักง่าย ๆ: ให้ AI ทำงานที่เป็นภาษาและตรรกะ แล้วให้สเปรดชีตเป็นผู้คิดเลข', 14.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ถามผู้เรียนว่าเคยเจอ AI ให้ตัวเลขผิดไหม แล้วอธิบายว่าสาเหตุคือแบบจำลองภาษาไม่ได้ออกแบบมาเพื่อคำนวณ ไม่ใช่เพราะมันโง่');
}

/* 6 — สูตรที่ดูถูกแต่ผิด */
{
  beginSlide('ดูถูกแต่ผิด');
  const s = wrapSlide(lightSlide());
  const y = head(s, 1, 'อันตรายที่สุดคือสูตรที่ดูถูกแต่ผิด', 'สูตรที่พังจะขึ้น error ให้เห็น แต่สูตรที่ผิดเงียบ ๆ จะไม่มีอะไรเตือนเลย');
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y, cw, 1.5, 'สูตรที่พัง',
    'ขึ้น #VALUE! หรือ #REF! เห็นทันที แก้ได้ทันที ไม่อันตราย', P.soft2, '00776A');
  textCard(s, 0.62 + cw + 0.4, y, cw, 1.5, 'สูตรที่ผิดเงียบ ๆ',
    'ให้ตัวเลขที่ดูสมเหตุสมผล เข้ารายงานได้ ผ่านตาผู้ตรวจได้ และไม่มีใครรู้', P.warn, 'A85A00');

  card(s, 0.62, y + 1.75, W - 1.24, 2.4, P.soft);
  s.addText('ตัวอย่างจริงจากไฟล์ฝึกปฏิบัติสาขาวิศวกรรมสิ่งแวดล้อม', {
    x: 0.9, y: y + 1.95, w: W - 1.8, h: 0.36, fontSize: 15.2, bold: true, color: P.deep, fontFace: FONT, margin: 0
  });
  reg('exTitle', 0.9, y + 1.95, W - 1.8, 0.36);
  s.addText('ชุดข้อมูลสถานีตรวจวัดใช้ค่า -999 แทนช่วงที่เซนเซอร์ไม่ทำงาน ถ้าเผลอหาค่าเฉลี่ยจากข้อมูลดิบ',
    { x: 0.9, y: y + 2.32, w: W - 1.8, h: 0.36, fontSize: 14.2, color: P.ink70, fontFace: FONT, margin: 0 });
  reg('exBody', 0.9, y + 2.32, W - 1.8, 0.36);
  const sw = (W - 1.8 - 1.0) / 3;
  stat(s, 0.9, y + 2.78, sw, '−3.1', 'ค่าเฉลี่ย PM2.5 จากข้อมูลดิบ ซึ่งเป็นไปไม่ได้', P.coral);
  stat(s, 0.9 + sw + 0.5, y + 2.78, sw, '30.2', 'ค่าเฉลี่ยที่ถูกต้องหลังตัด -999 ออก', P.teal);
  stat(s, 0.9 + (sw + 0.5) * 2, y + 2.78, sw, '27', 'จำนวนแถวที่มีค่า -999 ซ่อนอยู่', P.deep);
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนลองคำนวณค่าเฉลี่ยจากไฟล์ ED จริงในเครื่องตัวเอง แล้วถามว่าค่าฝุ่นติดลบเป็นไปได้อย่างไร นี่คือช่วงที่ทรงพลังที่สุดของโมดูลนี้');
}

/* 7 — AI ไม่เห็นข้อมูลของเรา */
{
  beginSlide('AI ไม่เห็นข้อมูล');
  const s = wrapSlide(lightSlide());
  const y = head(s, 1, 'AI ไม่เห็นไฟล์ของคุณ มันเดา', 'และมักเดาผิดเรื่องช่วงเซลล์ ชื่อคอลัมน์ และชนิดข้อมูล');
  const cw = (W - 1.24 - 0.5) / 2;
  card(s, 0.62, y, cw, 2.9, P.warn);
  s.addText('ถ้าคุณถามว่า', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.32, fontSize: 14.2, bold: true, color: 'A85A00', fontFace: FONT, margin: 0 });
  reg('q1', 0.86, y + 0.2, cw - 0.48, 0.32);
  s.addText('"หาค่าเฉลี่ย PM2.5 ให้หน่อย"', {
    x: 0.86, y: y + 0.58, w: cw - 0.48, h: 0.5, fontSize: 18.2, bold: true, color: P.ink, fontFace: FONT, margin: 0
  });
  reg('q1t', 0.86, y + 0.58, cw - 0.48, 0.5);
  s.addText('AI จะต้องเดาทั้งหมดนี้', { x: 0.86, y: y + 1.16, w: cw - 0.48, h: 0.32, fontSize: 13.7, color: P.muted, fontFace: FONT, margin: 0 });
  reg('q1s', 0.86, y + 1.16, cw - 0.48, 0.32);
  bullets(s, 0.86, y + 1.52, cw - 0.48, 1.25, [
    'ข้อมูลอยู่คอลัมน์ไหน เริ่มแถวที่เท่าไร',
    'ใช้ Excel หรือ Google Sheets เวอร์ชันใด',
    'มีค่าว่างหรือค่าพิเศษปนอยู่หรือไม่',
  ], 12.5);

  card(s, 0.62 + cw + 0.5, y, cw, 2.9, P.soft2);
  s.addText('สิ่งที่ต้องบอกทุกครั้ง', { x: 0.86 + cw + 0.5, y: y + 0.2, w: cw - 0.48, h: 0.32, fontSize: 14.2, bold: true, color: '00776A', fontFace: FONT, margin: 0 });
  reg('q2', 0.86 + cw + 0.5, y + 0.2, cw - 0.48, 0.32);
  bullets(s, 0.86 + cw + 0.5, y + 0.62, cw - 0.48, 2.15, [
    'โปรแกรมและเวอร์ชันที่ใช้',
    'ชื่อชีตและช่วงแถวของข้อมูล',
    'คอลัมน์ใดเก็บอะไร และเป็นชนิดใด',
    'ค่าพิเศษที่มีความหมาย เช่น -999 หรือ ND',
    'จะวางสูตรที่เซลล์ไหน และต้องลากลงได้หรือไม่',
  ], 12.5);

  card(s, 0.62, y + 3.15, W - 1.24, 0.78, P.soft);
  s.addText('ข้อมูลที่คุณมองว่าชัดเจนอยู่แล้ว คือข้อมูลที่ AI ไม่มีทางรู้ได้เลย', {
    x: 0.86, y: y + 3.3, w: W - 1.72, h: 0.5, fontSize: 15.7, bold: true, color: P.deep, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('note2', 0.86, y + 3.3, W - 1.72, 0.5);
  checkOverlaps();
  s.addNotes('เปรียบเทียบกับการสั่งงานผู้ช่วยที่เก่งมากแต่ตาบอด เขาทำได้ทุกอย่างถ้าเราอธิบายให้ครบ แต่ถ้าไม่บอกเขาก็ต้องเดา');
}

/* ภาพช่วงที่คลาดไปหนึ่งแถว */
{
  const s = mockPair('ภาพช่วงคลาดหนึ่งแถว', 1,
    'ความผิดพลาดที่คลาดไปเพียงหนึ่งแถว',
    'ผลลัพธ์ต่างกันไม่มากจนน่าสงสัย จึงไม่มีใครตรวจพบ',
    'สูตรที่ AI เดาช่วงเอง', 'สูตรที่ระบุช่วงจริง',
    { cols: ['วันที่', 'ค่า PM2.5'],
      rows: [[['2026-01-01', 'x'], ['51 ตกหล่น', 'xn']], ['2026-01-02', ['48', 'n']],
             ['2026-01-03', ['44', 'n']], [['เฉลี่ย AVERAGE(B3:B4)', 'x'], ['46.0', 'xn']]] },
    { cols: ['วันที่', 'ค่า PM2.5'],
      rows: [['2026-01-01', ['51', 'n']], ['2026-01-02', ['48', 'n']],
             ['2026-01-03', ['44', 'n']], ['เฉลี่ย AVERAGE(B2:B4)', ['47.7', 'n']]] },
    'AI ไม่เห็นไฟล์ของเราจึงเดาว่าข้อมูลเริ่มแถวไหน ค่าที่ได้คือ 46.0 กับ 47.7 ซึ่งไม่มีอะไรน่าสงสัยเลย นี่คือเหตุผลที่ต้องทดสอบด้วยข้อมูลชุดเล็กที่รู้คำตอบล่วงหน้า');
  s.addNotes('ให้ทุกคนเปิดไฟล์ของตนเองแล้วตรวจว่าสูตรที่ใช้อยู่ครอบคลุมแถวแรกและแถวสุดท้ายจริงหรือไม่');
}

/* 8 — โมดูล 2: prompt สี่ส่วน */
{
  beginSlide('prompt 4 ส่วน');
  const s = wrapSlide(lightSlide());
  const y = head(s, 2, 'สูตรคำสั่งสี่ส่วน', 'จำสี่ส่วนนี้แล้วใช้ได้กับทุกงาน ไม่ต้องท่องคำสั่งสำเร็จรูป');
  const parts = [
    ['1', 'บริบท', 'ทำงานอะไร ข้อมูลเกี่ยวกับอะไร ใช้โปรแกรมและเวอร์ชันใด', P.violet],
    ['2', 'โครงสร้างข้อมูล', 'คอลัมน์อะไรอยู่ที่ไหน แถวเริ่มที่เท่าไร ชนิดข้อมูลเป็นอะไร', P.teal],
    ['3', 'สิ่งที่ต้องการ', 'ผลลัพธ์ที่ต้องการ พร้อมเงื่อนไขให้ครบทุกข้อ', P.amber],
    ['4', 'รูปแบบคำตอบ', 'ขอสูตรอย่างเดียวหรือขอคำอธิบายด้วย ขอกี่ทางเลือก', P.deep],
  ];
  const cw = (W - 1.24 - 0.45) / 4, ch = 2.7;
  parts.forEach((p, i) => {
    const x = 0.62 + i * (cw + 0.15);
    card(s, x, y, cw, ch, P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: x + (cw - 0.62) / 2, y: y + 0.26, w: 0.62, h: 0.62, fill: { color: p[3] }, line: { width: 0 } });
    s.addText(p[0], { x: x + (cw - 0.62) / 2, y: y + 0.26, w: 0.62, h: 0.62, fontSize: 22, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('pnum', x + (cw - 0.62) / 2, y + 0.26, 0.62, 0.62);
    s.addText(p[1], { x: x + 0.16, y: y + 1.0, w: cw - 0.32, h: 0.4, fontSize: 16.2, bold: true, color: P.deep, fontFace: FONT, align: 'center', margin: 0 });
    reg('pname', x + 0.16, y + 1.0, cw - 0.32, 0.4, p[1], 15, { pad: 0.02 });
    s.addText(p[2], { x: x + 0.18, y: y + 1.45, w: cw - 0.36, h: 1.1, fontSize: 13.2, color: P.ink70, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('pdesc', x + 0.18, y + 1.45, cw - 0.36, 1.1, p[2], 12, { pad: 0.02 });
  });
  card(s, 0.62, y + 2.95, W - 1.24, 0.95, P.soft2);
  s.addText('ขาดส่วนที่ 2 บ่อยที่สุด และเป็นสาเหตุอันดับหนึ่งที่ได้สูตรที่อ้างช่วงเซลล์ผิด', {
    x: 0.86, y: y + 3.12, w: W - 1.72, h: 0.6, fontSize: 15.7, bold: true, color: '00776A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('pnote', 0.86, y + 3.12, W - 1.72, 0.6);
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนเปิดใบงาน Prompt Card แล้วเขียนสี่ส่วนนี้กับโจทย์ของสาขาตนเองทันที ไม่ต้องรอถึงช่วงแบบฝึกหัด');
}

/* 9 — เทียบ prompt */
{
  beginSlide('เทียบ prompt');
  const s = wrapSlide(lightSlide());
  const y = head(s, 2, 'คำสั่งเดียวกัน ต่างกันแค่รายละเอียด', 'ทั้งสองคำสั่งขอสิ่งเดียวกัน แต่ผลลัพธ์ต่างกันคนละเรื่อง');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.9, P.warn);
  s.addText('คำสั่งที่ได้สูตรใช้ไม่ได้', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.36, fontSize: 15.2, bold: true, color: 'A85A00', fontFace: FONT, margin: 0 });
  reg('bt', 0.86, y + 0.2, cw - 0.48, 0.36);
  s.addText('"หาค่าเฉลี่ย PM2.5 ของแต่ละตำบลให้หน่อย"', {
    x: 0.86, y: y + 0.64, w: cw - 0.48, h: 0.62, fontSize: 16.2, italic: true, color: P.ink, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('bq', 0.86, y + 0.64, cw - 0.48, 0.62, '"หาค่าเฉลี่ย PM2.5 ของแต่ละตำบลให้หน่อย"', 15, { pad: 0.02 });
  s.addText('ผลที่ได้', { x: 0.86, y: y + 1.42, w: cw - 0.48, h: 0.3, fontSize: 13.7, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
  reg('bl', 0.86, y + 1.42, cw - 0.48, 0.3);
  bullets(s, 0.86, y + 1.76, cw - 0.48, 1.95, [
    'เดาช่วงเซลล์เอง มักได้ A2:B100 ที่ไม่ตรงกับไฟล์จริง',
    'อาจใช้ฟังก์ชันที่ Excel รุ่นเก่าไม่มี',
    'ไม่รู้ว่ามีค่า -999 ปนอยู่ จึงเฉลี่ยรวมเข้าไปด้วย',
    'ต้องแก้ไปมาหลายรอบกว่าจะใช้ได้',
  ], 12.5);

  card(s, 0.62 + cw + 0.4, y, cw, 3.9, P.soft2);
  s.addText('คำสั่งที่ได้สูตรใช้ได้ทันที', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.36, fontSize: 15.2, bold: true, color: '00776A', fontFace: FONT, margin: 0 });
  reg('gt', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.36);
  s.addText('"ใช้ Google Sheets ชีต Data คอลัมน์ A=ตำบล (ข้อความ) B=วันที่ C=PM2.5 (ตัวเลข) แถว 2 ถึง 5000 มีค่า -999 แทนวันที่เซนเซอร์เสีย ต้องการค่าเฉลี่ยรายตำบลเฉพาะเดือนมกราคม โดยไม่นับ -999 ขอสูตรเดียววางที่ F2 แล้วลากลงได้"', {
    x: 0.86 + cw + 0.4, y: y + 0.64, w: cw - 0.48, h: 2.05, fontSize: 13.2, italic: true, color: P.ink, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('gq', 0.86 + cw + 0.4, y + 0.64, cw - 0.48, 2.05,
    '"ใช้ Google Sheets ชีต Data คอลัมน์ A=ตำบล (ข้อความ) B=วันที่ C=PM2.5 (ตัวเลข) แถว 2 ถึง 5000 มีค่า -999 แทนวันที่เซนเซอร์เสีย ต้องการค่าเฉลี่ยรายตำบลเฉพาะเดือนมกราคม โดยไม่นับ -999 ขอสูตรเดียววางที่ F2 แล้วลากลงได้"', 12, { pad: 0.02 });
  s.addText('ครบทั้งสี่ส่วน จึงได้สูตรที่วางแล้วใช้ได้เลย', {
    x: 0.86 + cw + 0.4, y: y + 2.82, w: cw - 0.48, h: 0.36, fontSize: 13.7, bold: true, color: '00776A', fontFace: FONT, margin: 0
  });
  reg('gn', 0.86 + cw + 0.4, y + 2.82, cw - 0.48, 0.36);
  s.addText('ยาวกว่าจริง แต่เขียนครั้งเดียวแล้วไม่ต้องแก้ซ้ำ ซึ่งเร็วกว่ามากเมื่อรวมเวลาทั้งหมด', {
    x: 0.86 + cw + 0.4, y: y + 3.22, w: cw - 0.48, h: 0.62, fontSize: 13.2, color: P.muted, fontFace: FONT, margin: 0
  });
  reg('gn2', 0.86 + cw + 0.4, y + 3.22, cw - 0.48, 0.62, 'ยาวกว่าจริง แต่เขียนครั้งเดียวแล้วไม่ต้องแก้ซ้ำ ซึ่งเร็วกว่ามากเมื่อรวมเวลาทั้งหมด', 12, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ประเด็นที่ต้องย้ำคือคำสั่งด้านขวาไม่ได้ใช้เทคนิคพิเศษอะไรเลย เพียงบอกสิ่งที่ตัวเองรู้อยู่แล้วให้ครบเท่านั้น');
}

/* 10 — รูปแบบ prompt 5 แบบ */
{
  beginSlide('prompt 5 แบบ');
  const s = wrapSlide(lightSlide());
  const y = head(s, 2, 'รูปแบบคำสั่งที่ใช้บ่อยห้าแบบ', 'ใช้เป็นแม่แบบได้ทันที เปลี่ยนเฉพาะรายละเอียดของงานตนเอง');
  const rows = [
    ['สร้างสูตร', 'อธิบายผลลัพธ์ที่ต้องการเป็นภาษาคน พร้อมโครงสร้างข้อมูล แล้วขอสูตรเดียวที่ลากได้'],
    ['แก้จุดผิด', 'วางสูตรที่มี บอกอาการที่เจอ และบอกผลลัพธ์ที่คาดหวัง แล้วขอให้ชี้ว่าผิดตรงไหนและเพราะอะไร'],
    ['อธิบายสูตรที่รับมา', 'วางสูตรยาวที่ได้จากคนอื่น ขอให้แยกอธิบายทีละส่วนเป็นภาษาคน'],
    ['ออกแบบโครงสร้างตาราง', 'บอกว่าจะเก็บข้อมูลอะไรและจะวิเคราะห์อะไรต่อ ขอให้เสนอคอลัมน์และชนิดข้อมูล'],
    ['สร้างข้อมูลทดสอบ', 'ขอชุดข้อมูลจำลอง 20 แถวที่มีโครงสร้างเหมือนของจริง ไว้ทดสอบสูตรก่อนใช้กับข้อมูลจริง'],
  ];
  const rh = 0.72, gap = 0.13;
  rows.forEach((r, i) => {
    const yy = y + 0.05 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 4 ? P.soft2 : P.soft);
    s.addText(r[0], { x: 0.88, y: yy + 0.08, w: 2.6, h: rh - 0.16, fontSize: 15.2, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('pk', 0.88, yy + 0.08, 2.6, rh - 0.16, r[0], 14, { pad: 0.02 });
    s.addText(r[1], { x: 3.62, y: yy + 0.08, w: W - 0.62 - 3.62 - 0.26, h: rh - 0.16, fontSize: 13.7, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('pv', 3.62, yy + 0.08, W - 0.62 - 3.62 - 0.26, rh - 0.16, r[1], 12.5, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('แบบที่ห้าคือแบบที่ผู้เรียนนึกไม่ถึงที่สุด แต่มีประโยชน์มาก เพราะแก้ปัญหาเรื่องความเป็นส่วนตัวของข้อมูลไปพร้อมกัน');
}

/* 11 — แบบฝึกหัด 3.1 */
{
  beginSlide('WS3.1');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'แบบฝึกหัด 3.1 · เขียนคำสั่งให้ได้สูตร', 'สูตรที่ใช้ได้จริง ใช้เวลา 20 นาที ทำกับไฟล์ของสาขาตนเอง');
  card(s, 0.62, y, W - 1.24, 2.5, P.soft);
  s.addText('สิ่งที่ต้องทำ', { x: 0.9, y: y + 0.2, w: W - 1.8, h: 0.36, fontSize: 16.2, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('wt', 0.9, y + 0.2, W - 1.8, 0.36);
  bullets(s, 0.9, y + 0.6, W - 1.8, 1.75, [
    'เปิดไฟล์ฝึกปฏิบัติของสาขาตนเอง อ่านชีตอ่านก่อนเริ่มให้จบ',
    'เลือกโจทย์ระดับ Advanced ในชีตนั้น แล้วเขียนคำสั่งตามสูตรสี่ส่วน',
    'บันทึกทั้งคำสั่งที่เขียนและสูตรที่ได้ ลงในชีตชื่อ บันทึกการตัดสินใจ',
    'ทดสอบสูตรกับข้อมูลจริง แล้วบันทึกว่าต้องแก้คำสั่งกี่รอบจึงใช้ได้',
  ], 13);
  const cw = (W - 1.24 - 0.4) / 2;
  textCard(s, 0.62, y + 2.72, cw, 1.35, 'เกณฑ์ผ่าน',
    'สูตรทำงานได้จริงกับไฟล์ และคำสั่งที่เขียนมีครบทั้งสี่ส่วน', P.soft2, '00776A');
  textCard(s, 0.62 + cw + 0.4, y + 2.72, cw, 1.35, 'สิ่งที่ผู้สอนจะดู',
    'จำนวนรอบที่ต้องแก้คำสั่ง ยิ่งน้อยยิ่งแสดงว่าเขียนคำสั่งได้ครบตั้งแต่แรก', P.warn, 'A85A00');
  checkOverlaps();
  s.addNotes('เดินดูว่าใครเขียนคำสั่งแล้วยังขาดส่วนโครงสร้างข้อมูล ซึ่งเป็นส่วนที่ขาดบ่อยที่สุด ให้ชี้ทีละคนแทนการอธิบายรวม');
}

/* 12 — โมดูล 3: ตัวเลขเปิดประเด็น */
{
  beginSlide('ตรวจสอบ intro');
  const s = wrapSlide(darkSlide());
  startSlideDark(s);
  checkOverlaps();
  function startSlideDark(sl) {
    sl.addText('โมดูลที่ 3', { x: 0.9, y: 1.5, w: 6, h: 0.4, fontSize: 15.2, color: P.teal, fontFace: FONT, margin: 0, charSpacing: 2 });
    reg('mlabel', 0.9, 1.5, 6, 0.4);
    sl.addText('การตรวจสอบผลลัพธ์อย่างเป็นระบบ', { x: 0.9, y: 1.95, w: 11.5, h: 0.85, fontSize: 36, bold: true, color: P.white, fontFace: FONT, margin: 0 });
    reg('mtitle', 0.9, 1.95, 11.5, 0.85, 'การตรวจสอบผลลัพธ์อย่างเป็นระบบ', 36, { pad: 0.05 });
    sl.addText('ถ้าจำอะไรจากคาบนี้ได้อย่างเดียว ให้จำโมดูลนี้', { x: 0.9, y: 2.9, w: 11.5, h: 0.5, fontSize: 18.2, color: 'CFC4E8', fontFace: FONT, margin: 0 });
    reg('msub', 0.9, 2.9, 11.5, 0.5, 'ถ้าจำอะไรจากคาบนี้ได้อย่างเดียว ให้จำโมดูลนี้', 17, { pad: 0.05 });
    sl.addText('สูตรที่ผิดส่วนใหญ่ไม่ขึ้น error แต่ให้ตัวเลขที่ผิด การไม่มี error จึงไม่ใช่หลักฐานว่าถูก',
      { x: 0.9, y: 3.75, w: 11.5, h: 0.5, fontSize: 15.2, color: '9C8FC0', fontFace: FONT, margin: 0 });
    reg('mnote', 0.9, 3.75, 11.5, 0.5, 'สูตรที่ผิดส่วนใหญ่ไม่ขึ้น error แต่ให้ตัวเลขที่ผิด การไม่มี error จึงไม่ใช่หลักฐานว่าถูก', 14, { pad: 0.05 });
  }
  s.addNotes('พูดช้าลงตรงนี้ ให้เป็นจุดเปลี่ยนจังหวะของคาบ');
}

/* 13 — วิธีตรวจสอบ */
{
  beginSlide('วิธีตรวจ');
  const s = wrapSlide(lightSlide());
  const y = head(s, 3, 'สี่วิธีตรวจที่ใช้ได้จริง', 'เรียงจากที่ให้หลักฐานหนักแน่นที่สุดลงมา');
  const rows = [
    ['ทดสอบกับกรณีที่รู้คำตอบ', 'คำนวณด้วยมือจากข้อมูลย่อย 5 ถึง 10 แถวก่อน แล้วเทียบกับผลของสูตร วิธีนี้ให้หลักฐานอิสระ จึงน่าเชื่อถือที่สุด', P.teal],
    ['ตรวจผลรวมย้อนกลับ', 'ผลรวมของกลุ่มย่อยทุกกลุ่มต้องเท่ากับผลรวมทั้งหมด ถ้าไม่เท่าแปลว่ามีแถวตกหล่นหรือถูกนับซ้ำ', P.violet],
    ['ตรวจจำนวนก่อนตรวจค่า', 'ดูจำนวนแถวที่เข้าคำนวณก่อนเสมอ ค่าเฉลี่ยที่สวยจากข้อมูลสองแถวไม่มีความหมาย', P.violet],
    ['ตรวจด้วยตรรกะของสาขาตนเอง', 'ค่าฝุ่นติดลบ ความหนาชั้นหินติดลบ pH เท่ากับ 15 เป็นไปไม่ได้ในทางกายภาพ ความรู้ในสาขาคือเครื่องมือตรวจที่ดีที่สุด', P.amber],
  ];
  const rh = 0.95, gap = 0.14;
  rows.forEach((r, i) => {
    const yy = y + 0.05 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 3 ? P.warn : P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: 0.86, y: yy + 0.24, w: 0.46, h: 0.46, fill: { color: r[2] }, line: { width: 0 } });
    s.addText(String(i + 1), { x: 0.86, y: yy + 0.24, w: 0.46, h: 0.46, fontSize: 16.2, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('vn', 0.86, yy + 0.24, 0.46, 0.46);
    s.addText(r[0], { x: 1.48, y: yy + 0.12, w: 3.3, h: rh - 0.24, fontSize: 14.7, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('vk', 1.48, yy + 0.12, 3.3, rh - 0.24, r[0], 13.5, { pad: 0.02 });
    s.addText(r[1], { x: 4.92, y: yy + 0.12, w: W - 0.62 - 4.92 - 0.26, h: rh - 0.24, fontSize: 13.2, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('vv', 4.92, yy + 0.12, W - 0.62 - 4.92 - 0.26, rh - 0.24, r[1], 12, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('ย้ำว่าการถาม AI ซ้ำว่าสูตรถูกไหม ไม่ใช่การตรวจสอบ เพราะแหล่งที่มาของคำตอบยังเป็นแหล่งเดิม');
}

/* ภาพการนับซ้ำ */
{
  const s = mockPair('ภาพนับซ้ำ', 3,
    'ตัวเลขที่ถูกโดยบังเอิญ',
    'ทั้งสองฝั่งได้คำตอบเป็นสามเท่ากัน แต่ฝั่งซ้ายผิด นี่คือกับดักที่ร้ายที่สุด',
    'นับจากข้อมูลที่ยังมีแถวซ้ำ', 'นับเฉพาะค่าที่ไม่ซ้ำ',
    { cols: ['รหัสตัวอย่าง', 'สถานะ'],
      rows: [['CB-014', ''], [['CB-014', 'x'], ['แถวซ้ำ', 'x']], ['CB-021', ''],
             [['COUNTA ได้ 3 ตัวอย่าง', 'x'], ['แต่มีจริง 2', 'x']]] },
    { cols: ['รหัสตัวอย่าง', 'สถานะ'],
      rows: [['CB-014', ''], ['CB-021', ''], ['CB-033', ''],
             ['นับค่าที่ไม่ซ้ำได้ 3 ตัวอย่าง', 'ถูกต้อง']] },
    'ฝั่งซ้ายนับตัวอย่างเดียวกันสองครั้งและตกอีกตัวอย่างหนึ่งไป แต่บังเอิญได้เลขเท่ากัน เมื่อไม่มีอะไรกระตุ้นให้สงสัย ความผิดพลาดจึงอยู่ในรายงานได้จนถึงมือผู้อ่าน');
  s.addNotes('เชื่อมกับแบบฝึกหัด 3.2 ที่จะทำต่อ บอกว่าอีกสี่กับดักที่เหลืออยู่ในนั้น');
}

/* 14 — WS3.2 Bug Hunt */
{
  beginSlide('WS3.2');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'แบบฝึกหัด 3.2 · ล่าสูตรผิด', 'แบบฝึกหัดที่สำคัญที่สุดของคาบนี้ ใช้เวลา 25 นาที ทำเป็นคู่');
  card(s, 0.62, y, W - 1.24, 1.55, P.warn);
  s.addText('โจทย์', { x: 0.9, y: y + 0.16, w: W - 1.8, h: 0.34, fontSize: 15.2, bold: true, color: 'A85A00', fontFace: FONT, margin: 0 });
  reg('bh1', 0.9, y + 0.16, W - 1.8, 0.34);
  s.addText('รับไฟล์ที่มีสูตรผิดซ่อนอยู่ห้าจุด ทุกจุดให้ผลลัพธ์ที่ดูสมเหตุสมผลและไม่มี error หาให้ครบแล้วอธิบายว่าผิดอย่างไร ใช้ AI ช่วยได้เต็มที่ แต่ต้องพิสูจน์ทุกข้อด้วยตนเอง', {
    x: 0.9, y: y + 0.54, w: W - 1.8, h: 0.85, fontSize: 14.2, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('bh2', 0.9, y + 0.54, W - 1.8, 0.85, 'รับไฟล์ที่มีสูตรผิดซ่อนอยู่ห้าจุด ทุกจุดให้ผลลัพธ์ที่ดูสมเหตุสมผลและไม่มี error หาให้ครบแล้วอธิบายว่าผิดอย่างไร ใช้ AI ช่วยได้เต็มที่ แต่ต้องพิสูจน์ทุกข้อด้วยตนเอง', 13, { pad: 0.02 });

  s.addText('ประเภทของจุดผิดที่ซ่อนไว้', { x: 0.62, y: y + 1.72, w: W - 1.24, h: 0.36, fontSize: 16.2, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('bh3', 0.62, y + 1.72, W - 1.24, 0.36);
  const traps = [
    'ช่วงอ้างอิงคลาดไปหนึ่งแถว',
    'AVERAGE กับช่วงที่มีเซลล์ว่าง',
    'VLOOKUP ที่ลืมใส่ FALSE',
    'นับซ้ำจากแถวที่ซ้ำกัน',
    'หน่วยวัดที่ไม่ตรงกัน',
  ];
  const tw = (W - 1.24 - 0.4) / 5;
  traps.forEach((t, i) => {
    const x = 0.62 + i * (tw + 0.1);
    card(s, x, y + 2.15, tw, 1.15, P.soft);
    s.addText(t, { x: x + 0.14, y: y + 2.25, w: tw - 0.28, h: 0.95, fontSize: 12.7, color: P.ink70, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('trap', x + 0.14, y + 2.25, tw - 0.28, 0.95, t, 11.5, { pad: 0.02 });
  });
  card(s, 0.62, y + 3.45, W - 1.24, 0.72, P.soft2);
  s.addText('เกณฑ์ผ่าน: พบอย่างน้อยสี่จากห้าจุด และอธิบายสาเหตุได้ถูกต้อง', {
    x: 0.9, y: y + 3.58, w: W - 1.8, h: 0.48, fontSize: 15.2, bold: true, color: '00776A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('bh4', 0.9, y + 3.58, W - 1.8, 0.48);
  checkOverlaps();
  s.addNotes('แบบฝึกหัดนี้ฝึกความไม่เชื่อโดยอัตโนมัติ ซึ่งเป็นสิ่งที่แยกผู้ใช้ AI อย่างมืออาชีพออกจากผู้ใช้ทั่วไป อย่าเฉลยเร็ว ปล่อยให้อึดอัดสักพัก');
}

/* 15 — กับดักจริงจาก 7 สาขา */
{
  beginSlide('กับดัก 7 สาขา');
  const s = wrapSlide(lightSlide());
  const y = head(s, 3, 'กับดักจริงจากไฟล์ของแต่ละสาขา', 'ทุกอันมาจากงานจริง ไม่มีอันไหนที่สร้างมาเพื่อแกล้ง');
  const rows = [
    ['AC', 'ใบสำคัญไม่ดุลหกใบ แต่ถ้าตรวจก่อนลบแถวซ้ำจะเจอสิบสามใบ ลำดับการทำงานเปลี่ยนคำตอบ'],
    ['ED', 'ค่า -999 ทำให้ค่าเฉลี่ยฝุ่นติดลบ และมีวันหายจากอนุกรมเวลาสี่วันโดยไม่มีอะไรเตือน'],
    ['FT', 'ผลตรวจจุลินทรีย์เขียนว่า น้อยกว่าสิบ หรือ ND ซึ่งถูกต้องตามธรรมเนียมแล็บ แต่คำนวณตรง ๆ ไม่ได้'],
    ['GS', 'สิบสามแถวที่ความลึกช่วงบนมากกว่าช่วงล่าง ผิดตรรกะกายภาพแต่โปรแกรมไม่เตือน'],
    ['BA', 'จำนวนสินค้าติดลบคือรายการคืนสินค้า ไม่ใช่ข้อผิดพลาด ลบทิ้งแล้วยอดขายจะผิด'],
    ['CB', 'แถวซ้ำสิบสองแถวทำให้ความถี่การพบสัตว์สูงเกินจริง กระทบข้อสรุปเชิงอนุรักษ์โดยตรง'],
    ['AG', 'ค่า -999 ปนในผลผลิต ดึงค่าเฉลี่ยลงโดยไม่มีสัญญาณเตือนใด ๆ'],
  ];
  const rh = 0.52, gap = 0.075;
  rows.forEach((r, i) => {
    const yy = y + 0.02 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i < 4 ? P.soft : P.soft2);
    s.addShape(pres.ShapeType.roundRect, { x: 0.84, y: yy + 0.09, w: 0.62, h: 0.34, rectRadius: 0.05, fill: { color: i < 4 ? P.deep : P.teal }, line: { width: 0 } });
    s.addText(r[0], { x: 0.84, y: yy + 0.09, w: 0.62, h: 0.34, fontSize: 12.7, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('code', 0.84, yy + 0.09, 0.62, 0.34);
    s.addText(r[1], { x: 1.6, y: yy + 0.04, w: W - 0.62 - 1.6 - 0.26, h: rh - 0.08, fontSize: 12.7, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('desc', 1.6, yy + 0.04, W - 0.62 - 1.6 - 0.26, rh - 0.08, r[1], 11.5, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('ให้แต่ละสาขาเล่าให้เพื่อนสาขาอื่นฟังว่ากับดักของตนคืออะไรและเจอในงานจริงอย่างไร ใช้เวลาไม่เกินห้านาที');
}

/* 16 — เบรก */
{
  beginSlide('เบรก');
  const s = wrapSlide(darkSlide());
  s.addShape(pres.ShapeType.ellipse, { x: -1.2, y: 4.4, w: 4.4, h: 4.4, fill: { color: P.deep }, line: { width: 0 } });
  s.addText('พักเบรก 10 นาที', { x: 0.9, y: 2.6, w: 11.5, h: 0.9, fontSize: 38, bold: true, color: P.white, fontFace: FONT, margin: 0, align: 'center' });
  reg('brk', 0.9, 2.6, 11.5, 0.9);
  s.addText('กลับมาแล้วเราจะเปลี่ยนจากการตรวจสอบ ไปสู่การสร้างระบบที่ทำงานเองได้', {
    x: 0.9, y: 3.6, w: 11.5, h: 0.5, fontSize: 17.2, color: 'CFC4E8', fontFace: FONT, margin: 0, align: 'center'
  });
  reg('brk2', 0.9, 3.6, 11.5, 0.5, 'กลับมาแล้วเราจะเปลี่ยนจากการตรวจสอบ ไปสู่การสร้างระบบที่ทำงานเองได้', 16, { pad: 0.05 });
  checkOverlaps();
  s.addNotes('ใช้ช่วงเบรกเก็บคำถามค้างจากแบบฝึกหัดล่าสูตรผิด และดูว่ามีคู่ไหนยังหาไม่ครบ');
}

/* 17 — โมดูล 4: Apps Script */
{
  beginSlide('Apps Script');
  const s = wrapSlide(lightSlide());
  const y = head(s, 4, 'ระบบอัตโนมัติเริ่มจากคำถามเดียว', 'งานไหนที่คุณทำซ้ำทุกสัปดาห์ด้วยขั้นตอนเดิมทุกครั้ง');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.4, P.soft);
  s.addText('Google Apps Script', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 18.2, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('as1', 0.86, y + 0.2, cw - 0.48, 0.4);
  s.addText('ภาษาเดียวกับ JavaScript รันบนเซิร์ฟเวอร์ของ Google เชื่อม Form ชีต Drive และ Gmail เข้าด้วยกันได้', {
    x: 0.86, y: y + 0.64, w: cw - 0.48, h: 0.7, fontSize: 13.7, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('as2', 0.86, y + 0.64, cw - 0.48, 0.7, 'ภาษาเดียวกับ JavaScript รันบนเซิร์ฟเวอร์ของ Google เชื่อม Form ชีต Drive และ Gmail เข้าด้วยกันได้', 12.5, { pad: 0.02 });
  s.addText('เหมาะกับ', { x: 0.86, y: y + 1.42, w: cw - 0.48, h: 0.3, fontSize: 13.7, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
  reg('as3', 0.86, y + 1.42, cw - 0.48, 0.3);
  bullets(s, 0.86, y + 1.76, cw - 0.48, 1.45, [
    'ส่งอีเมลอัตโนมัติเมื่อมีคนตอบฟอร์ม',
    'สรุปข้อมูลและเขียนผลลงชีตทุกคืน',
    'แจ้งเตือนเมื่อค่าเกินเกณฑ์ที่ตั้งไว้',
  ], 12.5);

  card(s, 0.62 + cw + 0.4, y, cw, 3.4, P.soft2);
  s.addText('Power Query', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 18.2, bold: true, color: '00776A', fontFace: FONT, margin: 0 });
  reg('pq1', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  s.addText('เครื่องมือใน Excel ที่บันทึกขั้นตอนการล้างข้อมูลไว้ แล้วเล่นซ้ำได้ทุกครั้งที่ข้อมูลใหม่เข้ามา', {
    x: 0.86 + cw + 0.4, y: y + 0.64, w: cw - 0.48, h: 0.7, fontSize: 13.7, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('pq2', 0.86 + cw + 0.4, y + 0.64, cw - 0.48, 0.7, 'เครื่องมือใน Excel ที่บันทึกขั้นตอนการล้างข้อมูลไว้ แล้วเล่นซ้ำได้ทุกครั้งที่ข้อมูลใหม่เข้ามา', 12.5, { pad: 0.02 });
  s.addText('เหมาะกับ', { x: 0.86 + cw + 0.4, y: y + 1.42, w: cw - 0.48, h: 0.3, fontSize: 13.7, bold: true, color: P.muted, fontFace: FONT, margin: 0 });
  reg('pq3', 0.86 + cw + 0.4, y + 1.42, cw - 0.48, 0.3);
  bullets(s, 0.86 + cw + 0.4, y + 1.76, cw - 0.48, 1.45, [
    'รวมไฟล์หลายไฟล์ที่มีโครงสร้างเหมือนกัน',
    'ล้างข้อมูลชุดเดิมซ้ำ ๆ ทุกเดือน',
    'แปลงข้อมูลจากระบบอื่นให้พร้อมใช้',
  ], 12.5);
  checkOverlaps();
  s.addNotes('เน้นว่าไม่ต้องเขียนโค้ดเป็นก็เริ่มได้ ให้ AI ร่างให้ แต่ต้องอ่านออกว่าแต่ละบรรทัดทำอะไร ซึ่งเป็นสิ่งที่จะฝึกในแบบฝึกหัดถัดไป');
}

/* 18 — ท่อข้อมูล */
{
  beginSlide('ท่อข้อมูล');
  const s = wrapSlide(lightSlide());
  const y = head(s, 4, 'ท่อข้อมูลที่จะสร้างในแบบฝึกหัด', 'สี่ขั้น เชื่อมกันแล้วทำงานเองทั้งหมด');
  const steps = [
    ['Google Form', 'เก็บข้อมูลจากหน้างาน', P.violet],
    ['Google Sheets', 'ข้อมูลไหลเข้าชีตอัตโนมัติ', P.deep],
    ['Apps Script', 'คำนวณ ตรวจเงื่อนไข ส่งแจ้งเตือน', P.teal],
    ['Dashboard', 'ผู้บริหารเปิดดูเองได้ตลอดเวลา', P.amber],
  ];
  const bw = 2.7, gapx = (W - 1.24 - bw * 4) / 3;
  steps.forEach((st, i) => {
    const x = 0.62 + i * (bw + gapx);
    card(s, x, y + 0.35, bw, 2.0, P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: x + (bw - 0.56) / 2, y: y + 0.55, w: 0.56, h: 0.56, fill: { color: st[2] }, line: { width: 0 } });
    s.addText(String(i + 1), { x: x + (bw - 0.56) / 2, y: y + 0.55, w: 0.56, h: 0.56, fontSize: 19.2, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sn', x + (bw - 0.56) / 2, y + 0.55, 0.56, 0.56);
    s.addText(st[0], { x: x + 0.14, y: y + 1.22, w: bw - 0.28, h: 0.38, fontSize: 15.2, bold: true, color: P.deep, fontFace: FONT, align: 'center', margin: 0 });
    reg('sname', x + 0.14, y + 1.22, bw - 0.28, 0.38, st[0], 14, { pad: 0.02 });
    s.addText(st[1], { x: x + 0.16, y: y + 1.62, w: bw - 0.32, h: 0.62, fontSize: 12.7, color: P.ink70, fontFace: FONT, align: 'center', margin: 0, valign: 'top' });
    reg('sdesc', x + 0.16, y + 1.62, bw - 0.32, 0.62, st[1], 11.5, { pad: 0.02 });
    if (i < 3) {
      s.addText('→', { x: x + bw + 0.02, y: y + 1.1, w: gapx - 0.04, h: 0.5, fontSize: 22, color: P.line, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    }
  });
  card(s, 0.62, y + 2.65, W - 1.24, 1.35, P.warn);
  s.addText('จุดที่ต้องระวังที่สุดของระบบอัตโนมัติ', { x: 0.9, y: y + 2.8, w: W - 1.8, h: 0.34, fontSize: 15.2, bold: true, color: 'A85A00', fontFace: FONT, margin: 0 });
  reg('wn1', 0.9, y + 2.8, W - 1.8, 0.34);
  s.addText('ระบบที่ทำงานเองจะทำผิดเองด้วย และทำผิดซ้ำทุกวันโดยไม่มีใครเห็น ก่อนเปิดใช้จริงต้องทดสอบด้วยข้อมูลที่รู้คำตอบ และต้องมีวิธีรู้ว่าระบบหยุดทำงานเมื่อไร', {
    x: 0.9, y: y + 3.18, w: W - 1.8, h: 0.7, fontSize: 13.7, color: P.ink70, fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('wn2', 0.9, y + 3.18, W - 1.8, 0.7, 'ระบบที่ทำงานเองจะทำผิดเองด้วย และทำผิดซ้ำทุกวันโดยไม่มีใครเห็น ก่อนเปิดใช้จริงต้องทดสอบด้วยข้อมูลที่รู้คำตอบ และต้องมีวิธีรู้ว่าระบบหยุดทำงานเมื่อไร', 12.5, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('เล่าตัวอย่างจริงว่าระบบส่งอีเมลอัตโนมัติที่ตั้งผิดเงื่อนไข ส่งอีเมลผิดไปหลายร้อยฉบับก่อนมีคนทัก');
}

/* 19 — WS3.3 และ 3.4 */
{
  beginSlide('WS3.3-3.4');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'แบบฝึกหัดช่วงท้าย', 'สองชิ้น ทำต่อเนื่องกัน 60 นาที');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 3.7, P.soft);
  s.addText('3.3 · ระบบส่งอีเมลอัตโนมัติ', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17.2, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w31', 0.86, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86, y + 0.68, cw - 0.48, 1.9, [
    'ใช้ AI ช่วยเขียน Apps Script ที่ส่งอีเมลเมื่อมีผู้ตอบฟอร์ม',
    'ทดสอบจริงด้วยการส่งฟอร์มด้วยตนเอง',
    'อธิบายให้ได้ว่าแต่ละบรรทัดทำอะไร',
  ], 12.5);
  card(s, 0.86, y + 2.7, cw - 0.48, 0.82, P.soft2);
  s.addText('เกณฑ์ผ่าน: สคริปต์รันได้ ส่งอีเมลถึงจริง และอธิบายโค้ดได้', {
    x: 1.05, y: y + 2.82, w: cw - 0.86, h: 0.58, fontSize: 13.2, bold: true, color: '00776A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('w32', 1.05, y + 2.82, cw - 0.86, 0.58, 'เกณฑ์ผ่าน: สคริปต์รันได้ ส่งอีเมลถึงจริง และอธิบายโค้ดได้', 12, { pad: 0.02 });

  card(s, 0.62 + cw + 0.4, y, cw, 3.7, P.soft);
  s.addText('3.4 · Dashboard โครงงานของฉัน', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.4, fontSize: 17.2, bold: true, color: P.deep, fontFace: FONT, margin: 0 });
  reg('w41', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.4);
  bullets(s, 0.86 + cw + 0.4, y + 0.68, cw - 0.48, 1.9, [
    'สร้าง Dashboard หนึ่งหน้าจากข้อมูลของตนเอง',
    'ต้องมีตัวเลขสำคัญสามตัว กราฟสองชิ้น',
    'เขียนข้อสรุปสามบรรทัดที่มีหลักฐานรองรับ',
  ], 12.5);
  card(s, 0.86 + cw + 0.4, y + 2.7, cw - 0.48, 0.82, P.soft2);
  s.addText('เกณฑ์ผ่าน: ตรวจด้วย rubric เน้นความเหมาะกับผู้รับสาร', {
    x: 1.05 + cw + 0.4, y: y + 2.82, w: cw - 0.86, h: 0.58, fontSize: 13.2, bold: true, color: '00776A', fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('w42', 1.05 + cw + 0.4, y + 2.82, cw - 0.86, 0.58, 'เกณฑ์ผ่าน: ตรวจด้วย rubric เน้นความเหมาะกับผู้รับสาร', 12, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ถ้าเวลาไม่พอ ให้ตัด 3.3 เหลือแค่รันสคริปต์ตัวอย่างที่เตรียมไว้ แล้วให้เวลากับ 3.4 เต็มที่ เพราะ 3.4 เป็นชิ้นงานที่ใช้ประเมิน');
}

/* 20 — จริยธรรม */
{
  beginSlide('จริยธรรม');
  const s = wrapSlide(lightSlide());
  const y = head(s, 5, 'ข้อมูลอะไรที่ห้ามวางลง AI', 'คำถามนี้ต้องตอบได้ก่อนกดวางทุกครั้ง');
  const cw = (W - 1.24 - 0.4) / 2;
  card(s, 0.62, y, cw, 2.75, P.warn);
  s.addText('ห้ามวาง', { x: 0.86, y: y + 0.2, w: cw - 0.48, h: 0.36, fontSize: 17.2, bold: true, color: 'A85A00', fontFace: FONT, margin: 0 });
  reg('e1', 0.86, y + 0.2, cw - 0.48, 0.36);
  bullets(s, 0.86, y + 0.64, cw - 0.48, 1.95, [
    'ชื่อ นามสกุล เบอร์โทร เลขบัตรประชาชน',
    'ข้อมูลสุขภาพหรือข้อมูลผู้ป่วย',
    'ข้อมูลวิจัยที่ยังไม่เผยแพร่',
    'ข้อมูลที่มีข้อตกลงรักษาความลับกำกับ',
  ], 12.5);

  card(s, 0.62 + cw + 0.4, y, cw, 2.75, P.soft2);
  s.addText('ทำแทนได้อย่างไร', { x: 0.86 + cw + 0.4, y: y + 0.2, w: cw - 0.48, h: 0.36, fontSize: 17.2, bold: true, color: '00776A', fontFace: FONT, margin: 0 });
  reg('e2', 0.86 + cw + 0.4, y + 0.2, cw - 0.48, 0.36);
  bullets(s, 0.86 + cw + 0.4, y + 0.64, cw - 0.48, 1.95, [
    'สร้างข้อมูลจำลองที่มีโครงสร้างคอลัมน์เหมือนกัน',
    'ส่งเฉพาะหัวตารางและชนิดข้อมูล ไม่ส่งเนื้อข้อมูล',
    'อธิบายโครงสร้างเป็นข้อความแทนการวางตาราง',
  ], 12.5);

  card(s, 0.62, y + 3.0, W - 1.24, 1.0, P.soft);
  s.addText('AI ต้องการเพียงโครงสร้างข้อมูล ไม่ใช่เนื้อข้อมูลจริง การส่งข้อมูลจำลองจึงได้ผลลัพธ์เท่าเดิมโดยไม่มีความเสี่ยงเลย', {
    x: 0.9, y: y + 3.18, w: W - 1.8, h: 0.66, fontSize: 15.2, bold: true, color: P.deep, fontFace: FONT, margin: 0, valign: 'middle'
  });
  reg('e3', 0.9, y + 3.18, W - 1.8, 0.66, 'AI ต้องการเพียงโครงสร้างข้อมูล ไม่ใช่เนื้อข้อมูลจริง การส่งข้อมูลจำลองจึงได้ผลลัพธ์เท่าเดิมโดยไม่มีความเสี่ยงเลย', 14, { pad: 0.02 });
  checkOverlaps();
  s.addNotes('ย้ำว่าการลบเฉพาะคอลัมน์ชื่อออกไม่พอ เพราะเบอร์โทรหรือรหัสนักศึกษาก็ระบุตัวบุคคลได้ และจำนวนแถวที่น้อยลงไม่ได้ทำให้ข้อมูลหมดสภาพการเป็นข้อมูลส่วนบุคคล');
}

/* 21 — การอ้างอิงการใช้ AI */
{
  beginSlide('อ้างอิง AI');
  const s = wrapSlide(lightSlide());
  const y = head(s, 5, 'ใช้ AI แล้วต้องเขียนบอกในรายงานไหม', 'คำตอบสั้น ๆ คือควรบอก และบอกให้ชัดว่าใช้ทำอะไร');
  const rows = [
    ['ใช้ AI ช่วยร่างสูตรหรือโค้ด', 'ควรระบุในภาคผนวกหรือกิตติกรรมประกาศว่าใช้เครื่องมือใดช่วยงานส่วนใด', P.teal],
    ['ใช้ AI ช่วยตรวจไวยากรณ์ภาษา', 'โดยทั่วไปไม่ต้องระบุ เทียบได้กับการใช้เครื่องมือตรวจคำสะกด', P.violet],
    ['ใช้ AI สร้างเนื้อหาเชิงวิชาการ', 'ต้องระบุเสมอ และต้องตรวจสอบข้อเท็จจริงทุกจุดด้วยตนเองก่อน', P.amber],
    ['ให้ AI ตีความผลการวิเคราะห์แทน', 'ไม่ควรทำ การตีความคือหน้าที่ของผู้วิจัย และเป็นส่วนที่ AI ผิดพลาดบ่อยที่สุด', P.coral],
  ];
  const rh = 0.85, gap = 0.16;
  rows.forEach((r, i) => {
    const yy = y + 0.1 + i * (rh + gap);
    card(s, 0.62, yy, W - 1.24, rh, i === 3 ? P.warn : P.soft);
    s.addShape(pres.ShapeType.roundRect, { x: 0.86, y: yy + 0.2, w: 0.16, h: 0.45, rectRadius: 0.05, fill: { color: r[2] }, line: { width: 0 } });
    reg('dot', 0.86, yy + 0.2, 0.16, 0.45);
    s.addText(r[0], { x: 1.2, y: yy + 0.1, w: 3.9, h: rh - 0.2, fontSize: 14.2, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('ak', 1.2, yy + 0.1, 3.9, rh - 0.2, r[0], 13, { pad: 0.02 });
    s.addText(r[1], { x: 5.24, y: yy + 0.1, w: W - 0.62 - 5.24 - 0.26, h: rh - 0.2, fontSize: 13.2, color: P.ink70, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('av', 5.24, yy + 0.1, W - 0.62 - 5.24 - 0.26, rh - 0.2, r[1], 12, { pad: 0.02 });
  });
  s.addText('แนวปฏิบัติของแต่ละวารสารและแต่ละสถาบันต่างกัน ให้ตรวจข้อกำหนดของที่ที่จะส่งงานเสมอ', {
    x: 0.62, y: y + 0.1 + 4 * (rh + gap) + 0.06, w: W - 1.24, h: 0.4, fontSize: 13.2, italic: true, color: P.muted, fontFace: FONT, margin: 0
  });
  reg('anote', 0.62, y + 0.1 + 4 * (rh + gap) + 0.06, W - 1.24, 0.4);
  checkOverlaps();
  s.addNotes('บอกผู้เรียนว่าแนวปฏิบัติเรื่องนี้ยังเปลี่ยนอยู่ตลอด ให้ตรวจข้อกำหนดล่าสุดของวารสารหรือของบัณฑิตวิทยาลัยก่อนส่งงานทุกครั้ง');
}

/* 22 — สรุป */
{
  beginSlide('สรุป');
  const s = wrapSlide(lightSlide());
  const y = head(s, null, 'สามอย่างที่อยากให้จำ');
  const items = [
    ['1', 'ให้ AI เขียนสูตร ให้สเปรดชีตคิดเลข', 'AI เก่งเรื่องภาษาและตรรกะ ไม่ได้ออกแบบมาเพื่อคำนวณ การแบ่งงานให้ถูกจึงสำคัญกว่าการเลือกเครื่องมือ', P.violet],
    ['2', 'ไม่มี error ไม่ได้แปลว่าถูก', 'สูตรที่ผิดส่วนใหญ่ให้ตัวเลขที่ดูดี ทดสอบกับกรณีที่รู้คำตอบก่อนเสมอ', P.amber],
    ['3', 'ความรู้ในสาขาคือเครื่องมือตรวจที่ดีที่สุด', 'ค่าฝุ่นติดลบหรือความหนาชั้นหินติดลบ คุณเห็นทันทีว่าผิด แต่ AI ไม่เห็น', P.teal],
  ];
  const ch = 1.28, gap = 0.2;
  items.forEach((it, i) => {
    const yy = y + 0.15 + i * (ch + gap);
    card(s, 0.62, yy, W - 1.24, ch, P.soft);
    s.addShape(pres.ShapeType.ellipse, { x: 0.9, y: yy + 0.34, w: 0.6, h: 0.6, fill: { color: it[3] }, line: { width: 0 } });
    s.addText(it[0], { x: 0.9, y: yy + 0.34, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: P.white, fontFace: FONT, align: 'center', valign: 'middle', margin: 0 });
    reg('sn2', 0.9, yy + 0.34, 0.6, 0.6);
    s.addText(it[1], { x: 1.72, y: yy + 0.18, w: W - 0.62 - 1.72 - 0.3, h: 0.42, fontSize: 17.2, bold: true, color: P.deep, fontFace: FONT, valign: 'middle', margin: 0 });
    reg('sk', 1.72, yy + 0.18, W - 0.62 - 1.72 - 0.3, 0.42, it[1], 16, { pad: 0.02 });
    s.addText(it[2], { x: 1.72, y: yy + 0.62, w: W - 0.62 - 1.72 - 0.3, h: 0.56, fontSize: 13.7, color: P.ink70, fontFace: FONT, valign: 'top', margin: 0 });
    reg('sv', 1.72, yy + 0.62, W - 0.62 - 1.72 - 0.3, 0.56, it[2], 12.5, { pad: 0.02 });
  });
  checkOverlaps();
  s.addNotes('ให้ผู้เรียนเขียนสามข้อนี้ลงในชีตบันทึกการตัดสินใจของตัวเอง แล้วจึงทำแบบทดสอบหลังเรียน');
}

/* 23 — ปิด */
{
  beginSlide('ปิด');
  const s = wrapSlide(darkSlide());
  s.addShape(pres.ShapeType.ellipse, { x: 10.8, y: 4.2, w: 3.6, h: 3.6, fill: { color: P.deep }, line: { width: 0 } });
  s.addText('แบบทดสอบหลังเรียน', { x: 0.9, y: 2.15, w: 9.5, h: 0.8, fontSize: 34, bold: true, color: P.white, fontFace: FONT, margin: 0 });
  reg('end1', 0.9, 2.15, 9.5, 0.8);
  s.addText('ใช้เวลา 5 นาที ทำโดยไม่เปิดเอกสาร เพื่อให้เห็นพัฒนาการของตัวเองได้จริง', {
    x: 0.9, y: 3.05, w: 9.5, h: 0.5, fontSize: 17.2, color: 'CFC4E8', fontFace: FONT, margin: 0
  });
  reg('end2', 0.9, 3.05, 9.5, 0.5, 'ใช้เวลา 5 นาที ทำโดยไม่เปิดเอกสาร เพื่อให้เห็นพัฒนาการของตัวเองได้จริง', 16, { pad: 0.05 });
  s.addText('ส่งแบบฝึกหัดทั้งหมดผ่าน Google Classroom ภายในหนึ่งสัปดาห์  ·  ผลคะแนนก่อนและหลังเรียนจะแจ้งให้ทราบพร้อมข้อเสนอแนะรายบุคคล', {
    x: 0.9, y: 4.35, w: 9.8, h: 0.7, fontSize: 14.2, color: '9C8FC0', fontFace: FONT, margin: 0, valign: 'top'
  });
  reg('end3', 0.9, 4.35, 9.8, 0.7, 'ส่งแบบฝึกหัดทั้งหมดผ่าน Google Classroom ภายในหนึ่งสัปดาห์  ·  ผลคะแนนก่อนและหลังเรียนจะแจ้งให้ทราบพร้อมข้อเสนอแนะรายบุคคล', 13, { pad: 0.05 });
  checkOverlaps();
  s.addNotes('ปิดด้วยการเชิญชวนให้เอาสิ่งที่เรียนไปใช้กับงานจริงภายในสัปดาห์นี้ เพราะทักษะที่ไม่ได้ใช้ภายในเจ็ดวันมักหายไป');
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
      const gap = e.unused;                       // พื้นที่ว่างที่เหลือหลังเลื่อนรอบนี้แล้ว
      const add = Math.max(0, Math.min(1.1, gap * 0.5));
      off[String(e.no)] = Math.round((prev + add) * 100) / 100;
    });
    fs.writeFileSync(OFFSET_FILE, JSON.stringify(off, null, 2));
    console.log('บันทึกระยะเลื่อนลงไฟล์ deck_offsets.json แล้ว');
  }
  if (n > 0) process.exitCode = 1;
});
