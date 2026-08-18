/**
 * เครื่องมือช่วยจัดวางสไลด์ + ตัวตรวจ layout เชิงตัวเลข
 * จำเป็นเพราะสภาพแวดล้อมที่ใช้ตรวจงานไม่มีฟอนต์ไทย ข้อความไทยจะ render เป็นกล่องสี่เหลี่ยม
 * การตรวจด้วยตาจึงบอกไม่ได้ว่าข้อความล้นกรอบหรือไม่ ต้องคำนวณเอาเอง
 */
const W = 13.333, H = 7.5, MARGIN = 0.5;

const P = {
  ink:    '231536',   // พื้นเข้ม
  deep:   '3B1E6E',   // ม่วงเข้ม สีหลัก
  violet: '6D45B8',
  teal:   '00A896',
  amber:  'F2A03D',
  coral:  'E05C5C',
  white:  'FFFFFF',
  soft:   'F4F1FA',   // พื้นการ์ดอ่อน
  soft2:  'EAF7F4',
  warn:   'FDF0E3',
  ink70:  '4A4160',
  muted:  '6B6480',
  line:   'DCD5EA'
};

const FONT = 'Tahoma';   // ฟอนต์ที่มีสระและวรรณยุกต์ไทยครบทั้งบน Windows และ Mac Office

// ---------------------------------------------------------------- ตัวตรวจ
const issues = [];
let boxes = [];      // กล่องบนสไลด์ปัจจุบัน
let slideNo = 0;

const extents = [];
function startSlide(name) {
  if (boxes.length) recordExtent();
  slideNo++;
  boxes = [];
  curName = name;
  return { name: name, no: slideNo };
}
let curName = '';
function recordExtent() {
  let bot = 0;
  boxes.forEach(b => { bot = Math.max(bot, b.y + b.h); });
  extents.push({ no: slideNo, name: curName, bottom: bot, unused: H - 0.35 - bot });
}

/** ประมาณจำนวนบรรทัดที่ข้อความต้องใช้ในกล่องกว้าง w นิ้ว ที่ขนาดฟอนต์ size pt */
function linesNeeded(text, w, size, pad) {
  pad = pad === undefined ? 0.2 : pad;          // ระยะขอบในกล่องข้อความของ pptx
  const usablePt = (w - pad) * 72;
  // ยิ่งฟอนต์ใหญ่ยิ่งต้องเผื่อมาก เพราะสระและวรรณยุกต์ไทยกินที่แนวนอนมากกว่าที่ประมาณไว้แบบเชิงเส้น
  const avg = (size >= 24 ? 0.72 : 0.58) * size;
  const cpl = Math.max(4, Math.floor(usablePt / avg));
  const paras = String(text).split('\n');
  let n = 0;
  paras.forEach(p => { n += Math.max(1, Math.ceil(p.length / cpl)); });
  return n;
}

function heightNeeded(text, w, size, pad) {
  return linesNeeded(text, w, size, pad) * (size * 1.32) / 72 + 0.1;
}

/** ลงทะเบียนกล่องเพื่อให้ตรวจขอบเขต การซ้อนทับ และการล้น */
function reg(kind, x, y, w, h, text, size, opts) {
  opts = opts || {};
  boxes.push({ kind, x, y, w, h, text: text || '', size: size || 14, soft: !!opts.soft });
  if (text && size) {
    const need = heightNeeded(text, w, size, opts.pad);
    if (need > h + 0.02) {
      issues.push(`สไลด์ ${slideNo}: ข้อความล้นกรอบ (${kind}) ต้องการ ${need.toFixed(2)}" มี ${h.toFixed(2)}" · "${String(text).slice(0, 42)}..."`);
    }
  }
  if (x < MARGIN - 0.001 || y < 0.2 - 0.001 || x + w > W - MARGIN + 0.001 || y + h > H - 0.25 + 0.001) {
    issues.push(`สไลด์ ${slideNo}: กล่องออกนอกขอบเขต (${kind}) x=${x} y=${y} w=${w} h=${h}`);
  }
}

function checkOverlaps() {
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      const a = boxes[i], b = boxes[j];
      if (a.soft || b.soft) continue;                       // การ์ดพื้นหลังอนุญาตให้ซ้อนกับข้อความในตัวมันเอง
      const ox = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
      const oy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
      if (ox > 0.04 && oy > 0.04) {
        issues.push(`สไลด์ ${slideNo}: กล่องซ้อนทับ ${a.kind} กับ ${b.kind} (ซ้อน ${ox.toFixed(2)}"×${oy.toFixed(2)}")`);
      }
    }
  }
}

function report() {
  if (boxes.length) recordExtent();
  const loose = extents.filter(e => e.unused > 1.0);
  if (loose.length) {
    console.log('สไลด์ที่เหลือพื้นที่ว่างด้านล่างเกิน 1 นิ้ว (ควรขยายเนื้อหาให้เต็มหน้า)');
    loose.forEach(e => console.log(`  สไลด์ ${e.no} ${e.name}: เนื้อหาจบที่ ${e.bottom.toFixed(2)}" เหลือว่าง ${e.unused.toFixed(2)}"`));
    console.log('');
  }
  if (issues.length === 0) { console.log('ตรวจ layout: ผ่านทั้งหมด ไม่พบข้อความล้นกรอบหรือกล่องซ้อนทับ'); return 0; }
  console.log(`ตรวจ layout: พบ ${issues.length} จุด`);
  issues.forEach(s => console.log('  - ' + s));
  return issues.length;
}

module.exports = { W, H, MARGIN, P, FONT, startSlide, reg, checkOverlaps, report, heightNeeded, linesNeeded, issues, extents };
