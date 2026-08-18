/**
 * สร้างไฟล์พรีวิวของหน้าเว็บที่ปกติต้อง deploy บน Apps Script ก่อนจึงจะเห็น
 * ใช้สำหรับตรวจงานเท่านั้น เปิดด้วยเบราว์เซอร์ได้ทันที ไม่ต้องมีบัญชี Google
 *
 * วิธีทำงาน: โหลดโค้ด .gs เข้ามารันในสภาพแวดล้อมจำลอง แล้วเรียกฟังก์ชันที่ประกอบ HTML
 * จึงได้หน้าตาตรงกับของจริงทุกประการ ไม่ใช่การทำหน้าจำลองขึ้นใหม่
 */
const fs = require('fs');
const path = require('path');

const OUTDIR = process.argv[2];
// ไฟล์ .gs อยู่คนละโฟลเดอร์ เพราะเป็นของที่ต้องนำไปวางใน Apps Script ไม่ใช่ซอร์สของเว็บ
const GSDIR = process.argv[3] || path.join(__dirname, '..', '07 สคริปต์ Apps Script');
fs.mkdirSync(OUTDIR, { recursive: true });

/* ---------- สภาพแวดล้อมจำลองของ Apps Script เท่าที่โค้ดต้องใช้ ---------- */
function makeSandbox() {
  return {
    Logger: { log: () => {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => null }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    DriveApp: { getFileById: () => ({ addViewer: () => {}, removeViewer: () => {} }) },
    FormApp: {},
    Session: { getActiveUser: () => ({ getEmail: () => '' }) },
    MailApp: { sendEmail: () => {} },
    HtmlService: { createHtmlOutput: (h) => ({ setTitle: () => ({ addMetaTag: () => h }) }) },
  };
}

function loadGs(file, extra) {
  const code = fs.readFileSync(path.join(GSDIR, file), 'utf8');
  const sb = makeSandbox();
  const ctx = Object.assign({}, sb, extra || {});
  const names = Object.keys(ctx);
  const vals = names.map(n => ctx[n]);
  // คืนค่า scope ทั้งหมดกลับมาเพื่อเรียกฟังก์ชันภายในได้
  const fn = new Function(...names, code + '\n;return this;');
  const self = {};
  const runner = new Function(...names, 'body', 'return (function(){' + code + '\nreturn { get: function(n){ return eval(n); } };}).call(this)');
  return runner(...vals);
}

const BANNER = (title, note) =>
  '<div style="background:#FFF4E5;border-bottom:2px solid #E8A33D;padding:14px 18px;'
  + 'font-family:Sarabun,\'Noto Sans Thai\',-apple-system,\'Segoe UI\',sans-serif;font-size:14px;line-height:1.6">'
  + '<b style="color:#8A5A00">พรีวิวสำหรับตรวจงาน · ' + title + '</b><br>'
  + '<span style="color:#6B5320">' + note + '</span></div>';

function wrap(inner, styles, title, note) {
  return '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<title>' + title + '</title>' + styles + '</head><body>'
    + BANNER(title, note)
    + '<div class="wrap">' + inner + '</div></body></html>';
}

/* ============================================================ หน้าผลแบบทดสอบ */
{
  const m = loadGs('3_ResultsWebApp.gs');
  const decide = loadGs('2_ScoreAndNotify.gs');

  // ต้องใช้ decideLevel จากไฟล์ที่สอง เพราะไฟล์ที่สามเรียกใช้ข้ามไฟล์กันในโปรเจกต์เดียว
  const combined = fs.readFileSync(path.join(GSDIR, '2_ScoreAndNotify.gs'), 'utf8')
    + '\n' + fs.readFileSync(path.join(GSDIR, '3_ResultsWebApp.gs'), 'utf8');
  const sb = makeSandbox();
  const names = Object.keys(sb), vals = names.map(n => sb[n]);
  const api = new Function(...names,
    'return (function(){' + combined + '\nreturn { cardResult: cardResult, libStyles: styles, styles: styles, decideLevel: decideLevel, footer: footer };})()')(...vals);

  const cases = [
    {
      file: 'พรีวิว-2-ผลแบบทดสอบ-Advanced.html',
      title: 'หน้าผลแบบทดสอบจัดระดับ (กรณีได้ระดับ Advanced)',
      found: { email: 'student01@student.mahidol.ac.th', name: 'สมชาย ใจดี', confidence: 4 },
      s: { A: 6, B: 5, C: 4, anchorA: 1, anchorB: 1, anchorC: 1, answered: 18, correct: 15,
           reachedB: true, reachedC: true, wrongTopics: ['การตรวจสอบผลลัพธ์', 'คุณภาพของ prompt'], ratio: 15 / 18 },
    },
    {
      file: 'พรีวิว-2-ผลแบบทดสอบ-Beginner.html',
      title: 'หน้าผลแบบทดสอบจัดระดับ (กรณีได้ระดับ Beginner และประเมินตนเองสูงกว่าผล)',
      found: { email: 'student02@student.mahidol.ac.th', name: 'สมหญิง ตั้งใจ', confidence: 5 },
      s: { A: 2, B: 0, C: 0, anchorA: 0, anchorB: 0, anchorC: 0, answered: 6, correct: 2,
           reachedB: false, reachedC: false, wrongTopics: ['การอ้างอิงเซลล์แบบสัมบูรณ์', 'ชนิดข้อมูล'], ratio: 2 / 6 },
    },
  ];
  cases.forEach(c => {
    const r = api.decideLevel(c.s, c.found.confidence);
    const html = wrap(api.cardResult(c.found, c.s, r) + api.footer(), api.styles(), c.title,
      'หน้านี้คือสิ่งที่ผู้ทำแบบทดสอบเห็นหลังกดส่ง ของจริงจะดึงคำตอบจากฟอร์มมาคำนวณสด '
      + 'พรีวิวนี้ใช้ข้อมูลสมมติเพื่อให้เห็นหน้าตาโดยไม่ต้อง deploy');
    fs.writeFileSync(path.join(OUTDIR, c.file), html);
    console.log('เขียน', c.file, '· ระดับที่ได้', r.level, '(' + r.path + ')');
  });
}

/* ============================================================ หน้าคลังไฟล์ */
{
  const code = fs.readFileSync(path.join(GSDIR, '4_LibraryWebApp.gs'), 'utf8');
  const sb = makeSandbox();
  const names = Object.keys(sb), vals = names.map(n => sb[n]);
  const api = new Function(...names,
    'return (function(){' + code + '\nreturn { LIB: LIB, cardLibrary: cardLibrary, cardNotRegistered: cardNotRegistered, cardNeedSignIn: cardNeedSignIn, libStyles: libStyles, DATASETS: DATASETS };})()')(...vals);

  // ตั้ง File ID สมมติ เพื่อให้ปุ่มดาวน์โหลดปรากฏในพรีวิว
  Object.keys(api.LIB.FILES).forEach(c => {
    api.LIB.FILES[c] = { xlsx: 'PREVIEW_' + c, sheet: 'PREVIEWSHEET_' + c };
  });

  const member = { email: 'student01@student.mahidol.ac.th', name: 'สมชาย ใจดี', level: 'Intermediate', prog: 'AG' };
  const pages = [
    ['พรีวิว-3-คลังไฟล์-ผ่านสิทธิ์.html', 'หน้าคลังไฟล์ (กรณีผู้ลงทะเบียนแล้ว)',
      api.cardLibrary(member, { ok: true, cached: true, errors: [] }),
      'หน้านี้คือสิ่งที่ผู้ลงทะเบียนแล้วเห็น โจทย์ระดับที่ตนลงทะเบียนจะถูกไฮไลต์สีเหลือง '
      + 'และชุดข้อมูลของสาขาตนเองมีป้ายกำกับ · ปุ่มดาวน์โหลดในพรีวิวยังกดไม่ได้เพราะยังไม่ได้ใส่ File ID จริง'],
    ['พรีวิว-3-คลังไฟล์-ยังไม่ลงทะเบียน.html', 'หน้าคลังไฟล์ (กรณียังไม่ลงทะเบียน)',
      api.cardNotRegistered('outsider@example.com'),
      'หน้านี้คือสิ่งที่คนที่ยังไม่ได้ลงทะเบียนเห็น สังเกตว่าไม่มีลิงก์ไฟล์ใด ๆ ปรากฏเลย'],
    ['พรีวิว-3-คลังไฟล์-ยังไม่เข้าสู่ระบบ.html', 'หน้าคลังไฟล์ (กรณีเปิดด้วยบัญชีนอกโดเมน)',
      api.cardNeedSignIn(),
      'หน้านี้แสดงเมื่อระบบระบุตัวผู้ใช้ไม่ได้ เช่น เปิดด้วยบัญชีส่วนตัวแทนบัญชีของสถาบัน'],
  ];
  pages.forEach(([file, title, inner, note]) => {
    fs.writeFileSync(path.join(OUTDIR, file), wrap(inner, api.libStyles(), title, note));
    console.log('เขียน', file);
  });
}

/* ============================================================ หน้าสารบัญพรีวิว */
{
  const files = fs.readdirSync(OUTDIR).filter(f => f.startsWith('พรีวิว')).sort();
  const links = files.map(f =>
    '<li><a href="' + encodeURIComponent(f) + '">' + f.replace('.html', '').replace(/^พรีวิว-/, '') + '</a></li>').join('');
  const html = '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1"><title>สารบัญหน้าพรีวิว</title>'
    + '<style>body{font-family:Sarabun,"Noto Sans Thai",-apple-system,"Segoe UI",sans-serif;background:#F5F7FA;'
    + 'color:#1a1a1a;line-height:1.7;margin:0;padding:40px 20px}'
    + '.box{max-width:760px;margin:0 auto;background:#fff;border-radius:12px;padding:32px;'
    + 'box-shadow:0 1px 3px rgba(16,24,40,.08)}'
    + 'h1{font-size:24px;color:#1F3864;margin:0 0 6px}p{color:#6b7684;margin:0 0 22px;font-size:15px}'
    + 'ol{padding-left:22px}li{margin:10px 0;font-size:15.5px}a{color:#2E74B5}'
    + '.note{background:#FFF4E5;border-left:4px solid #E8A33D;padding:14px 18px;border-radius:0 8px 8px 0;'
    + 'font-size:14px;margin-top:26px;color:#6B5320}</style></head><body><div class="box">'
    + '<h1>สารบัญหน้าพรีวิวทั้งหมด</h1>'
    + '<p>เปิดดูได้ทันทีโดยไม่ต้อง deploy และไม่ต้องเข้าสู่ระบบ ใช้สำหรับตรวจงานก่อนเปิดใช้จริง</p>'
    + '<ol><li><a href="../index.html">หน้าเว็บสาธารณะ (index.html)</a> — หน้าหลักที่ทุกคนเข้าได้</li>'
    + links + '</ol>'
    + '<div class="note"><b>ข้อควรทราบ</b><br>หน้าพรีวิวเหล่านี้เป็นภาพนิ่งที่สร้างจากโค้ดจริง '
    + 'จึงมีหน้าตาตรงกับของจริงทุกประการ แต่ปุ่มและลิงก์ยังกดไม่ได้ เพราะยังไม่ได้ตั้งค่าลิงก์และยังไม่ได้ deploy '
    + 'เมื่อตรวจเสร็จแล้วสามารถลบโฟลเดอร์นี้ทิ้งได้</div>'
    + '</div></body></html>';
  fs.writeFileSync(path.join(OUTDIR, 'เริ่มที่นี่.html'), html);
  console.log('เขียน เริ่มที่นี่.html · รวม', files.length + 1, 'หน้า');
}
