const fs = require('fs');
const bank = JSON.parse(fs.readFileSync(__dirname + '/items.json', 'utf8'));

// ---- clean stem for the live form (strip the anchor marker used in the doc) ----
function cleanStem(s) {
  return s.replace(/^★\s*ข้อชี้ขาดของ Block [ABC]\s*—\s*/, '');
}
// deterministic shuffle so the doc's key and the form agree via TEXT, not position
function shuffle(arr, seed) {
  const a = arr.slice();
  let s = seed;
  for (let i = a.length - 1; i > 0; i--) {
    s = (s * 1103515245 + 12345) % 2147483648;
    const j = s % (i + 1);
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}
const q = s => "'" + String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";

/* ================= FILE 1 : BUILD FORM ================= */
let f1 = `/**
 * 1_BuildPlacementForm.gs
 * สร้าง Google Form แบบทดสอบจัดระดับ (Placement Test) ทั้งฉบับโดยอัตโนมัติ
 * พร้อมการแตกสาขา (branching) ตามข้อชี้ขาดของแต่ละบล็อก
 *
 * ชุดการเรียนรู้ Excel & Google Sheets 3 ระดับ
 * ผู้พัฒนา: อาจารย์ ดร.เนติยา การะเกตุ
 *
 * วิธีใช้: รันฟังก์ชัน buildPlacementForm() หนึ่งครั้ง แล้วดู URL ของฟอร์มใน Execution log
 *
 * หมายเหตุสำคัญ 1: สคริปต์นี้ตั้งใจไม่เปิดโหมด Quiz ของ Google Form
 * เพราะโหมด Quiz จะแสดงคะแนนและเฉลยให้ผู้ตอบเห็นทันที ซึ่งทำให้ข้อสอบรั่ว
 * การให้คะแนนทั้งหมดทำในไฟล์ 2_ScoreAndNotify.gs และ 3_ResultsWebApp.gs
 *
 * หมายเหตุสำคัญ 2: หน้าจบทั้งสามหน้าในฟอร์มใช้ข้อความกลาง ๆ เหมือนกัน ไม่บอกระดับ
 * เพราะการแตกสาขาในฟอร์มอาศัยข้อชี้ขาดเพียงข้อเดียว ซึ่งอาจไม่ตรงกับผลที่คำนวณจากคะแนนทุกข้อ
 * หน้าทั้งสามทำหน้าที่เป็นจุดจบเส้นทางเท่านั้น ส่วนการบอกผลเป็นหน้าที่ของหน้าเว็บผลลัพธ์
 */

var FORM_TITLE = 'แบบทดสอบจัดระดับ — ชุดการเรียนรู้ Excel & Google Sheets';

// URL ของหน้าเว็บผลลัพธ์ (ได้จากการ deploy ไฟล์ 3_ResultsWebApp.gs เป็น Web app)
// ตอนสร้างฟอร์มครั้งแรกยังไม่มีค่านี้ ให้ deploy ก่อน แล้วกลับมาแก้ตรงนี้
// จากนั้นรันฟังก์ชัน updateConfirmationMessage() เพื่ออัปเดตข้อความหน้ายืนยันโดยไม่ต้องสร้างฟอร์มใหม่
var RESULTS_URL = 'วาง Web app URL ที่นี่';

var FORM_INTRO = [
  'แบบทดสอบนี้ช่วยให้ท่านเลือกระดับการอบรมที่เหมาะกับตนเองที่สุด',
  '',
  '• ไม่มีผลต่อคะแนนในรายวิชาใด ๆ และไม่ใช่การสอบเพื่อคัดออก',
  '• กรุณาทำด้วยตนเองโดยไม่เปิดโปรแกรมหรือค้นหาคำตอบ เพราะหากผลสูงกว่าความสามารถจริง',
  '  ท่านจะได้รับคำแนะนำให้ลงระดับที่ยากเกินไป ซึ่งเสียเวลาของท่านเอง',
  '• ใช้เวลาประมาณ 6–14 นาที ระบบจะจัดเส้นทางคำถามให้เหมาะกับท่านโดยอัตโนมัติ',
  '• เมื่อกดส่งแล้ว ท่านจะได้ลิงก์ดูผลฉบับเต็มทันที และระบบจะส่งสำเนาผลไปยังอีเมลของท่านไว้ด้วย',
  '',
  'ข้อมูลที่เก็บจะใช้เพื่อการจัดกลุ่มผู้เรียนและการปรับปรุงหลักสูตรเท่านั้น',
  'ตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล'
].join('\\n');

function buildPlacementForm() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription(FORM_INTRO)
      .setCollectEmail(true)
      .setProgressBar(true)
      .setAllowResponseEdits(false)
      .setLimitOneResponsePerUser(true)
      .setShowLinkToRespondAgain(false)
      .setIsQuiz(false)
      .setConfirmationMessage(confirmationText());

  // ---------- ส่วนที่ 1 : ข้อมูลผู้เรียน ----------
  form.addTextItem().setTitle('1. ชื่อ–นามสกุล').setRequired(true);
  form.addTextItem().setTitle('2. รหัสนักศึกษา (บุคลากรหรือบุคคลภายนอกเว้นว่างได้)').setRequired(false);
  form.addTextItem().setTitle('3. สาขาวิชา / หน่วยงาน').setRequired(true);
  form.addMultipleChoiceItem()
      .setTitle('4. สถานภาพ')
      .setChoiceValues(['นักศึกษาปี 1', 'นักศึกษาปี 2', 'นักศึกษาปี 3', 'นักศึกษาปี 4', 'บัณฑิตศึกษา', 'บุคลากร', 'อื่น ๆ'])
      .setRequired(true);
  form.addScaleItem()
      .setTitle('5. ท่านประเมินความสามารถการใช้สเปรดชีตของตนเองอยู่ในระดับใด')
      .setBounds(1, 5)
      .setLabels('ไม่เคยใช้เลย', 'ใช้ได้คล่องมาก')
      .setRequired(true);
  form.addCheckboxItem()
      .setTitle('6. ท่านเคยทำสิ่งใดต่อไปนี้มาก่อน (ตอบได้หลายข้อ)')
      .setChoiceValues([
        'ป้อนข้อมูลและจัดรูปแบบตาราง',
        'ใช้สูตรพื้นฐาน เช่น SUM หรือ AVERAGE',
        'สร้างกราฟ',
        'ใช้ VLOOKUP หรือ XLOOKUP',
        'สร้าง PivotTable',
        'เขียนสคริปต์หรือมาโคร',
        'ใช้ AI ช่วยเขียนสูตรหรือโค้ด',
        'ยังไม่เคยทำข้อใดเลย'
      ])
      .setRequired(true);

  // ---------- สร้างหน้า (page break) ทั้งหมดก่อน เพื่อให้อ้างอิงข้ามหน้าได้ ----------
  var pageA = form.addPageBreakItem()
      .setTitle('ส่วนที่ 2 — พื้นฐานการจัดการข้อมูล')
      .setHelpText('มี 6 ข้อ ตอบตามความเข้าใจของท่าน ไม่ต้องเปิดโปรแกรมประกอบ');
  var itemsA = addBlock(form, BLOCK_A);

  var pageB = form.addPageBreakItem()
      .setTitle('ส่วนที่ 3 — การเชื่อมโยงและสรุปข้อมูล')
      .setHelpText('ยินดีด้วย ท่านผ่านส่วนแรกแล้ว ส่วนนี้มี 6 ข้อ');
  var itemsB = addBlock(form, BLOCK_B);

  var pageC = form.addPageBreakItem()
      .setTitle('ส่วนที่ 4 — การทำงานร่วมกับ AI และระบบอัตโนมัติ')
      .setHelpText('ส่วนสุดท้าย มี 6 ข้อ');
  var itemsC = addBlock(form, BLOCK_C);

  // หน้าจบสามหน้า ใช้ข้อความเหมือนกันทั้งหมด ไม่บอกระดับ (ดูเหตุผลในหมายเหตุสำคัญ 2 ด้านบน)
  // ยังต้องมีสามหน้า เพราะแต่ละหน้าเป็นปลายทางของเส้นทางที่ต่างกัน
  var pageBeg = form.addPageBreakItem().setTitle('ทำแบบทดสอบครบแล้ว').setHelpText(END_TEXT);
  var pageInt = form.addPageBreakItem().setTitle('ทำแบบทดสอบครบแล้ว').setHelpText(END_TEXT);
  var pageAdv = form.addPageBreakItem().setTitle('ทำแบบทดสอบครบแล้ว').setHelpText(END_TEXT);

  // หน้าผลลัพธ์ทุกหน้าต้องจบแบบทดสอบทันที
  pageBeg.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  pageInt.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  pageAdv.setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // ---------- ผูกการแตกสาขาไว้กับข้อชี้ขาดของแต่ละบล็อก ----------
  // A6 : ถูก -> ทำ Block B ต่อ / ผิด -> จบที่ผลระดับ Beginner
  setAnchorRouting(itemsA.anchor, itemsA.anchorSpec, pageB, pageBeg);
  // B6 : ถูก -> ทำ Block C ต่อ / ผิด -> จบที่ผลระดับ Intermediate
  setAnchorRouting(itemsB.anchor, itemsB.anchorSpec, pageC, pageInt);
  // C6 : ถูก -> ผลระดับ Advanced / ผิด -> ผลระดับ Intermediate
  setAnchorRouting(itemsC.anchor, itemsC.anchorSpec, pageAdv, pageInt);

  Logger.log('สร้างฟอร์มเรียบร้อย');
  Logger.log('ลิงก์สำหรับผู้ตอบ : ' + form.getPublishedUrl());
  Logger.log('ลิงก์สำหรับแก้ไข : ' + form.getEditUrl());
  return form.getEditUrl();
}

/** เพิ่มคำถามทั้งบล็อก คืนค่าอ็อบเจกต์ของข้อชี้ขาดไว้ผูกการแตกสาขาภายหลัง */
function addBlock(form, spec) {
  var anchorItem = null, anchorSpec = null;
  for (var i = 0; i < spec.length; i++) {
    var s = spec[i];
    var item = form.addMultipleChoiceItem()
        .setTitle(s.id + '. ' + s.stem)
        .setRequired(true);
    if (s.anchor) {
      // ยังไม่ตั้งตัวเลือก เพราะต้องรอให้สร้างหน้าปลายทางครบก่อน
      anchorItem = item; anchorSpec = s;
    } else {
      item.setChoiceValues(s.choices);
    }
  }
  return { anchor: anchorItem, anchorSpec: anchorSpec };
}

/** ตั้งตัวเลือกของข้อชี้ขาด พร้อมกำหนดหน้าปลายทางของแต่ละตัวเลือก */
function setAnchorRouting(item, spec, pageIfCorrect, pageIfWrong) {
  var choices = [];
  for (var i = 0; i < spec.choices.length; i++) {
    var text = spec.choices[i];
    var target = (text === spec.answer) ? pageIfCorrect : pageIfWrong;
    choices.push(item.createChoice(text, target));
  }
  item.setChoices(choices);
}

var END_TEXT = 'ท่านทำแบบทดสอบครบตามเส้นทางที่ระบบจัดให้แล้ว\\n\\n'
  + 'กรุณากดปุ่มส่ง จากนั้นระบบจะแสดงลิงก์ให้ดูผลฉบับเต็มทันที '
  + 'ซึ่งจะบอกระดับที่เหมาะกับท่าน คะแนนรายส่วน หัวข้อที่ควรให้ความสนใจ และลิงก์ลงทะเบียน';

/** ข้อความบนหน้ายืนยันหลังกดส่ง — Google Form จะทำ URL ให้เป็นลิงก์กดได้โดยอัตโนมัติ */
function confirmationText() {
  return 'ส่งแบบทดสอบเรียบร้อยแล้ว ขอบคุณครับ/ค่ะ\\n\\n'
    + '▶ ดูผลฉบับเต็มของท่านได้ทันทีที่ลิงก์นี้\\n'
    + RESULTS_URL + '\\n\\n'
    + '(กรุณาเปิดด้วยบัญชี Google ของสถาบันบัญชีเดียวกับที่ใช้ทำแบบทดสอบ '
    + 'ระบบใช้บัญชีนี้ยืนยันตัวตน จึงไม่มีผู้อื่นเห็นผลของท่านได้)\\n\\n'
    + 'ระบบได้ส่งสำเนาผลไปยังอีเมลของท่านไว้ด้วยแล้ว';
}

/**
 * อัปเดตข้อความหน้ายืนยันหลังจาก deploy หน้าเว็บผลลัพธ์แล้ว
 * ใช้เมื่อสร้างฟอร์มไปก่อนแล้วเพิ่งได้ Web app URL มา จะได้ไม่ต้องสร้างฟอร์มใหม่
 * วิธีใช้: แก้ค่า RESULTS_URL ด้านบน แล้วรันฟังก์ชันนี้พร้อมใส่รหัสฟอร์ม
 */
function updateConfirmationMessage(formId) {
  var id = formId || WEBAPP.FORM_ID;   // ใช้ค่าจากไฟล์ 3_ResultsWebApp.gs ได้ถ้ากรอกไว้แล้ว
  FormApp.openById(id).setConfirmationMessage(confirmationText());
  Logger.log('อัปเดตข้อความหน้ายืนยันเรียบร้อย');
}

// ============================================================
// คลังข้อสอบ — แก้ไขข้อความได้ที่นี่ แล้วรัน buildPlacementForm() ใหม่
// ลำดับตัวเลือกถูกสลับไว้แล้ว คำตอบที่ถูกระบุไว้ในคีย์ answer
// ============================================================
`;

bank.blocks.forEach(blk => {
  f1 += `\nvar BLOCK_${blk.code} = [\n`;
  blk.items.forEach((it, idx) => {
    const sh = shuffle(it.choices, 7919 + idx * 131 + blk.code.charCodeAt(0));
    f1 += `  {\n`;
    f1 += `    id: ${q(it.id)},\n`;
    f1 += `    anchor: ${it.anchor ? 'true' : 'false'},\n`;
    f1 += `    stem: ${q(cleanStem(it.stem))},\n`;
    f1 += `    choices: [\n${sh.map(c => '      ' + q(c)).join(',\n')}\n    ],\n`;
    f1 += `    answer: ${q(it.choices[it.answer])}\n`;
    f1 += `  }${idx < blk.items.length - 1 ? ',' : ''}\n`;
  });
  f1 += `];\n`;
});

fs.writeFileSync(__dirname + '/1_BuildPlacementForm.gs', f1);

/* ================= FILE 2 : SCORE AND NOTIFY ================= */
const keyLines = [];
bank.blocks.forEach(blk => blk.items.forEach(it => {
  keyLines.push(`  { id: ${q(it.id)}, block: ${q(blk.code)}, anchor: ${it.anchor ? 'true' : 'false'}, topic: ${q(it.topic)}, answer: ${q(it.choices[it.answer])} }`);
}));

const f2 = `/**
 * 2_ScoreAndNotify.gs
 * คำนวณคะแนนแบบทดสอบจัดระดับ ตัดสินระดับที่แนะนำ บันทึกผล และส่งอีเมลแจ้งผู้เรียน
 *
 * ชุดการเรียนรู้ Excel & Google Sheets 3 ระดับ
 * ผู้พัฒนา: อาจารย์ ดร.เนติยา การะเกตุ
 *
 * วิธีติดตั้ง
 *   1) เชื่อมฟอร์มเข้ากับสเปรดชีตนี้แล้ว (Responses > Link to Sheets)
 *   2) แก้ค่าคงที่ในส่วน CONFIG ด้านล่างให้เป็นลิงก์และอีเมลจริง
 *   3) รันฟังก์ชัน installTrigger() หนึ่งครั้ง
 *   4) ทดสอบด้วยฟังก์ชัน testDecisionTable() เพื่อตรวจว่าเกณฑ์ตัดสินทำงานถูกต้อง
 */

// ======================= CONFIG =======================
var CONFIG = {
  REGISTER_URL: {
    Beginner:     'https://forms.gle/แก้เป็นลิงก์ลงทะเบียนระดับ-Beginner',
    Intermediate: 'https://forms.gle/แก้เป็นลิงก์ลงทะเบียนระดับ-Intermediate',
    Advanced:     'https://forms.gle/แก้เป็นลิงก์ลงทะเบียนระดับ-Advanced'
  },
  LEVEL_FULLNAME: {
    Beginner:     'Beginner — รากฐานของข้อมูลที่เชื่อถือได้',
    Intermediate: 'Intermediate — จากตารางดิบสู่ข้อสรุป',
    Advanced:     'Advanced with AI — ออกแบบระบบให้ข้อมูลทำงานแทนเรา'
  },
  // URL ของหน้าเว็บผลลัพธ์ (จาก Deploy ไฟล์ 3_ResultsWebApp.gs) ใช้แนบไปกับอีเมลด้วย
  RESULTS_URL: 'วาง Web app URL ที่นี่',
  ADMIN_EMAIL: 'netiya.karaket@gmail.com',   // แก้เป็นอีเมลผู้ดูแลจริง
  RESULT_SHEET: 'Placement_Results',
  SENDER_NAME: 'ทีมอบรม Excel & Google Sheets',
  NOTIFY_ADMIN_ON_FLAG: true,                 // ส่งสำเนาให้ผู้ดูแลเมื่อมีธงเตือน
  // ผู้ทำแบบทดสอบเห็นผลฉบับเต็มบนหน้าเว็บทันทีอยู่แล้ว อีเมลจึงทำหน้าที่เป็นบันทึกสำรอง
  // ตั้งเป็น false ได้ถ้าไม่ต้องการส่งอีเมลเลย หรือเมื่อโควตาอีเมลของบัญชีใกล้เต็ม
  SEND_EMAIL: true
};

// ======================= ANSWER KEY =======================
// ตรวจคำตอบจาก “ข้อความ” ไม่ใช่ตำแหน่ง จึงสลับลำดับตัวเลือกในฟอร์มได้อย่างปลอดภัย
var KEY = [
${keyLines.join(',\n')}
];

// ======================= TRIGGER =======================
function installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'onPlacementSubmit') ScriptApp.deleteTrigger(existing[i]);
  }
  ScriptApp.newTrigger('onPlacementSubmit').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log('ติดตั้งทริกเกอร์เรียบร้อย');
}

// ======================= MAIN =======================
function onPlacementSubmit(e) {
  try {
    var nv = e.namedValues;
    var email = firstValue(nv, 'Email Address') || firstValue(nv, 'อีเมล') || findByPrefix(nv, 'Email');
    var name = findByPrefix(nv, '1.');
    var studentId = findByPrefix(nv, '2.');
    var major = findByPrefix(nv, '3.');
    var status = findByPrefix(nv, '4.');
    var confidence = parseInt(findByPrefix(nv, '5.'), 10) || 0;

    var scored = scoreResponse(nv);
    var result = decideLevel(scored, confidence);

    writeResultRow(e, {
      email: email, name: name, studentId: studentId, major: major,
      status: status, confidence: confidence
    }, scored, result);

    if (email && CONFIG.SEND_EMAIL) sendResultEmail(email, name, scored, result);
    if (CONFIG.NOTIFY_ADMIN_ON_FLAG && result.flag && CONFIG.ADMIN_EMAIL) {
      MailApp.sendEmail(CONFIG.ADMIN_EMAIL,
        '[Placement] ต้องพิจารณา: ' + (name || email),
        'ผู้เรียน: ' + name + ' (' + email + ')\\n' +
        'ระดับที่ระบบแนะนำ: ' + result.level + '\\n' +
        'ธงเตือน: ' + result.flag + '\\n' +
        'คะแนน A/B/C: ' + scored.A + '/' + (scored.reachedB ? scored.B : '-') + '/' + (scored.reachedC ? scored.C : '-'));
    }
  } catch (err) {
    Logger.log('ERROR: ' + err);
    if (CONFIG.ADMIN_EMAIL) {
      MailApp.sendEmail(CONFIG.ADMIN_EMAIL, '[Placement] เกิดข้อผิดพลาด', String(err));
    }
  }
}

// ======================= SCORING =======================
/** ตรวจคำตอบทุกข้อที่ผู้ตอบได้ทำ (ข้อที่ระบบข้ามให้จะไม่มีคำตอบ) */
function scoreResponse(nv) {
  var s = {
    A: 0, B: 0, C: 0,
    anchorA: 0, anchorB: 0, anchorC: 0,
    answered: 0, correct: 0,
    reachedB: false, reachedC: false,
    wrongTopics: []
  };
  for (var i = 0; i < KEY.length; i++) {
    var k = KEY[i];
    var given = findByPrefix(nv, k.id + '.');
    if (given === '' || given === null) continue;   // ผู้ตอบไม่ได้ทำข้อนี้
    s.answered++;
    if (k.block === 'B') s.reachedB = true;
    if (k.block === 'C') s.reachedC = true;
    var ok = (normalize(given) === normalize(k.answer));
    if (ok) {
      s.correct++;
      s[k.block]++;
      if (k.anchor) s['anchor' + k.block] = 1;
    } else {
      s.wrongTopics.push(k.topic);
    }
  }
  s.ratio = s.answered > 0 ? (s.correct / s.answered) : 0;
  return s;
}

/**
 * ตารางตัดสิน (ตรงกับเอกสารออกแบบ ข้อ 5.1) และตาข่ายนิรภัย (ข้อ 5.2)
 * คืนค่า { level, path, flag, note }
 */
function decideLevel(s, confidence) {
  var level, path, flag = '', note = '';

  if (s.anchorA === 0) {
    level = 'Beginner';
    if (s.A >= 4) { path = 'P2'; flag = 'ทบทวน: ทำข้ออื่นได้ดีแต่พลาดข้อชี้ขาด อาจอนุญาตให้ข้ามระดับได้'; }
    else { path = 'P1'; }
  } else if (!s.reachedB || s.anchorB === 0) {
    level = 'Intermediate';
    if (s.B >= 4) { path = 'P4'; flag = 'ทบทวน: ทำข้ออื่นได้ดีแต่พลาดข้อชี้ขาด อาจอนุญาตให้ข้ามระดับได้'; }
    else { path = 'P3'; }
  } else {
    // ผ่านทั้ง A และ B จึงได้ทำ Block C
    if (s.B <= 3) {
      level = 'Intermediate'; path = 'P5';
      flag = 'ทบทวน: เข้าใจหลักการเชิงระบบแต่ทักษะปฏิบัติใน Block B ยังไม่แน่น';
    } else if (s.C >= 5) {
      level = 'Advanced'; path = 'P7';
      note = 'พื้นฐานดีมาก เหมาะจะรับบทพี่เลี้ยงเพื่อนร่วมชั้น';
    } else {
      level = 'Advanced'; path = 'P6';
    }
  }

  // ตาข่ายนิรภัย: ตอบถูกน้อยกว่าครึ่งของข้อที่ได้ทำ ให้ลดระดับลงหนึ่งขั้น
  if (s.ratio < 0.5 && level !== 'Beginner') {
    level = (level === 'Advanced') ? 'Intermediate' : 'Beginner';
    path = path + '+SN';
    flag = 'ตาข่ายนิรภัย: ตอบถูกต่ำกว่าครึ่งของข้อที่ได้ทำ จึงลดระดับที่แนะนำลงหนึ่งขั้น';
  }

  // ธงการประเมินตนเองที่คลาดเคลื่อน
  var calib = 'สอดคล้อง';
  if (confidence >= 4 && level === 'Beginner') calib = 'ประเมินตนเองสูงกว่าผล';
  else if (confidence <= 2 && level === 'Advanced') calib = 'ประเมินตนเองต่ำกว่าผล';

  return { level: level, path: path, flag: flag, note: note, calibration: calib };
}

// ======================= OUTPUT =======================
function writeResultRow(e, p, s, r) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.RESULT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.RESULT_SHEET);
    sh.appendRow(['Timestamp', 'Email', 'ชื่อ-สกุล', 'รหัสนักศึกษา', 'สาขา/หน่วยงาน', 'สถานภาพ',
      'ความมั่นใจ (1-5)', 'A', 'B', 'C', 'anchorA', 'anchorB', 'anchorC',
      'ข้อที่ทำ', 'ตอบถูก', 'สัดส่วนถูก', 'เส้นทาง', 'ระดับที่แนะนำ', 'ธงเตือน', 'การประเมินตนเอง', 'หัวข้อที่ควรทบทวน']);
    sh.getRange(1, 1, 1, 21).setFontWeight('bold').setBackground('#D9E2F3');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), p.email, p.name, p.studentId, p.major, p.status, p.confidence,
    s.A, s.reachedB ? s.B : '', s.reachedC ? s.C : '',
    s.anchorA, s.reachedB ? s.anchorB : '', s.reachedC ? s.anchorC : '',
    s.answered, s.correct, Math.round(s.ratio * 100) / 100,
    r.path, r.level, r.flag, r.calibration, s.wrongTopics.join(' · ')]);
}

function sendResultEmail(email, name, s, r) {
  var full = CONFIG.LEVEL_FULLNAME[r.level];
  var url = CONFIG.REGISTER_URL[r.level];
  var rows = 'Block A (พื้นฐานการจัดการข้อมูล): ' + s.A + '/6';
  rows += '\\nBlock B (เชื่อมโยงและสรุปข้อมูล): ' + (s.reachedB ? s.B + '/6' : 'ไม่ได้ทำ — ระบบจัดเส้นทางให้');
  rows += '\\nBlock C (AI และระบบอัตโนมัติ): ' + (s.reachedC ? s.C + '/6' : 'ไม่ได้ทำ — ระบบจัดเส้นทางให้');

  var review = '';
  if (s.wrongTopics.length > 0) {
    var uniq = [];
    for (var i = 0; i < s.wrongTopics.length && uniq.length < 4; i++) {
      if (uniq.indexOf(s.wrongTopics[i]) < 0) uniq.push(s.wrongTopics[i]);
    }
    review = '\\n\\nหัวข้อที่แนะนำให้ทบทวนหรือให้ความสนใจเป็นพิเศษระหว่างอบรม\\n• ' + uniq.join('\\n• ');
  }

  var calibNote = '';
  if (r.calibration === 'ประเมินตนเองสูงกว่าผล') {
    calibNote = '\\n\\nหมายเหตุ: ท่านประเมินความสามารถของตนเองไว้ค่อนข้างสูง ซึ่งเป็นเรื่องปกติ '
      + 'ระดับที่แนะนำครอบคลุมกับดักที่ผู้ใช้สเปรดชีตมาหลายปีจำนวนมากก็ยังพลาด และใช้เวลาเพียง 3 ชั่วโมง';
  } else if (r.calibration === 'ประเมินตนเองต่ำกว่าผล') {
    calibNote = '\\n\\nหมายเหตุ: ท่านประเมินความสามารถของตนเองไว้ต่ำกว่าที่ผลทดสอบสะท้อนออกมา '
      + 'ขอให้มั่นใจและลงระดับที่ระบบแนะนำได้เลย';
  }
  if (r.note) calibNote += '\\n\\n' + r.note;

  var bodyText = 'เรียน ' + (name || 'ผู้เข้าร่วมอบรม') + '\\n\\n'
    + 'อีเมลฉบับนี้เป็นสำเนาผลไว้ให้ท่านเก็บไว้ ท่านสามารถเปิดดูผลฉบับเต็มพร้อมรายละเอียดได้ที่\\n'
    + CONFIG.RESULTS_URL + '\\n\\n'
    + 'ระดับที่แนะนำ: ' + full + ' (3 ชั่วโมง)\\n\\n'
    + 'คะแนนรายส่วน\\n' + rows
    + review + calibNote + '\\n\\n'
    + 'ลงทะเบียนระดับที่แนะนำได้ที่ลิงก์นี้\\n' + url + '\\n\\n'
    + 'หากท่านเห็นว่าผลนี้ยังไม่ตรงกับความต้องการ เช่น มีงานเฉพาะที่ต้องใช้ทักษะระดับอื่น '
    + 'สามารถตอบกลับอีเมลฉบับนี้เพื่อขอลงระดับอื่นได้ ผู้สอนจะพิจารณาเป็นรายกรณี '
    + 'ผลการทดสอบเป็นคำแนะนำ ไม่ใช่คำตัดสิน\\n\\n'
    + 'ด้วยความนับถือ\\n' + CONFIG.SENDER_NAME + '\\n\\n'
    + '— ข้อมูลของท่านถูกเก็บเพื่อการจัดกลุ่มผู้เรียนและการปรับปรุงหลักสูตรเท่านั้น '
    + 'ตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล';

  MailApp.sendEmail({
    to: email,
    subject: 'ผลการทดสอบจัดระดับ: ระบบแนะนำระดับ ' + r.level + ' — Excel & Google Sheets',
    body: bodyText,
    name: CONFIG.SENDER_NAME
  });
}

// ======================= HELPERS =======================
function normalize(v) {
  return String(v).replace(/\\s+/g, ' ').trim();
}
function firstValue(nv, key) {
  if (nv[key] && nv[key][0]) return String(nv[key][0]).trim();
  return '';
}
/** ค้นคำตอบจากหัวข้อคำถามที่ขึ้นต้นด้วย prefix เช่น 'A6.' ทนต่อการแก้ข้อความคำถามภายหลัง */
function findByPrefix(nv, prefix) {
  for (var k in nv) {
    if (k.indexOf(prefix) === 0) {
      var v = nv[k];
      return (v && v[0] !== undefined) ? String(v[0]).trim() : '';
    }
  }
  return '';
}

// ======================= SELF-TEST =======================
/** ทดสอบเกณฑ์ตัดสินด้วยกรณีจำลอง รันแล้วดูผลใน Execution log */
function testDecisionTable() {
  var cases = [
    { name: 'P1 พื้นฐานยังไม่แน่น', s: mk({ A: 2, anchorA: 0, answered: 6, correct: 2 }), conf: 2, expect: 'Beginner' },
    { name: 'P2 พลาดเฉพาะข้อชี้ขาด', s: mk({ A: 4, anchorA: 0, answered: 6, correct: 4 }), conf: 5, expect: 'Beginner' },
    { name: 'P3 ผ่าน A ตก B', s: mk({ A: 6, anchorA: 1, B: 2, anchorB: 0, reachedB: true, answered: 12, correct: 8 }), conf: 3, expect: 'Intermediate' },
    { name: 'P4 ตก B แบบก้ำกึ่ง', s: mk({ A: 6, anchorA: 1, B: 4, anchorB: 0, reachedB: true, answered: 12, correct: 10 }), conf: 3, expect: 'Intermediate' },
    { name: 'P5 ถึง C แต่ B อ่อน', s: mk({ A: 5, anchorA: 1, B: 3, anchorB: 1, reachedB: true, C: 4, anchorC: 1, reachedC: true, answered: 18, correct: 12 }), conf: 3, expect: 'Intermediate' },
    { name: 'P6 เหมาะกับ Advanced', s: mk({ A: 6, anchorA: 1, B: 5, anchorB: 1, reachedB: true, C: 4, anchorC: 1, reachedC: true, answered: 18, correct: 15 }), conf: 4, expect: 'Advanced' },
    { name: 'P7 พื้นดีมาก', s: mk({ A: 6, anchorA: 1, B: 6, anchorB: 1, reachedB: true, C: 6, anchorC: 1, reachedC: true, answered: 18, correct: 18 }), conf: 2, expect: 'Advanced' },
    // เดาถูกเฉพาะข้อชี้ขาดทั้งสามข้อ แต่คะแนนรวมเพียง 6/18 ตาข่ายนิรภัยต้องดึงลงจาก P5 (Intermediate) เป็น Beginner
    { name: 'ตาข่ายนิรภัย เดาถูกเฉพาะข้อชี้ขาด', s: mk({ A: 2, anchorA: 1, B: 2, anchorB: 1, reachedB: true, C: 2, anchorC: 1, reachedC: true, answered: 18, correct: 6 }), conf: 3, expect: 'Beginner' },
    // คะแนนรวม 8/18 = 0.44 ผ่าน B ระดับ 4 จึงได้ P6 แล้วถูกตาข่ายนิรภัยดึงลงเป็น Intermediate
    { name: 'ตาข่ายนิรภัย ดึง Advanced ลงเป็น Intermediate', s: mk({ A: 2, anchorA: 1, B: 4, anchorB: 1, reachedB: true, C: 2, anchorC: 1, reachedC: true, answered: 18, correct: 8 }), conf: 3, expect: 'Intermediate' }
  ];
  var pass = 0;
  for (var i = 0; i < cases.length; i++) {
    var r = decideLevel(cases[i].s, cases[i].conf);
    var ok = (r.level === cases[i].expect);
    if (ok) pass++;
    Logger.log((ok ? 'PASS' : 'FAIL') + ' | ' + cases[i].name + ' | ได้ ' + r.level + ' (' + r.path + ') คาดหวัง ' + cases[i].expect + (r.flag ? ' | ธง: ' + r.flag : ''));
  }
  Logger.log('ผลรวม: ' + pass + '/' + cases.length + ' กรณีผ่าน');
  return pass + '/' + cases.length;
}
function mk(o) {
  var s = { A: 0, B: 0, C: 0, anchorA: 0, anchorB: 0, anchorC: 0, answered: 0, correct: 0, reachedB: false, reachedC: false, wrongTopics: [] };
  for (var k in o) s[k] = o[k];
  s.ratio = s.answered > 0 ? s.correct / s.answered : 0;
  return s;
}
`;

fs.writeFileSync(__dirname + '/2_ScoreAndNotify.gs', f2);

/* ================= CSV ITEM BANK ================= */
const esc = v => '"' + String(v).replace(/"/g, '""') + '"';
let csv = ['block', 'item_id', 'anchor', 'clo', 'topic', 'bloom', 'stem', 'choice_correct', 'distractor_1', 'distractor_2', 'distractor_3', 'rationale'].join(',') + '\n';
bank.blocks.forEach(blk => blk.items.forEach(it => {
  const wrong = it.choices.filter((_, i) => i !== it.answer);
  csv += [blk.code, it.id, it.anchor ? 'YES' : '', it.clo, it.topic, it.bloom,
    cleanStem(it.stem), it.choices[it.answer], wrong[0], wrong[1], wrong[2], it.rationale].map(esc).join(',') + '\n';
}));
fs.writeFileSync(__dirname + '/Placement_ItemBank.csv', '﻿' + csv);

console.log('generated: 1_BuildPlacementForm.gs, 2_ScoreAndNotify.gs, Placement_ItemBank.csv');
