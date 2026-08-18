/**
 * 6_BuildRegisterForm.gs
 * สร้าง Google Form สำหรับลงทะเบียนเรียน ผูกคำตอบเข้ากับสเปรดชีตผล Placement Test เดิม
 * (สเปรดชีตเดียวกับที่ 2_ScoreAndNotify.gs เขียนชีต Placement_Results ไว้)
 * แล้วเปลี่ยนชื่อชีตคำตอบที่ได้เป็น "Roster" เพื่อให้ 7_LessonWebApp.gs อ่านต่อได้
 *
 * ชุดการเรียนรู้ Excel & Google Sheets 3 ระดับ
 * ผู้พัฒนา: อาจารย์ ดร.เนติยา การะเกตุ
 *
 * วิธีใช้: รันฟังก์ชัน buildRegisterForm() หนึ่งครั้ง แล้วดูผลใน Execution log
 * ล็อกจะมีลิงก์แก้ไขฟอร์ม และลิงก์ลงทะเบียนแบบเลือกระดับไว้ล่วงหน้า 3 ลิงก์
 * เอาลิงก์ 3 อันนั้นไปใส่ใน CONFIG.register ของ build_site.py แทนของเดิมได้เลย
 * (คนกดยังแก้ checkbox เพิ่มระดับอื่นในฟอร์มได้ ไม่ได้จำกัดแค่ระดับเดียว)
 */

var FORM_TITLE = 'ลงทะเบียนเรียน — ชุดการเรียนรู้ Excel & Google Sheets';
var ROSTER_SPREADSHEET_ID = '1edMHpMAsWFikceTolleq5_yTRcx7pl_8wtdnlAZZbRk';
var ROSTER_SHEET_NAME = 'Roster';

// ต้องขึ้นต้นด้วยคำนี้ตรงตัว เพราะ 7_LessonWebApp.gs ใช้ indexOf หาว่าเลือกระดับนี้ไว้หรือไม่
var LEVEL_CHOICES = [
  'Beginner — รากฐานของข้อมูลที่เชื่อถือได้',
  'Intermediate — จากตารางดิบสู่ข้อสรุป',
  'Advanced — ออกแบบระบบให้ข้อมูลทำงานแทนเรา'
];

function buildRegisterForm() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription(
      'ลงทะเบียนเพื่อปลดล็อกหน้าเรียนของระดับที่ท่านต้องการ\n\n' +
      '• เลือกได้มากกว่าหนึ่งระดับ ถ้าต้องการเรียนหลายระดับ\n' +
      '• ใช้อีเมลบัญชี Google เดียวกับที่จะใช้เปิดหน้าเรียน ระบบยืนยันตัวตนด้วยอีเมลนี้เท่านั้น\n' +
      '• ลงทะเบียนแล้วสามารถกลับมาเพิ่มระดับอื่นได้ภายหลังโดยส่งฟอร์มนี้อีกครั้ง'
    )
    .setCollectEmail(true)
    .setProgressBar(false)
    .setAllowResponseEdits(false)
    .setLimitOneResponsePerUser(false)
    .setConfirmationMessage(
      'ลงทะเบียนเรียบร้อยแล้ว ขอบคุณครับ/ค่ะ\n\n' +
      'กลับไปที่หน้าแรกของเว็บ แล้วกด "เปิดหน้าเรียน" ของระดับที่ลงทะเบียนไว้ได้เลย ' +
      'ระบบจะขอให้เข้าสู่ระบบด้วยบัญชี Google อีเมลเดียวกับที่กรอกในฟอร์มนี้'
    );

  form.addTextItem().setTitle('ชื่อ–นามสกุล').setRequired(true);
  form.addTextItem().setTitle('รหัสนักศึกษา (บุคลากรหรือบุคคลภายนอกเว้นว่างได้)').setRequired(false);

  var levelItem = form.addCheckboxItem()
      .setTitle('ระดับที่ต้องการลงทะเบียน (เลือกได้มากกว่า 1 ข้อ)')
      .setChoiceValues(LEVEL_CHOICES)
      .setRequired(true);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ROSTER_SPREADSHEET_ID);

  var sheet = findResponseSheet_(SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID), form);
  if (sheet) sheet.setName(ROSTER_SHEET_NAME);

  Logger.log('สร้างฟอร์มเรียบร้อย');
  Logger.log('ลิงก์สำหรับแก้ไข : ' + form.getEditUrl());
  Logger.log('ลิงก์ลงทะเบียนทั่วไป (ไม่เลือกระดับล่วงหน้า) : ' + form.getPublishedUrl());

  LEVEL_CHOICES.forEach(function (choiceText) {
    var prefilled = form.createResponse()
        .withItemResponse(levelItem.createResponse([choiceText]))
        .toPrefilledUrl();
    Logger.log('ลิงก์ลงทะเบียนระดับ ' + choiceText.split(' — ')[0] + ' : ' + prefilled);
  });

  Logger.log('ชื่อชีตคำตอบ : ' + (sheet ? sheet.getName() : '(หาไม่เจอ ตรวจสอบ ROSTER_SPREADSHEET_ID)'));
  return form.getEditUrl();
}

/**
 * หาชีตคำตอบที่เพิ่งผูกกับฟอร์มนี้ จากสเปรดชีตปลายทาง
 * getFormUrl() ของชีตคืนค่าในรูปแบบเดียวกับ getEditUrl() ของฟอร์ม (มี form ID ตรงตัว)
 * ไม่ใช่รูปแบบเดียวกับ getPublishedUrl() จึงต้องเทียบด้วย form ID แทนการเทียบ URL ทั้งเส้น
 */
function findResponseSheet_(ss, form) {
  var sheets = ss.getSheets();
  var id = form.getId();
  for (var i = 0; i < sheets.length; i++) {
    var url = sheets[i].getFormUrl();
    if (url && url.indexOf(id) >= 0) return sheets[i];
  }
  return null;
}

/**
 * ใช้ซ่อมชื่อชีตครั้งเดียว ถ้า buildRegisterForm() รันไปแล้วแต่หาไม่เจอตอนเปลี่ยนชื่อ (เช่นก่อนแก้ findResponseSheet_)
 * ห้ามรัน buildRegisterForm() ซ้ำ เพราะจะสร้างฟอร์มใหม่ซ้ำอีกฉบับ ใช้ฟังก์ชันนี้แทน
 */
function fixRosterSheetName() {
  var form = FormApp.openById('1JDuPydnTmRe62wVNFl0Qg5xkpHC43VzagsSg-W21y5Q');
  var sheet = findResponseSheet_(SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID), form);
  if (!sheet) { Logger.log('ยังหาชีตไม่เจอ'); return; }
  sheet.setName(ROSTER_SHEET_NAME);
  Logger.log('เปลี่ยนชื่อชีตเป็น ' + ROSTER_SHEET_NAME + ' เรียบร้อย');
}
