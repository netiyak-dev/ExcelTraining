/**
 * 7_LessonWebApp.gs
 * เสิร์ฟหน้าเรียนแบบกั้นสิทธิ์ ตรวจอีเมลผู้เปิดกับชีต Roster ก่อนแสดงเนื้อหา
 * ใช้ ROSTER_SPREADSHEET_ID และ ROSTER_SHEET_NAME ร่วมกับ 6_BuildRegisterForm.gs (ประกาศไว้ที่ไฟล์นั้น)
 *
 * เนื้อหาหน้าเรียนแต่ละระดับเก็บเป็นไฟล์ HTML แยกในโปรเจกต์นี้ (Beginner.html, Intermediate.html, Advanced.html)
 * ก็อปมาจากผลลัพธ์ของ build_all.sh web ในโฟลเดอร์หลักของโปรเจกต์ แล้ว push เข้ามาด้วย clasp
 * แก้เนื้อหาบทเรียนแล้วต้อง build ใหม่ + copy ไฟล์มาทับ + clasp push ทุกครั้ง ไม่งั้นหน้าเรียนจะไม่อัปเดต
 *
 * วิธี deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone within <โดเมนสถาบัน>
 */

var LEVELS_ = ['Beginner', 'Intermediate', 'Advanced'];

// วางลิงก์ลงทะเบียนทั่วไป (จาก Logger ตอนรัน buildRegisterForm) แทนบรรทัดนี้
var REGISTER_FORM_URL_ = 'https://docs.google.com/forms/d/e/1FAIpQLSc9N1kALMw3LzBC1Qoo-z92pCO6dfhQlhmhTvFnh3veQvMWAw/viewform';

function doGet(e) {
  var level = e.parameter.level;
  if (LEVELS_.indexOf(level) < 0) {
    return page_(
      '<div class="head"><div class="eyebrow">หน้าเรียน</div><div class="hello">ไม่พบระดับที่ระบุ</div></div>'
      + '<div class="card"><p class="why">กรุณากลับไปที่หน้าแรกแล้วกด "เปิดหน้าเรียน" ใหม่อีกครั้ง</p></div>',
      'หน้าเรียน');
  }

  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (err) { email = ''; }

  if (!email) return page_(cardNoIdentity_(), 'หน้าเรียน');
  if (!isRegistered_(email, level)) return page_(cardNotRegistered_(email, level), 'หน้าเรียน');

  return HtmlService.createHtmlOutputFromFile(level)
      .setTitle('บทเรียน ' + level)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** ตรวจว่าอีเมลนี้ลงทะเบียนระดับที่ขอไว้หรือไม่ จากชีต Roster (คอลัมน์ระดับอาจมีหลายค่าคั่นด้วยจุลภาค) */
function isRegistered_(email, level) {
  var sheet = SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID).getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) return false;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;

  var header = data[0];
  var emailCol = findColumn_(header, 'Email');
  var levelCol = findColumn_(header, 'ระดับที่ต้องการลงทะเบียน');
  if (emailCol < 0 || levelCol < 0) return false;

  var target = String(email).toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol]).toLowerCase().trim();
    if (rowEmail !== target) continue;
    var rowLevels = String(data[i][levelCol]);
    if (rowLevels.indexOf(level) >= 0) return true;
  }
  return false;
}

function findColumn_(header, prefix) {
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).indexOf(prefix) === 0) return i;
  }
  return -1;
}

function page_(inner, title) {
  var html = '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<title>' + esc_(title) + '</title>' + styles_() + '</head><body><div class="wrap">'
    + inner + '</div></body></html>';
  return HtmlService.createHtmlOutput(html)
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function cardNoIdentity_() {
  return '<div class="head"><div class="eyebrow">หน้าเรียน</div><div class="hello">ยังระบุตัวตนไม่ได้</div></div>'
    + '<div class="card"><p class="why">หน้านี้ต้องเปิดด้วยบัญชี Google ขององค์กร '
    + 'กรุณาเข้าสู่ระบบด้วยบัญชีสถาบันแล้วลองใหม่อีกครั้ง</p></div>';
}

function cardNotRegistered_(email, level) {
  return '<div class="head"><div class="eyebrow">หน้าเรียน</div><div class="hello">ยังไม่ได้ลงทะเบียนระดับนี้</div></div>'
    + '<div class="card"><p class="why">อีเมล ' + esc_(email) + ' ยังไม่ได้ลงทะเบียนระดับ ' + esc_(level) + '</p>'
    + '<a class="btn" href="' + REGISTER_FORM_URL_ + '" target="_blank">ไปหน้าลงทะเบียน</a></div>';
}

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function styles_() {
  return '<style>'
    + 'body{margin:0;background:#F4F6F9;font-family:"Sarabun","Noto Sans Thai",-apple-system,"Segoe UI",sans-serif;color:#1a1a1a;line-height:1.65}'
    + '.wrap{max-width:560px;margin:0 auto;padding:40px 18px}'
    + '.eyebrow{font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:#7a869a}'
    + '.hello{font-size:24px;font-weight:700;color:#1F3864;margin-top:4px;margin-bottom:18px}'
    + '.card{background:#fff;border-radius:12px;padding:22px;box-shadow:0 1px 3px rgba(16,24,40,.08)}'
    + '.why{font-size:15px;margin:0 0 14px}'
    + '.btn{display:inline-block;padding:12px 26px;border-radius:8px;background:#2E74B5;color:#fff;text-decoration:none;font-weight:700}'
    + '</style>';
}
