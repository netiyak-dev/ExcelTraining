/**
 * 8_LibraryWebApp.gs
 * จุดจ่ายงานของ Web App นี้ (doGet) + หน้าคลังไฟล์ฝึกปฏิบัติ 7 สาขา แบบกั้นสิทธิ์
 *
 * เส้นทาง (ผูกกับ URL เดียวกันของ deployment นี้):
 *   ?level=Beginner|Intermediate|Advanced  -> renderLesson_() ใน 7_LessonWebApp.gs
 *   ?view=library                          -> renderLibrary_() ในไฟล์นี้
 *   ไม่มีพารามิเตอร์                        -> หน้าอธิบายสั้น ๆ
 *
 * ตรวจสิทธิ์ด้วย Roster เดียวกับหน้าเรียน (ลงทะเบียนระดับใดก็ได้ ถือว่าเข้าคลังไฟล์ได้ทั้งหมด
 * เพราะคลังไฟล์เป็นทรัพยากรกลางที่ใช้ร่วมกันทั้งสามระดับ ไม่ได้แยกสิทธิ์ตามระดับเหมือนหน้าเรียน)
 *
 * เมื่อเข้าถึงได้ สคริปต์จะเรียก addViewer(email) ให้อัตโนมัติ จึงกดเปิดไฟล์จริงใน Drive ได้ทันที
 * ไม่ต้องขอสิทธิ์เพิ่มเอง ไฟล์ต้นฉบับอยู่ในโฟลเดอร์ส่วนตัวของผู้สอน ไม่มีใครเห็นได้นอกจากคนที่ลงทะเบียนแล้ว
 */

var LIBRARY_FILES_ = [
  { code: 'AG', title: 'AG_แปลงทดลองข้าวโพด', major: 'วิทยาศาสตร์การเกษตร',
    fileId: '1IdgIDvrstQ18L0Gn9BrhpwkZYzCWaifC',
    note: 'ค่า -999 ที่ปนอยู่ในคอลัมน์ผลผลิต ถ้าเผลอนำไปเฉลี่ยจะทำให้ค่าเฉลี่ยติดลบหรือต่ำผิดปกติ' },
  { code: 'AC', title: 'AC_สมุดรายวันทั่วไป', major: 'การบัญชี',
    fileId: '1FJHfociqQ5Ajy2MCTkTkicPnZ6JwHxYj',
    note: 'มีใบสำคัญ 6 ใบที่เดบิตไม่เท่าเครดิต มองไม่เห็นถ้าไม่จับคู่ตามเลขที่ใบสำคัญ และซ่อนตัวดีขึ้นเพราะบางยอดถูกเก็บเป็นข้อความ' },
  { code: 'BA', title: 'BA_ยอดขายร้านกาแฟ', major: 'บริหารธุรกิจ',
    fileId: '11j54UKj71hvhndQHXsGpJSdIq-Ab2Hp2',
    note: 'ส่วนลดที่ปนกันระหว่างบาทกับเปอร์เซ็นต์ ถ้าเผลอบวกรวมกันจะได้ยอดที่ผิดโดยไม่มีสัญญาณเตือน' },
  { code: 'CB', title: 'CB_สำรวจกล้องดักถ่าย', major: 'ชีววิทยาเชิงอนุรักษ์',
    fileId: '1wopLH0FXyf5nB9YaBI73O5p1-03RTIOY',
    note: 'แถวซ้ำจากการบันทึกภาพเดียวกันสองครั้ง ทำให้ความถี่การพบสูงเกินจริง ส่งผลตรงต่อข้อสรุปเชิงอนุรักษ์' },
  { code: 'ED', title: 'ED_สถานีตรวจวัดสิ่งแวดล้อม', major: 'วิศวกรรมสิ่งแวดล้อมและการจัดการภัยพิบัติ',
    fileId: '1xzjGV53uq8zoWO_joAaKXZZWAkfTxwz1',
    note: 'วันที่หายไปจากอนุกรมเวลา 22 วัน ไม่มีอะไรบอกให้รู้เลยถ้าไม่ตรวจ ทำให้ค่าเฉลี่ยเคลื่อนที่คลาดเคลื่อนโดยไม่มี error' },
  { code: 'FT', title: 'FT_บันทึกควบคุมคุณภาพสายการผลิต', major: 'เทคโนโลยีการอาหาร',
    fileId: '1YjcaQuaFd6V5qaPsxbujqTGNBvDwjXQ1',
    note: 'ผลตรวจจุลินทรีย์ที่เป็น <10, ND, TNTC เป็นค่าที่ถูกต้องตามธรรมเนียมห้องปฏิบัติการ แต่คำนวณตรง ๆ ไม่ได้' },
  { code: 'GS', title: 'GS_ข้อมูลเจาะสำรวจน้ำบาดาล', major: 'ธรณีศาสตร์',
    fileId: '1ZXxi2wVZHgKRTnfLWRRmBBXcymmAPhWU',
    note: 'แถวที่ความลึกช่วงบนมากกว่าช่วงล่าง เป็นไปไม่ได้ทางกายภาพ แต่ Excel ไม่เตือน พบเมื่อคำนวณความหนาแล้วเห็นค่าติดลบ' }
];

function doGet(e) {
  if (e.parameter.level) return renderLesson_(e);
  if (e.parameter.view === 'library') return renderLibrary_(e);
  return page_(
    '<div class="head"><div class="eyebrow">ชุดการเรียนรู้ Excel &amp; Google Sheets</div>'
    + '<div class="hello">ไม่พบหน้าที่ระบุ</div></div>'
    + '<div class="card"><p class="why">กรุณากลับไปที่หน้าแรกของเว็บแล้วกดปุ่มที่ต้องการใหม่อีกครั้ง</p></div>',
    'ชุดการเรียนรู้ Excel & Google Sheets');
}

function renderLibrary_(e) {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (err) { email = ''; }

  if (!email) return page_(cardNoIdentity_(), 'คลังไฟล์ฝึกปฏิบัติ');
  if (!isRegisteredAny_(email)) return page_(cardNotRegisteredLibrary_(email), 'คลังไฟล์ฝึกปฏิบัติ');

  var h = '<div class="head"><div class="eyebrow">คลังไฟล์ฝึกปฏิบัติ</div>'
    + '<div class="hello">7 สาขาวิชา ใช้ได้ทุกระดับที่ลงทะเบียนไว้</div></div>';

  LIBRARY_FILES_.forEach(function (item) {
    var file;
    try {
      file = DriveApp.getFileById(item.fileId);
      file.addViewer(email);
    } catch (err) {
      // ผู้ใช้อาจเป็นเจ้าของไฟล์เองอยู่แล้ว (เช่นทดสอบด้วยบัญชีผู้สอน) addViewer จะ error เฉย ๆ ข้ามได้
    }
    var url = 'https://drive.google.com/file/d/' + item.fileId + '/view';
    h += '<div class="card">'
      + '<div class="label">' + esc_(item.code) + ' · ' + esc_(item.major) + '</div>'
      + '<div class="filetitle">' + esc_(item.title) + '</div>'
      + '<p class="why">' + esc_(item.note) + '</p>'
      + '<a class="btn" href="' + url + '" target="_blank">เปิดไฟล์ใน Drive</a>'
      + '</div>';
  });

  h += '<p class="note">เปิดแล้วให้ทำสำเนา (File &rarr; Make a copy) ก่อนเริ่มแก้ไข เพื่อไม่ให้ไฟล์ต้นฉบับเปลี่ยน</p>';
  return page_(h, 'คลังไฟล์ฝึกปฏิบัติ');
}

/** ลงทะเบียนไว้ระดับใดก็ได้อย่างน้อยหนึ่งระดับ ถือว่าเข้าคลังไฟล์ได้ */
function isRegisteredAny_(email) {
  var sheet = SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID).getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) return false;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;

  var emailCol = findColumn_(data[0], 'Email');
  if (emailCol < 0) return false;

  var target = String(email).toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailCol]).toLowerCase().trim() === target) return true;
  }
  return false;
}

function cardNotRegisteredLibrary_(email) {
  return '<div class="head"><div class="eyebrow">คลังไฟล์ฝึกปฏิบัติ</div><div class="hello">ยังไม่ได้ลงทะเบียน</div></div>'
    + '<div class="card"><p class="why">อีเมล ' + esc_(email) + ' ยังไม่ได้ลงทะเบียนเรียนเลย ต้องลงทะเบียนอย่างน้อยหนึ่งระดับก่อน</p>'
    + '<a class="btn" href="' + REGISTER_FORM_URL_ + '" target="_blank">ไปหน้าลงทะเบียน</a></div>';
}
