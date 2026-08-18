/**
 * 9_BuildClassrooms.gs
 * สร้างห้องเรียน Google Classroom สามระดับให้อัตโนมัติ ใช้ Classroom Advanced Service
 * (ต้องเปิด Google Classroom API ใน Cloud project ของสคริปต์นี้ก่อน — ดูวิธีด้านล่าง)
 *
 * ชุดการเรียนรู้ Excel & Google Sheets 3 ระดับ
 *
 * วิธีใช้: รันฟังก์ชัน buildClassrooms() หนึ่งครั้ง แล้วดูผลใน Execution log
 * รันซ้ำได้อย่างปลอดภัย ถ้าเจอชื่อห้องที่มีอยู่แล้วจะข้ามไม่สร้างซ้ำ
 *
 * ถ้ารันแล้ว error ว่า "Classroom API has not been used..." ให้กดลิงก์ที่ error แจ้งมา
 * เพื่อเปิดใช้ Google Classroom API ใน Cloud Console แล้วรอ 1-2 นาทีค่อยรันใหม่
 */

var CLASSROOM_COURSES_ = [
  { level: 'Beginner', name: 'Excel & Google Sheets — Beginner', section: 'รากฐานของข้อมูลที่เชื่อถือได้' },
  { level: 'Intermediate', name: 'Excel & Google Sheets — Intermediate', section: 'จากตารางดิบสู่ข้อสรุป' },
  { level: 'Advanced', name: 'Excel & Google Sheets — Advanced with AI', section: 'ออกแบบระบบให้ข้อมูลทำงานแทนเรา' }
];

function buildClassrooms() {
  var existing = listMyCourses_();

  CLASSROOM_COURSES_.forEach(function (spec) {
    var found = existing.filter(function (c) { return c.name === spec.name; })[0];
    if (found) {
      Logger.log('มีอยู่แล้ว ข้าม: ' + spec.level + ' -> ' + joinLink_(found));
      return;
    }

    var course = Classroom.Courses.create({
      name: spec.name,
      section: spec.section,
      ownerId: 'me',
      courseState: 'ACTIVE'
    });

    Logger.log('สร้างห้องเรียนระดับ ' + spec.level + ' เรียบร้อย');
    Logger.log('ลิงก์เชิญเข้าห้องเรียน ' + spec.level + ' : ' + joinLink_(course));
    Logger.log('(ถ้าลิงก์ข้างบนใช้ไม่ได้ ให้เปิด classroom.google.com เข้าห้องนี้ด้วยตนเอง '
      + 'แล้วไปที่ Settings คัดลอก "ลิงก์เชิญ" จากหน้านั้นแทน)');
  });
}

function listMyCourses_() {
  var courses = [];
  var pageToken;
  do {
    var res = Classroom.Courses.list({ teacherId: 'me', pageToken: pageToken });
    if (res.courses) courses = courses.concat(res.courses);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return courses;
}

function joinLink_(course) {
  return 'https://classroom.google.com/c/' + course.id + '?cjc=' + course.enrollmentCode;
}
