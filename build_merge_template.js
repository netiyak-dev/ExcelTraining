/**
 * แม่แบบจดหมายเวียน .docx ที่มีเขตข้อมูลพร้อมใช้
 * เขียนเขตข้อมูลเป็นข้อความ «ชื่อ» ไว้ให้เห็นตำแหน่ง แล้วให้ผู้เรียนแทนด้วยเขตข้อมูลจริงตามขั้นตอน
 * เหตุผลที่ไม่ฝังเขตข้อมูล MERGEFIELD มาให้เลย คือขั้นการแทรกเขตข้อมูลเองเป็นสิ่งที่ต้องฝึก
 */
const fs = require('fs');
const L = require('./docxlib.js');
const { P, H, T, SP, BOX, FONT, C, Document, Packer, Paragraph, TextRun, AlignmentType } = L;

const OUT = process.argv[2];
const body = [];

function line(text, opts) {
  return new Paragraph({
    spacing: { after: (opts && opts.after) || 120 },
    alignment: (opts && opts.align) || AlignmentType.LEFT,
    indent: (opts && opts.indent) ? { firstLine: opts.indent } : undefined,
    children: [new TextRun({
      text, size: (opts && opts.size) || 30, bold: !!(opts && opts.bold),
      font: FONT, color: (opts && opts.color) || '000000'
    })]
  });
}

body.push(line('บันทึกข้อความ', { size: 40, bold: true, align: AlignmentType.CENTER, after: 200 }));
body.push(line('มหาวิทยาลัยมหิดล วิทยาเขตกาญจนบุรี', { size: 26, align: AlignmentType.CENTER, color: '595959', after: 320 }));

body.push(line('ที่  ศธ ๐๕๑๗.๑๑/«เลขที่หนังสือ»', { after: 60 }));
body.push(line('วันที่  «วันที่หนังสือ»', { after: 200 }));
body.push(line('เรื่อง  «เรื่อง»', { bold: true, after: 200 }));
body.push(line('เรียน  «คำนำหน้า»«ชื่อ» «นามสกุล»  «ตำแหน่ง»', { after: 240 }));

body.push(line('ตามที่ «หน่วยงาน» ได้ดำเนินโครงการอบรมเชิงปฏิบัติการ เรื่อง การจัดการข้อมูลด้วย Excel '
  + 'และ Google Sheets เพื่อพัฒนาทักษะการวิเคราะห์ข้อมูลของบุคลากรและนักศึกษา นั้น', { indent: 720, after: 160 }));

body.push(line('ในการนี้ ขอเรียนว่า «คำนำหน้า»«ชื่อ» «นามสกุล» ได้เข้าร่วมการอบรมระดับ «ระดับที่อบรม» '
  + 'ในวันที่ «วันที่อบรม» รวมระยะเวลา «จำนวนชั่วโมง» ชั่วโมง และมีผลการประเมินหลังการอบรมคิดเป็นร้อยละ «คะแนนหลังเรียน» '
  + 'ซึ่งอยู่ในเกณฑ์ «ผลการประเมิน» จึงขอส่งหนังสือรับรองการเข้าร่วมมาพร้อมนี้', { indent: 720, after: 160 }));

body.push(line('จึงเรียนมาเพื่อโปรดทราบ', { indent: 720, after: 400 }));

body.push(line('ขอแสดงความนับถือ', { align: AlignmentType.CENTER, after: 700 }));
body.push(line('(อาจารย์ ดร.เนติยา การะเกตุ)', { align: AlignmentType.CENTER, after: 60 }));
body.push(line('ผู้รับผิดชอบโครงการ', { align: AlignmentType.CENTER, after: 400 }));

body.push(SP());
body.push(BOX('วิธีใช้ไฟล์นี้ · ลบกรอบนี้ทิ้งก่อนพิมพ์จริง', [
  'ข้อความที่อยู่ในเครื่องหมาย « » ทุกจุด คือตำแหน่งที่ต้องแทนด้วยเขตข้อมูลจริง ยังไม่ใช่เขตข้อมูล',
  'ขั้นที่ 1 เปิดแท็บ การส่งจดหมาย เลือก เริ่มจดหมายเวียน แล้วเลือก จดหมาย',
  'ขั้นที่ 2 เลือก เลือกผู้รับ แล้วเลือก ใช้รายการที่มีอยู่ ชี้ไปที่ไฟล์ รายชื่อผู้รับ.xlsx แล้วเลือกชีต รายชื่อ',
  'ขั้นที่ 3 ลากคลุมข้อความ «ชื่อ» แล้วกด แทรกเขตข้อมูลผสาน เลือกชื่อคอลัมน์ที่ตรงกัน ทำให้ครบทุกจุด',
  'ขั้นที่ 4 กด ดูตัวอย่างผลลัพธ์ แล้วกดลูกศรไล่ดูให้ครบทุกคน อย่าดูแค่คนแรก',
  'ขั้นที่ 5 กด เสร็จสิ้นและผสาน เลือก แก้ไขเอกสารแต่ละฉบับ เพื่อได้ไฟล์เดียวที่มีทุกฉบับต่อกัน',
  'ขั้นที่ 6 ตรวจสามอย่างก่อนพิมพ์ คือวันที่แสดงแบบไทยถูกต้อง ตัวเลขไม่มีทศนิยมเกิน และไม่มีจุดที่ยังเป็น « »'
], C.head2));
body.push(SP());
body.push(BOX('สองปัญหาที่ต้องแก้ด้วยรหัสรูปแบบ ไม่ใช่แก้ที่ Excel', [
  'วันที่กลายเป็น 9/15/2026 แทน 15 กันยายน 2569 ให้กด Alt+F9 เพื่อดูรหัสเขตข้อมูล '
  + 'แล้วเติมท้ายเป็น { MERGEFIELD วันที่อบรม \\@ "d MMMM yyyy" } แล้วกด Alt+F9 กลับ และกด F9 เพื่อรีเฟรช',
  'ตัวเลขกลายเป็น 87.33333333 แทน 87.33 ให้เติมท้ายเป็น { MERGEFIELD คะแนนหลังเรียน \\# "#,##0.00" } '
  + 'การจัดรูปแบบใน Excel ไม่ติดมากับจดหมายเวียน เพราะ Word อ่านค่าดิบ ไม่ได้อ่านรูปแบบที่แสดง',
  'สองเรื่องนี้คือคำถามที่ถูกถามมากที่สุดในการสอนจดหมายเวียน จึงต้องรู้ก่อนลงมือ'
], C.warn));

const doc = new Document({
  creator: 'Dr. Netiya Karaket',
  title: 'แม่แบบจดหมายเวียน หนังสือรับรองการเข้าร่วมอบรม',
  styles: { default: { document: { run: { font: FONT, size: 30 } } } },
  sections: [{
    properties: { page: { margin: { top: 1200, right: 1200, bottom: 1200, left: 1400 } } },
    children: body
  }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('เขียนแล้ว', OUT, (b.length / 1024).toFixed(0) + ' KB'); });
