#!/usr/bin/env bash
# สร้างสื่อทั้งชุดจากซอร์สโค้ด · รันจากในโฟลเดอร์นี้เท่านั้น เพราะสคริปต์อ้างไฟล์ข้อมูลแบบชื่อล้วน
#
#   bash build_all.sh          สร้างทุกอย่าง
#   bash build_all.sh web      สร้างเฉพาะหน้าเว็บ ใช้ตอนแก้เนื้อหาบทเรียน ซึ่งเป็นงานที่ทำบ่อยที่สุด
#   NOINDEX=1 bash build_all.sh web   สร้างหน้าเว็บพร้อมกันเครื่องมือค้นหาเก็บหน้าเรียน
set -euo pipefail
cd "$(dirname "$0")"

OUT=ผลลัพธ์
WEB=$OUT/หน้าเว็บ
ONLY=${1:-all}

say() { printf '\n\033[1;34m== %s\033[0m\n' "$1"; }

say "หน้าเว็บ"
python3 build_site.py    "$WEB/index.html"
python3 build_lesson.py  "$WEB/หน้าเรียน-Beginner.html"     lesson_beginner
python3 build_lesson.py  "$WEB/หน้าเรียน-Intermediate.html" lesson_intermediate
python3 build_lesson.py  "$WEB/หน้าเรียน-Advanced.html"     lesson_advanced

if [ "$ONLY" = "web" ]; then
  say "เสร็จเฉพาะหน้าเว็บ"
  exit 0
fi

say "ไฟล์ฝึกปฏิบัติ 7 สาขา"
python3 build_samples.py  "$OUT/ไฟล์ตัวอย่าง"
python3 verify_samples.py "$OUT/ไฟล์ตัวอย่าง"
python3 answer_key.py     "$OUT/ไฟล์ตัวอย่าง"

say "ไฟล์เฉลยและ csv สำหรับนำเข้า Google Sheets"
python3 build_answers.py  "$OUT/ไฟล์ตัวอย่าง" "$OUT/เฉลย" "$OUT/csv"

say "เอกสาร"
node build.js              "$OUT/เอกสาร/Blueprint_หลักสูตร_Excel_GoogleSheets_3ระดับ.docx"
[ -f items.json ] && node build_placement.js "$OUT/เอกสาร/Placement_Test_แบบทดสอบจัดระดับ.docx" \
  || echo "  ข้ามเอกสารแบบทดสอบจัดระดับ เพราะไม่พบ items.json"
node build_guide.js        "$OUT/เอกสาร/คู่มืออาจารย์_คลังไฟล์ฝึกปฏิบัติ7สาขา.docx"
node build_setup_guide.js  "$OUT/เอกสาร/คู่มือติดตั้งระบบ.docx"

say "จดหมายเวียน"
node   build_merge_template.js "$OUT/จดหมายเวียน/แม่แบบจดหมายเวียน_หนังสือรับรองการอบรม.docx"
python3 build_merge_data.py    "$OUT/จดหมายเวียน/รายชื่อผู้รับ.xlsx"

say "สคริปต์แบบทดสอบจัดระดับ"
# items.json ถูกกันไม่ให้ขึ้น repo เพราะเป็นข้อสอบพร้อมเฉลย ดูเหตุผลใน README
# ถ้าไม่มีไฟล์นี้ ให้ข้ามขั้นนี้ไปโดยไม่ถือว่า build ล้มเหลว
if [ -f items.json ]; then
  node gen_gs.js "$OUT/เอกสาร"
else
  echo "  ข้ามขั้นนี้ เพราะไม่พบ items.json"
  echo "  ถ้าต้องการสร้างแบบทดสอบจัดระดับใหม่ ให้คัดลอก items.json มาวางในโฟลเดอร์นี้ก่อน"
fi

say "สไลด์สอน · รอบวัดระยะแล้วรอบสร้างจริง"
for d in beginner:Beginner intermediate:Intermediate advanced:Advanced_with_AI; do
  n=${d%%:*}; N=${d##*:}
  node "build_deck_$n.js" "$OUT/สไลด์/${N}_สไลด์สอน.pptx" --measure > /dev/null
  node "build_deck_$n.js" "$OUT/สไลด์/${N}_สไลด์สอน.pptx" | tail -2
done

say "เสร็จทั้งหมด · ผลลัพธ์อยู่ในโฟลเดอร์ $OUT"
