#!/usr/bin/env python3
"""ประกอบหน้าเรียนหนึ่งระดับเป็นไฟล์ HTML ไฟล์เดียว

วิธีใช้  python3 build_lesson.py <ไฟล์ผลลัพธ์> [โมดูลเนื้อหา]
โมดูลเนื้อหาปริยายคือ lesson_beginner ตัวเรนเดอร์นี้ใช้ร่วมกันได้ทุกระดับ
"""
import sys, os, html, json, importlib

OUT = sys.argv[1]
CONTENT = sys.argv[2] if len(sys.argv) > 2 else "lesson_beginner"
# ตั้งตัวแปรแวดล้อม NOINDEX=1 เมื่อนำขึ้นโฮสต์สาธารณะ เพื่อไม่ให้เครื่องมือค้นหาเก็บหน้านี้
# เพราะหน้าเรียนมีเฉลยอยู่ในโค้ดฝั่งผู้ใช้ ซึ่งเปิดดูได้ด้วยคำสั่งดูซอร์ส
NOINDEX = os.environ.get("NOINDEX") == "1"
sys.argv = [sys.argv[0], OUT]
L = importlib.import_module(CONTENT)

e = html.escape
C = L.C


# ---------------------------------------------------------------- ชิ้นส่วนเนื้อหา
def block(b):
    k = b["kind"]
    if k == "grid4":
        cells = "".join(
            '<div class="g4c"><b>{}</b><span>{}</span></div>'.format(e(a), e(t))
            for a, t in b["items"])
        return '<h4>{}</h4><div class="g4">{}</div>'.format(e(b["title"]), cells)

    if k == "twocol":
        left = "".join("<li>{}</li>".format(e(x)) for x in b["good"])
        right = "".join("<li>{}</li>".format(e(x)) for x in b["bad"])
        lcls, rcls = ("box ok", "box no") if not b.get("swap") else ("box ok", "box ok2")
        return ('<h4>{}</h4><div class="two">'
                '<div class="{}"><b>{}</b><ul>{}</ul></div>'
                '<div class="{}"><b>{}</b><ul>{}</ul></div></div>').format(
            e(b["title"]), lcls, e(b["good_h"]), left, rcls, e(b["bad_h"]), right)

    if k == "table":
        head = "".join("<th>{}</th>".format(e(h)) for h in b["head"])
        rows = "".join("<tr>" + "".join(
            ("<td><b>{}</b></td>" if i == 0 else "<td>{}</td>").format(e(c))
            for i, c in enumerate(r)) + "</tr>" for r in b["rows"])
        return ('<h4>{}</h4><div class="tw"><table><thead><tr>{}</tr></thead>'
                '<tbody>{}</tbody></table></div>').format(e(b["title"]), head, rows)

    if k == "steps":
        items = "".join(
            '<div class="stp"><span class="n">{}</span><div><b>{}</b><span>{}</span></div></div>'.format(
                i + 1, e(a), e(t)) for i, (a, t) in enumerate(b["items"]))
        return '<h4>{}</h4><div class="stps">{}</div>'.format(e(b["title"]), items)

    if k == "sheetpair":
        return sheetpair(b)

    if k == "stat":
        cards = "".join(
            '<div class="st {}"><div class="stn">{}</div><div class="stl">{}</div></div>'.format(
                cls, e(n), e(lab)) for n, lab, cls in b["stats"])
        return ('<h4>{}</h4><div class="statbox"><p class="lead">{}</p>'
                '<div class="sts">{}</div><p class="note">{}</p></div>').format(
            e(b["title"]), e(b["lead"]), cards, e(b["note"]))

    cls = {"tip": "cal tip", "warn": "cal warn", "why": "cal why", "rule": "cal rule"}[k]
    return '<div class="{}"><b>{}</b><p>{}</p></div>'.format(cls, e(b["title"]), e(b["text"]))



def mini_sheet(spec, tone):
    """วาดภาพจำลองหน้าจอสเปรดชีต ใช้ตารางจริงเพื่อให้อ่านออกได้ด้วยโปรแกรมอ่านหน้าจอ"""
    colhead = '<th class="cn"></th>' + "".join(
        '<th class="cl">{}</th>'.format(e(c)) for c in spec["cols"])
    body = ""
    for i, row in enumerate(spec["rows"], start=1):
        cells = ""
        for txt, span, cls in row:
            cells += '<td class="{}"{}>{}</td>'.format(
                cls, ' colspan="{}"'.format(span) if span > 1 else "", e(txt))
        body += '<tr><th class="rn">{}</th>{}</tr>'.format(i, cells)
    tabs = ""
    if spec.get("tabs"):
        tabs = '<div class="tabs">' + "".join(
            '<span class="{}">{}</span>'.format("on" if act else "", e(nm))
            for nm, act in spec["tabs"]) + "</div>"
    return ('<div class="xls {t}"><table>'
            '<thead><tr>{ch}</tr></thead><tbody>{b}</tbody></table>{tabs}</div>').format(
        t=tone, ch=colhead, b=body, tabs=tabs)


def sheetpair(b):
    out = '<h4>{}</h4><p class="sp-lead">{}</p>'.format(e(b["title"]), e(b["lead"]))
    for it in b["items"]:
        out += (
            '<div class="sp">'
            '<div class="sp-h">{label}</div>'
            '<div class="sp-two">'
            '<div class="sp-side"><div class="sp-tag no">แบบที่ใช้ต่อไม่ได้</div>{bad}</div>'
            '<div class="sp-side"><div class="sp-tag ok">แบบที่ควรเป็น</div>{good}</div>'
            '</div>'
            '<div class="sp-why"><b>เกิดอะไรขึ้น</b> {why}</div>'
            '</div>').format(
            label=e(it["label"]), bad=mini_sheet(it["bad"], "bad"),
            good=mini_sheet(it["good"], "good"), why=e(it["why"]))
    return out


def answer_field(code, key, label, rows, ph):
    """ช่องบันทึกคำตอบหนึ่งช่อง เก็บอัตโนมัติในเครื่องของผู้เรียนเอง"""
    return ('<div class="fld"><label>{lb}</label>'
            '<textarea rows="{r}" data-save="ws:{c}:{k}" placeholder="{p}"></textarea></div>').format(
        lb=e(label), r=rows, c=e(code), k=e(key), p=html.escape(ph, quote=True))


def workshop(w):
    tasks = "".join("<li>{}</li>".format(e(t)) for t in w["tasks"])
    flds = "".join(answer_field(w["code"], str(i), "ข้อ {} · {}".format(i + 1, t), 2,
                                "บันทึกสิ่งที่ทำ ผลที่ได้ หรือสูตรที่ใช้")
                   for i, t in enumerate(w["tasks"]))
    flds += answer_field(w["code"], "think", "คิดต่อ · " + w["think"], 3,
                         "ตอบด้วยความคิดของตนเอง ไม่มีคำตอบถูกผิดตายตัว")
    flds += answer_field(w["code"], "stuck", "ติดตรงไหน หรืออยากถามอะไรอาจารย์", 2,
                         "เว้นว่างไว้ได้ถ้าไม่มี")
    builder = L.BUILDERS.get(w.get("builder"), "")
    return ('<div class="ws"><div class="wsh"><span class="wsc">แบบฝึกหัด {}</span>'
            '<b>{}</b><span class="wsm">{}</span></div>'
            '<ul>{}</ul>'
            '<div class="wsp"><b>เกณฑ์ผ่าน</b> {}</div>'
            '<div class="wst"><b>คิดต่อ</b> {}</div>'
            '{}'
            '<div class="wsans"><b>ช่องบันทึกงานของท่าน</b>'
            '<p class="wsnote">พิมพ์แล้วระบบบันทึกให้เองในเครื่องของท่าน '
            'ปิดหน้าเว็บแล้วเปิดใหม่ข้อความยังอยู่ และจะถูกรวมไปในไฟล์ที่ส่งเข้า Classroom ด้วย</p>'
            '{}'
            '<div class="wsbtn"><button type="button" class="btn plain sm" data-savebtn="1">'
            'บันทึกเดี๋ยวนี้</button><span class="wsmsg"></span></div>'
            '</div></div>').format(
        e(w["code"]), e(w["title"]), e(w["mins"]), tasks, e(w["pass"]), e(w["think"]),
        builder, flds)


def module(m):
    body = "".join(block(b) for b in m["blocks"])
    wlist = m.get("workshops") or ([m["workshop"]] if m.get("workshop") else [])
    ws = "".join(workshop(w) for w in wlist)
    return ('<section class="mod" id="{id}">'
            '<div class="modh"><span class="modn">{no}</span>'
            '<div><h3>{title}</h3><span class="modt">{time}</span></div></div>'
            '<p class="modi">{intro}</p>{body}{ws}</section>').format(
        id=m["id"], no=m["no"], title=e(m["title"]), time=e(m["time"]),
        intro=e(m["intro"]), body=body, ws=ws)




nav_extra = "".join('<a href="#{}"><span>&#9998;</span>{}</a>'.format(i, e(t))
                    for i, t in L.LEVEL["navlabs"])

modules_html = "".join(module(m) for m in L.MODULES)
nav_html = '<a href="#start"><span>&#9654;</span>ก่อนเริ่ม</a>' + "".join(
    '<a href="#{}"><span>{}</span>{}</a>'.format(m["id"], m["no"], e(m["title"]))
    for m in L.MODULES)

quiz_html = "".join(
    '<div class="qz" data-a="{a}"><p class="qq"><span class="qm">โมดูล {m}</span>{q}</p>'
    '<div class="qo">{opts}</div>'
    '<div class="qw"><b>เฉลย</b> {why}</div></div>'.format(
        a=q["a"], m=e(q["m"]), q=e(q["q"]),
        opts="".join('<button type="button" data-i="{}">{}</button>'.format(i, e(c))
                     for i, c in enumerate(q["choices"])),
        why=e(q["why"]))
    for q in L.QUIZ)

check_html = "".join(
    '<label><input type="checkbox"><span>{}</span></label>'.format(e(x)) for x in L.CHECKLIST)

HTML = """<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>หน้าเรียน Beginner — รากฐานของข้อมูลที่เชื่อถือได้</title>
%(noindex)s<meta name="description" content="หน้าเรียนระดับ Beginner ของชุดการเรียนรู้ Excel และ Google Sheets สำหรับอ่านเองและทบทวนหลังคาบ">
<style>
:root{--deep:%(deep)s;--blue:%(blue)s;--sky:%(sky)s;--ink:%(ink)s;--amber:%(amber)s;
--coral:%(coral)s;--soft:%(soft)s;--soft2:%(soft2)s;--warn:%(warn)s;--ink70:%(ink70)s;
--muted:%(muted)s;--line:%(line)s}
*{box-sizing:border-box}
body{margin:0;background:#F5F7FA;color:var(--ink);line-height:1.75;
 font-family:"Sarabun","Noto Sans Thai",-apple-system,"Segoe UI",Roboto,sans-serif;font-size:16px}
a{color:var(--blue)}
h3,h4{line-height:1.4;margin:0}
.wrap{max-width:1180px;margin:0 auto;padding:0 20px}

header.hero{background:linear-gradient(160deg,#122240 0%%,#2E74B5 100%%);color:#fff;padding:46px 0 40px}
.eyebrow{font-size:14px;color:#9FC4E8;letter-spacing:.08em}
.hero h1{font-size:32px;font-weight:800;margin:8px 0 0}
.hero .tag{font-size:19px;color:#C3D9EF;margin-top:6px}
.hero .concept{margin-top:20px;background:rgba(255,255,255,.12);border-radius:10px;
 padding:16px 20px;font-size:17px;font-weight:700;max-width:760px}
.hero .meta{margin-top:16px;font-size:14px;color:#9FC4E8}

.layout{display:grid;grid-template-columns:250px minmax(0,1fr);gap:32px;padding:32px 0 60px;align-items:start}
nav.toc{position:sticky;top:20px;background:#fff;border-radius:12px;padding:18px;
 box-shadow:0 1px 3px rgba(16,24,40,.07)}
nav.toc b{display:block;font-size:13px;color:var(--muted);margin-bottom:10px}
nav.toc a{display:flex;gap:9px;align-items:flex-start;padding:8px 6px;border-radius:7px;
 text-decoration:none;color:var(--ink70);font-size:13.5px;line-height:1.45}
nav.toc a:hover{background:var(--soft)}
nav.toc a span{flex:none;width:20px;height:20px;border-radius:50%%;background:var(--blue);color:#fff;
 font-size:11.5px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:2px}

.mywork{margin-top:14px;border-top:1px solid var(--line);padding-top:12px}
.mywork b{display:block;font-size:12.5px;color:var(--muted);margin-bottom:4px}
.mywork span{display:block;font-size:12.5px;color:var(--deep);line-height:1.55;font-weight:700}
.mywork a.dl{display:inline-block;margin-top:8px;padding:6px 12px;background:var(--soft);
 color:var(--deep);font-size:12.5px;font-weight:700;border-radius:7px;text-decoration:none}
.mywork a.dl:hover{background:#D7E7F7}
.mywork.warn span{color:#A83A36}

.mod{background:#fff;border-radius:12px;padding:28px 30px;margin-bottom:20px;
 box-shadow:0 1px 3px rgba(16,24,40,.07);scroll-margin-top:16px}
.modh{display:flex;gap:14px;align-items:flex-start;border-bottom:1px solid var(--line);
 padding-bottom:16px;margin-bottom:16px}
.modn{flex:none;width:38px;height:38px;border-radius:50%%;background:var(--blue);color:#fff;
 font-size:18px;font-weight:800;display:flex;align-items:center;justify-content:center}
.mod h3{font-size:21px;color:var(--deep)}
.modt{font-size:13px;color:var(--muted)}
.modi{font-size:16px;color:var(--ink70);margin:0 0 22px}
.mod h4{font-size:16.5px;color:var(--deep);margin:26px 0 12px}
.mod h4:first-of-type{margin-top:0}

.g4{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
.g4c{background:var(--soft);border-radius:10px;padding:14px 16px}
.g4c b{display:block;color:var(--deep);font-size:15px;margin-bottom:4px}
.g4c span{font-size:13.5px;color:var(--ink70);line-height:1.6}

.two{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}
.box{border-radius:10px;padding:16px 18px}
.box.ok{background:var(--soft)}
.box.ok2{background:var(--soft2)}
.box.no{background:var(--warn)}
.box b{display:block;font-size:15px;margin-bottom:8px}
.box.ok b{color:var(--deep)}.box.ok2 b{color:#8A5A00}.box.no b{color:#A83A36}
.box ul{margin:0;padding-left:20px}.box li{font-size:14px;margin:6px 0;line-height:1.65}

.tw{overflow-x:auto}
table{width:100%%;border-collapse:collapse;font-size:14.5px}
thead th{background:var(--soft);color:var(--deep);text-align:left;padding:11px 14px;
 font-size:13.5px;font-weight:700;border-bottom:2px solid var(--line)}
tbody td{padding:11px 14px;border-bottom:1px solid var(--line);vertical-align:top;line-height:1.65}
tbody tr:last-child td{border-bottom:none}
tbody td b{color:var(--deep)}

.stps{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}
.stp{background:var(--soft);border-radius:10px;padding:14px 16px;display:flex;gap:11px}
.stp .n{flex:none;width:26px;height:26px;border-radius:50%%;background:var(--blue);color:#fff;
 font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center}
.stp b{display:block;font-size:14.5px;color:var(--deep);margin-bottom:3px}
.stp span{font-size:13px;color:var(--ink70);line-height:1.6}

.statbox{background:var(--soft);border-radius:10px;padding:18px 20px}
.statbox .lead{margin:0 0 14px;font-size:15px;font-weight:700;color:var(--deep)}
.sts{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px}
.st{background:#fff;border-radius:9px;padding:14px;text-align:center}
.stn{font-size:34px;font-weight:800;line-height:1.2}
.st.bad .stn{color:var(--coral)}.st.good .stn{color:var(--blue)}.st.warn .stn{color:var(--amber)}
.stl{font-size:12.5px;color:var(--muted);margin-top:4px;line-height:1.5}
.statbox .note{font-size:13.5px;color:var(--ink70);margin:14px 0 0}

.sp-lead{font-size:14.5px;color:var(--ink70);margin:0 0 16px}
.sp{background:#FAFCFE;border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin-bottom:16px}
.sp-h{font-size:15.5px;font-weight:800;color:var(--deep);margin-bottom:14px}
.sp-two{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px}
.sp-side{min-width:0}
.sp-tag{display:inline-block;font-size:12px;font-weight:700;padding:3px 10px;border-radius:5px;margin-bottom:8px}
.sp-tag.no{background:var(--warn);color:#A83A36}
.sp-tag.ok{background:#E7F3E9;color:#1E5C2A}
.sp-why{margin-top:14px;background:var(--soft);border-radius:8px;padding:11px 14px;font-size:14px;line-height:1.7}
.sp-why b{color:var(--deep)}

.xls{border:1px solid #C6D2E0;border-radius:7px;overflow-x:auto;background:#fff}
.xls table{width:100%%;border-collapse:collapse;font-size:12.5px;
 font-family:Calibri,Arial,"Sarabun",sans-serif}
.xls th,.xls td{border:1px solid #DAE2EC;padding:5px 8px;text-align:left;white-space:nowrap}
.xls th.cl,.xls th.cn{background:#EDF1F6;color:#5B6B7F;text-align:center;font-weight:600;font-size:11.5px}
.xls th.rn{background:#EDF1F6;color:#5B6B7F;text-align:center;font-weight:600;font-size:11.5px;width:26px}
.xls td.h{background:#DCE7F3;font-weight:700;color:var(--deep)}
.xls td.n{text-align:right}
.xls td.bad{background:#FBDDDB;color:#8E2F2B;font-weight:700}
.xls.good td.h{background:#DCEEDF;color:#1E5C2A}
.xls .tabs{display:flex;gap:2px;padding:5px 6px;background:#EDF1F6;border-top:1px solid #DAE2EC}
.xls .tabs span{font-size:11.5px;padding:3px 10px;border-radius:4px 4px 0 0;background:#DDE5EE;color:#5B6B7F}
.xls .tabs span.on{background:#fff;color:var(--deep);font-weight:700}
.xls.bad .tabs span.on{color:#8E2F2B}
.labwrap{overflow-x:auto;margin-bottom:16px}
.xls.lab td{padding:0}
.xls.lab td input{width:100%%;min-width:110px;border:none;background:transparent;padding:6px 8px;
 font-family:"SF Mono",Consolas,Menlo,monospace;font-size:12.5px;color:var(--ink);box-sizing:border-box}
.xls.lab td input:focus{outline:2px solid var(--sky);outline-offset:-2px;background:#F7FBFF}
.xls.lab td input.fixed{background:#EAF6EC}
.labbtns{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:6px 0 16px}
.btn{font-family:inherit;font-size:14.5px;font-weight:700;padding:10px 20px;border-radius:8px;
 border:1.5px solid transparent;cursor:pointer}
.btn.go{background:var(--blue);color:#fff}
.btn.go:hover{background:var(--deep)}
.btn.plain{background:#fff;color:var(--deep);border-color:var(--line)}
.btn.plain:hover{border-color:var(--sky)}
.labscore{font-size:14.5px;font-weight:700;color:var(--muted)}
.labscore.ok{color:#1E5C2A}.labscore.no{color:#A83A36}
.labchk{list-style:none;margin:0 0 4px;padding:0}
.labchk li{display:flex;gap:11px;align-items:flex-start;padding:9px 0;font-size:14.5px;
 color:var(--ink70);border-bottom:1px solid #F0F4F8;line-height:1.65}
.labchk li:last-child{border-bottom:none}
.labchk .dot{flex:none;width:19px;height:19px;border-radius:50%%;margin-top:3px;
 background:#E3E9F0;color:#8A97A8;font-size:12px;font-weight:700;display:flex;
 align-items:center;justify-content:center}
.labchk li.pass .dot{background:#5BA860;color:#fff}
.labchk li.pass .dot:after{content:"\\2713"}
.labchk li.fail .dot{background:var(--coral);color:#fff}
.labchk li.fail .dot:after{content:"\\2715"}
.labchk li.pass{color:#1E5C2A}
textarea{width:100%%;box-sizing:border-box;font-family:inherit;font-size:14.5px;line-height:1.7;
 padding:12px 14px;border:1.5px solid var(--line);border-radius:9px;color:var(--ink);resize:vertical}
textarea:focus{outline:none;border-color:var(--sky)}
.fld{margin-bottom:10px}
.fld label{display:block;font-size:13px;color:var(--muted);margin-bottom:4px}
.fld input{width:100%%;box-sizing:border-box;font-family:inherit;font-size:14.5px;padding:9px 12px;
 border:1.5px solid var(--line);border-radius:8px;color:var(--ink);background:#fff}
.fld input:focus{outline:none;border-color:var(--sky)}

.cal{border-radius:0 10px 10px 0;padding:15px 20px;margin:20px 0}
.cal b{display:block;font-size:15px;margin-bottom:5px}
.cal p{margin:0;font-size:14.5px;line-height:1.7}
.cal.tip{background:var(--soft);border-left:4px solid var(--blue)}
.cal.tip b{color:var(--deep)}
.cal.warn{background:var(--warn);border-left:4px solid var(--coral)}
.cal.warn b{color:#A83A36}
.cal.why{background:#F7FAFD;border-left:4px solid var(--sky)}
.cal.why b{color:var(--deep)}
.cal.rule{background:var(--soft2);border-left:4px solid var(--amber)}
.cal.rule b{color:#8A5A00}

.ws{background:#F7FAFD;border:1.5px dashed #B9D2EA;border-radius:12px;padding:20px 22px;margin-top:26px}
.wsh{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap;margin-bottom:10px}
.wsc{background:var(--blue);color:#fff;font-size:12.5px;font-weight:700;padding:3px 11px;border-radius:6px}
.wsh b{font-size:16.5px;color:var(--deep)}
.wsm{font-size:13px;color:var(--muted)}
.ws ul{margin:0;padding-left:20px}.ws li{font-size:14.5px;margin:7px 0}
.wsp,.wst{margin-top:12px;font-size:14px;line-height:1.7;padding:11px 14px;border-radius:8px}
.wsp{background:var(--soft)}.wsp b{color:var(--deep)}
.wst{background:var(--soft2)}.wst b{color:#8A5A00}

.bld{background:#fff;border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-top:18px}
.bldh b{display:block;font-size:15.5px;color:var(--deep)}
.bldh span{display:block;font-size:13px;color:var(--muted);margin:3px 0 12px;line-height:1.6}
.bldbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:10px 0}
.bldwrap{overflow-x:auto}
.bldgrid table{min-width:100%%}
.bldgrid td input,.bldgrid th.ch input{width:100%%;min-width:120px;border:none;background:transparent;
 padding:6px 8px;font-family:"SF Mono",Consolas,Menlo,monospace;font-size:12.5px;box-sizing:border-box}
.bldgrid th.ch{background:#DCE7F3;padding:0}
.bldgrid th.ch input{font-family:inherit;font-weight:700;color:var(--deep);text-align:center}
.bldgrid th.ch input::placeholder{font-weight:400;color:#8FA6BF}
.bldgrid td input:focus,.bldgrid th.ch input:focus{outline:2px solid var(--sky);outline-offset:-2px;background:#F7FBFF}

.pvctl{display:flex;gap:12px;flex-wrap:wrap;margin:8px 0 14px}
.pvctl label{display:flex;flex-direction:column;gap:4px;font-size:12.5px;color:var(--muted);font-weight:700}
.pvctl select{font-family:inherit;font-size:14px;padding:8px 10px;border:1.5px solid var(--line);
 border-radius:8px;background:#fff;color:var(--ink);min-width:150px}
.pvctl select:focus{outline:none;border-color:var(--sky)}
.fld select{width:100%%;box-sizing:border-box;font-family:inherit;font-size:14.5px;padding:9px 12px;
 border:1.5px solid var(--line);border-radius:8px;background:#fff;color:var(--ink)}
.pvempty{padding:16px;font-size:14px;color:var(--muted)}
b.lkh{display:block;font-size:14.5px;color:var(--deep);margin-bottom:8px}
.lknote{margin:0;padding-left:20px}
.lknote li{font-size:14px;color:var(--ink70);margin:7px 0;line-height:1.65}
.fx{background:#fff;border:1px solid var(--line);border-radius:10px;padding:14px 16px;margin-bottom:12px}
.fxq{font-size:15px;color:var(--ink);line-height:1.65;margin-bottom:10px}
.fxn{display:inline-flex;width:22px;height:22px;border-radius:50%%;background:var(--blue);color:#fff;
 font-size:12.5px;font-weight:700;align-items:center;justify-content:center;margin-right:9px;vertical-align:1px}
.fx input{width:100%%;box-sizing:border-box;font-family:"SF Mono",Consolas,Menlo,monospace;font-size:13.5px;
 padding:10px 12px;border:1.5px solid var(--line);border-radius:8px;background:#FAFCFE;color:var(--ink)}
.fx input:focus{outline:none;border-color:var(--sky);background:#fff}
.fxr{font-size:13.5px;line-height:1.65;margin-top:8px;display:none}
.fxr.on{display:block}
.fxr.ok{color:#1E5C2A}.fxr.no{color:#A83A36}

.bhfx{font-family:"SF Mono",Consolas,Menlo,monospace;font-size:13.5px;background:#F5F2FB;
 border-left:3px solid var(--sky);padding:9px 12px;border-radius:0 7px 7px 0;margin-bottom:10px;
 color:var(--ink);overflow-x:auto;white-space:pre}
.fx select{width:100%%;box-sizing:border-box;font-family:inherit;font-size:14px;padding:9px 11px;
 border:1.5px solid var(--line);border-radius:8px;background:#fff;color:var(--ink)}
.fx select:focus{outline:none;border-color:var(--sky)}
.pmout{background:#F7F5FC;border:1px solid var(--line);border-radius:9px;padding:14px 16px;
 font-family:inherit;font-size:14px;line-height:1.75;color:var(--ink70);white-space:pre-wrap;
 margin:0 0 4px;max-height:280px;overflow-y:auto}

.wsans{margin-top:18px;border-top:1.5px dashed #B9D2EA;padding-top:16px}
.wsans>b{display:block;font-size:15.5px;color:var(--deep);margin-bottom:3px}
.wsnote{font-size:13px;color:var(--muted);margin:0 0 14px;line-height:1.65}
.wsans .fld label{font-size:13.5px;color:var(--ink70);line-height:1.6}
.wsans textarea{background:#fff;font-size:14px}
.wsbtn{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:4px}
.btn.sm{font-size:13px;padding:6px 14px;font-weight:700}
.wsmsg{font-size:13px;color:#1E5C2A;font-weight:700}
#saveBar .btn.sm{margin-left:8px;vertical-align:middle}

.qz{background:#fff;border-radius:12px;padding:20px 22px;margin-bottom:14px;
 box-shadow:0 1px 3px rgba(16,24,40,.07)}
.qq{margin:0 0 12px;font-size:16px;font-weight:700;color:var(--ink)}
.qm{background:var(--soft);color:var(--deep);font-size:12px;font-weight:700;padding:2px 9px;
 border-radius:5px;margin-right:9px}
.qo{display:grid;gap:8px}
.qo button{text-align:left;background:#fff;border:1.5px solid var(--line);border-radius:8px;
 padding:11px 14px;font-size:14.5px;font-family:inherit;color:var(--ink70);cursor:pointer;line-height:1.6}
.qo button:hover{border-color:var(--sky)}
.qo button.right{background:#EAF6EC;border-color:#5BA860;color:#1E5C2A;font-weight:700}
.qo button.wrong{background:var(--warn);border-color:var(--coral);color:#A83A36}
.qw{display:none;margin-top:12px;background:var(--soft);border-radius:8px;padding:12px 15px;font-size:14px;line-height:1.7}
.qw b{color:var(--deep)}
.qz.done .qw{display:block}

.chk{background:#fff;border-radius:12px;padding:24px 26px;box-shadow:0 1px 3px rgba(16,24,40,.07)}
.chk label{display:flex;gap:11px;align-items:flex-start;padding:9px 0;font-size:15px;cursor:pointer;
 border-bottom:1px solid #F0F4F8}
.chk label:last-child{border-bottom:none}
.chk input{margin-top:5px;width:17px;height:17px;flex:none;accent-color:%(blue)s}
.chk input:checked + span{color:var(--muted);text-decoration:line-through}

section.blk{padding:30px 0 10px}
.shead{font-size:23px;font-weight:800;color:var(--deep);margin:0 0 6px}
.slead{color:var(--muted);margin:0 0 20px;font-size:15.5px}
footer{background:var(--ink);color:#9FC4E8;padding:30px 0;font-size:13.5px;margin-top:30px;line-height:1.8}
.back{display:inline-block;color:#9FC4E8;text-decoration:none;font-size:14px;margin-bottom:14px}
.back:hover{color:#fff}
.totop{position:fixed;right:22px;bottom:22px;width:44px;height:44px;border-radius:50%%;
 background:var(--deep);color:#fff;text-decoration:none;font-size:20px;display:none;
 align-items:center;justify-content:center;box-shadow:0 2px 8px rgba(16,24,40,.25);z-index:20}
.totop.on{display:flex}
@media(max-width:900px){.layout{grid-template-columns:1fr}nav.toc{position:static}}
@media print{nav.toc,footer{display:none}.mod{box-shadow:none;border:1px solid #ccc}}
</style>
</head>
<body>

%(hero)s

<div class="wrap">
  <div class="layout">
    <nav class="toc">
      <b>หัวข้อในหน้านี้</b>
      %(nav)s
      %(navextra)s
      <a href="#submit"><span>&#8681;</span>ส่งงานเข้า Classroom</a>
      <a href="#quiz"><span>?</span>ตรวจความเข้าใจ</a>
      <div class="mywork" id="myWork">
        <b>งานของฉัน</b>
        <span id="myWorkState">ยังไม่ได้เริ่มพิมพ์</span>
        <a class="dl" href="#submit">ดาวน์โหลดเก็บไว้</a>
      </div>
      <a href="#check"><span>&#10003;</span>รายการตรวจก่อนส่งงาน</a>
    </nav>

    <main>
      %(start)s

      %(modules)s


      %(labs)s

      <section class="mod" id="submit">
        <div class="modh"><span class="modn">&#8681;</span>
          <div><h3>ส่งงานเข้า Google Classroom</h3>
          <span class="modt">ดาวน์โหลดไฟล์งานจากหน้านี้ แล้วอัปโหลดในงานที่อาจารย์มอบหมาย</span></div></div>

        <div class="two">
          <div class="box ok"><b>กรอกก่อนดาวน์โหลด</b>
            <div class="fld"><label>ชื่อ-นามสกุล</label><input id="stName" data-save="me:name" placeholder="สมชาย ใจดี"></div>
            <div class="fld"><label>รหัสนักศึกษา</label><input id="stId" data-save="me:id" placeholder="6512345"></div>
            <div class="fld"><label>สาขา</label><input id="stProg" data-save="me:prog" placeholder="AG"></div>
          </div>
          <div class="box ok2"><b>ไฟล์ที่ได้จะมีอะไรบ้าง</b><ul>%(explist)s</ul></div>
        </div>

        <div class="cal tip" id="saveBar"><b id="saveState">ยังไม่ได้เริ่มพิมพ์</b>
          <p><b style="display:inline">งานถูกเก็บไว้ที่ไหน</b> เก็บอยู่ในหน่วยความจำของเบราว์เซอร์บนเครื่องที่ท่านใช้อยู่ขณะนี้เท่านั้น
          ไม่ได้ส่งออกไปที่ใดและอาจารย์ยังมองไม่เห็น จึงเปิดต่อได้เฉพาะเมื่อกลับมาที่หน้านี้ด้วยเครื่องเดิมและเบราว์เซอร์เดิม
          ถ้าเปลี่ยนเครื่อง เปลี่ยนเบราว์เซอร์ ใช้โหมดไม่ระบุตัวตน หรือล้างประวัติการเข้าชม งานจะหายทั้งหมด
          <b style="display:inline">การส่งงานจริงเกิดขึ้นเมื่อกดดาวน์โหลดไฟล์ด้านล่างแล้วนำไปอัปโหลดใน Classroom เท่านั้น</b>
          <button type="button" class="btn plain sm" id="clearAll">ล้างงานทั้งหมดแล้วเริ่มใหม่</button></p></div>

        <div class="labbtns">
          <button type="button" class="btn go" id="expCsv">ดาวน์โหลดไฟล์งาน (.csv)</button>
          <button type="button" class="btn plain" id="expCopy">คัดลอกสรุปงานเป็นข้อความ</button>
          <span class="labscore" id="expMsg"></span>
        </div>

        <div class="cal why"><b>วิธีส่งใน Google Classroom</b>
          <p>เปิดงานที่อาจารย์มอบหมาย กดเพิ่มหรือสร้าง แล้วเลือกไฟล์ อัปโหลดไฟล์ .csv ที่เพิ่งดาวน์โหลด
          จากนั้นกดส่ง หากอาจารย์ให้ส่งเป็น Google Sheets ให้อัปโหลดไฟล์ขึ้น Drive แล้วเปิดด้วย Google Sheets ก่อนหนึ่งครั้ง</p></div>

        <div class="cal warn"><b>ข้อควรระวังเรื่องไฟล์ .csv ภาษาไทย</b>
          <p>ไฟล์ที่ดาวน์โหลดจากหน้านี้ใส่เครื่องหมายกำกับการเข้ารหัส UTF-8 มาให้แล้ว เปิดใน Excel ได้เลยโดยภาษาไทยไม่เพี้ยน
          แต่ถ้านำไฟล์ .csv จากที่อื่นมาเปิด ให้ใช้วิธีนำเข้าตามโมดูล 3 เสมอ</p></div>
      </section>

      <section class="mod" id="quiz">
        <div class="modh"><span class="modn">?</span>
          <div><h3>ตรวจความเข้าใจด้วยตัวเอง</h3>
          <span class="modt">6 ข้อ · เลือกคำตอบแล้วจะเห็นเฉลยพร้อมเหตุผลทันที</span></div></div>
        <p class="modi">ข้อเหล่านี้ไม่ใช่ข้อสอบและไม่มีการเก็บคะแนน
          ใช้เพื่อตรวจว่าตนเองเข้าใจแต่ละโมดูลแล้วจริงหรือยัง ถ้าตอบผิดข้อไหน ให้กลับไปอ่านโมดูลนั้นอีกครั้ง</p>
        %(quiz)s
      </section>

      <section class="mod" id="check">
        <div class="modh"><span class="modn">&#10003;</span>
          <div><h3>รายการตรวจก่อนส่งงาน</h3>
          <span class="modt">ติ๊กไล่ทีละข้อก่อนกดส่งใน Google Classroom</span></div></div>
        <div class="chk">%(check)s</div>
      </section>

      <section class="mod">
        <div class="modh"><span class="modn">&#9733;</span>
          <div><h3>สามอย่างที่อยากให้จำ</h3>
          <span class="modt">ถ้าลืมอย่างอื่นหมด ขอให้จำสามข้อนี้ไว้</span></div></div>
        <div class="two" style="grid-template-columns:1fr">
          <div class="box ok"><b>1. ดูการจัดชิดของเซลล์ก่อนเสมอ</b>
            <ul><li>ชิดซ้ายคือข้อความ ชิดขวาคือตัวเลข สัญญาณเล็ก ๆ นี้บอกได้ทันทีว่าทำไมสูตรถึงไม่ทำงาน</li></ul></div>
          <div class="box ok2"><b>2. สำรวจให้ครบก่อนลงมือแก้</b>
            <ul><li>การรีบแก้ทีละจุดโดยไม่เห็นภาพรวม มักทำให้แก้ผิดลำดับแล้วต้องย้อนกลับมาทำใหม่</li></ul></div>
          <div class="box ok"><b>3. บันทึกเหตุผลของทุกการตัดสินใจ</b>
            <ul><li>สิ่งที่ทำให้งานวิเคราะห์เชื่อถือได้ ไม่ใช่ตัวเลขที่สวย แต่คือการที่คนอื่นตรวจสอบย้อนกลับได้</li></ul></div>
        </div>
        <div class="cal why"><b>ระดับถัดไป</b>
          <p>ระดับ Intermediate จะสอนการเชื่อมตารางและ PivotTable
          โดยใช้ข้อมูลชุดเดิมที่คุณทำความสะอาดไว้แล้วในคาบนี้ งานวันนี้จึงไม่สูญเปล่า</p></div>
      </section>
    </main>
  </div>
</div>

<footer>
  <div class="wrap">
    <b>ชุดการเรียนรู้ Excel &amp; Google Sheets · ระดับ Beginner</b><br>
    อาจารย์ ดร.เนติยา การะเกตุ · หลักสูตรวิทยาศาสตร์การเกษตร มหาวิทยาลัยมหิดล วิทยาเขตกาญจนบุรี<br>
    ข้อมูลตัวอย่างที่ใช้ในคาบนี้เป็นข้อมูลจำลองเพื่อการเรียนการสอน ไม่ใช่ข้อมูลจริงของหน่วยงานใด
  </div>
</footer>

<a class="totop" href="#" id="toTop" aria-label="กลับขึ้นบนสุด">&#8593;</a>

<script>
(function () {
  // ปุ่มกลับขึ้นบนสุด แสดงเมื่อเลื่อนลงมาแล้ว
  var top = document.getElementById('toTop');
  window.addEventListener('scroll', function () {
    top.classList.toggle('on', window.scrollY > 500);
  });
  top.addEventListener('click', function (ev) {
    ev.preventDefault();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });



  /* ---------- บันทึกงานอัตโนมัติในเครื่องของผู้เรียน ---------- */
  var STORE = '%(storekey)s';
  var canStore = (function () {
    try { localStorage.setItem('__t', '1'); localStorage.removeItem('__t'); return true; }
    catch (err) { return false; }
  })();
  var memStore = {};

  function readStore() {
    if (!canStore) return memStore;
    try { return JSON.parse(localStorage.getItem(STORE) || '{}'); } catch (err) { return {}; }
  }
  function fields() { return [].slice.call(document.querySelectorAll('[data-save]')); }

  function restore() {
    var o = readStore(), n = 0;
    fields().forEach(function (el) {
      var k = el.getAttribute('data-save');
      if (Object.prototype.hasOwnProperty.call(o, k)) { el.value = o[k]; n++; }
    });
    return n;
  }
  function persist() {
    var o = {};
    fields().forEach(function (el) { o[el.getAttribute('data-save')] = el.value; });
    if (!canStore) { memStore = o; return false; }
    try { localStorage.setItem(STORE, JSON.stringify(o)); return true; }
    catch (err) { return false; }
  }

  function stamp() {
    var d = new Date();
    return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2) + ' น.';
  }
  function announce(msg, bad) {
    var s = document.getElementById('saveState');
    if (s) { s.textContent = msg; }
    var w = document.getElementById('myWorkState');
    if (w) { w.textContent = msg; }
    var box = document.getElementById('myWork');
    if (box) { box.classList.toggle('warn', !!bad); }
  }

  var saveTimer = null;
  function scheduleSave(btnMsg) {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      var ok = persist();
      announce(ok ? 'เก็บไว้ในเบราว์เซอร์เครื่องนี้แล้ว เวลา ' + stamp()
                  : 'เบราว์เซอร์นี้เก็บงานให้ไม่ได้ ต้องดาวน์โหลดไฟล์ก่อนปิดหน้า', !ok);
      if (btnMsg) {
        btnMsg.textContent = ok ? 'เก็บไว้ในเบราว์เซอร์เครื่องนี้แล้ว ' + stamp()
                                : 'เก็บไม่ได้ กรุณาดาวน์โหลดไฟล์งาน';
      }
    }, btnMsg ? 0 : 700);
  }

  var restored = restore();
  if (!canStore) {
    announce('เบราว์เซอร์นี้ไม่อนุญาตให้เก็บงาน ข้อความจะหายเมื่อปิดหน้า ต้องดาวน์โหลดไฟล์ก่อนปิด', true);
  } else if (restored > 0) {
    announce('เปิดงานที่เก็บไว้ในเบราว์เซอร์เครื่องนี้ขึ้นมาแล้ว');
  }

  document.addEventListener('input', function (ev) {
    if (ev.target && ev.target.hasAttribute && ev.target.hasAttribute('data-save')) scheduleSave(null);
  });

  [].forEach.call(document.querySelectorAll('[data-savebtn]'), function (b) {
    b.addEventListener('click', function () {
      scheduleSave(b.parentNode.querySelector('.wsmsg'));
    });
  });

  var clr = document.getElementById('clearAll');
  if (clr) clr.addEventListener('click', function () {
    if (!confirm('ลบงานทั้งหมดที่บันทึกไว้ในเครื่องนี้ ทั้งคำตอบแบบฝึกหัดและตารางที่แก้ไว้ ยืนยันหรือไม่')) return;
    try { localStorage.removeItem(STORE); } catch (err) { }
    memStore = {};
    location.reload();
  });

  /* ---------- แบบตรวจความเข้าใจ นับคะแนนไว้ใช้ตอนส่งออก ---------- */
  var quizRight = 0, quizDone = 0;

  /* ทะเบียนงานในหน้านี้ ห้องฝึกแต่ละชุดลงทะเบียนตัวเองไว้ ปุ่มส่งออกจึงเก็บครบโดยไม่ต้องรู้จักกันเอง
     รูปแบบที่ลงทะเบียน { title, run() -> { score, items[], cols[], rows[][], notes[[ป้าย,ค่า]] } } */
  var WORK = [];


  /* ---------- กระดานสร้างตารางของแบบฝึกหัด 1.1 ---------- */
  var bldBox = document.getElementById('bldGrid');
  if (bldBox) {
    var store = document.getElementById('bldStore');
    var COLS0 = 5, ROWS0 = 10;

    function bldModel() {
      // อ่านจากช่องที่บันทึกไว้ก่อน ถ้าไม่มีจึงเริ่มด้วยตารางเปล่า
      if (store.value) {
        try {
          var o = JSON.parse(store.value);
          if (o && o.cols && o.rows) return o;
        } catch (err) { }
      }
      var c = [], r = [], i, j;
      for (i = 0; i < COLS0; i++) { c.push(''); }
      for (i = 0; i < ROWS0; i++) { var rr = []; for (j = 0; j < COLS0; j++) { rr.push(''); } r.push(rr); }
      return { cols: c, rows: r };
    }

    function colName(i) {
      var s2 = '';
      i++;
      while (i > 0) { var m = (i - 1) %% 26; s2 = String.fromCharCode(65 + m) + s2; i = (i - m - 1) / 26; }
      return s2;
    }

    function bldDraw(m) {
      var h = '<table><thead><tr><th class="cn"></th>';
      m.cols.forEach(function (c, j) { h += '<th class="cl">' + colName(j) + '</th>'; });
      h += '</tr><tr><th class="rn">1</th>';
      m.cols.forEach(function (c, j) {
        h += '<th class="ch"><input value="' + c.replace(/"/g, '&quot;')
          + '" placeholder="ตั้งชื่อคอลัมน์" data-h="' + j + '"></th>';
      });
      h += '</tr></thead><tbody>';
      m.rows.forEach(function (row, i) {
        h += '<tr><th class="rn">' + (i + 2) + '</th>';
        row.forEach(function (v, j) {
          h += '<td><input value="' + String(v).replace(/"/g, '&quot;')
            + '" data-br="' + i + '" data-bc="' + j + '"></td>';
        });
        h += '</tr>';
      });
      bldBox.innerHTML = h + '</tbody></table>';
    }

    function bldRead() {
      var cols = [].slice.call(bldBox.querySelectorAll('th.ch input')).map(function (i) { return i.value; });
      var rows = [];
      [].forEach.call(bldBox.querySelectorAll('tbody tr'), function (tr) {
        rows.push([].slice.call(tr.querySelectorAll('input')).map(function (i) { return i.value; }));
      });
      return { cols: cols, rows: rows };
    }
    function bldSync() {
      store.value = JSON.stringify(bldRead());
      scheduleSave(null);
    }

    bldDraw(bldModel());
    bldBox.addEventListener('input', bldSync);

    [].forEach.call(document.querySelectorAll('[data-bld]'), function (b) {
      b.addEventListener('click', function () {
        var m = bldRead(), what = b.getAttribute('data-bld'), i, j;
        if (what === 'row') {
          for (i = 0; i < 4; i++) {
            var rr = []; for (j = 0; j < m.cols.length; j++) { rr.push(''); } m.rows.push(rr);
          }
        } else if (what === 'col') {
          m.cols.push(''); m.rows.forEach(function (r) { r.push(''); });
        } else if (what === 'delcol') {
          if (m.cols.length <= 2) { return; }
          m.cols.pop(); m.rows.forEach(function (r) { r.pop(); });
        } else if (what === 'trim') {
          while (m.rows.length > 1 && m.rows[m.rows.length - 1].every(function (v) { return v.trim() === ''; })) {
            m.rows.pop();
          }
        }
        bldDraw(m); bldSync();
      });
    });

    /* กฎตรวจโครงสร้าง อิงกฎสามข้อของโมดูล 1 */
    var DATE_RE = /^\\d{4}-\\d{2}-\\d{2}$/;
    var NUM_RE = /^-?\\d+(\\.\\d+)?$/;

    function bldRules(m) {
      var cols = m.cols.map(function (c) { return c.trim(); });
      var body = m.rows.map(function (r) { return r.map(function (v) { return String(v).trim(); }); });
      var isFull = function (r) { return r.some(function (v) { return v !== ''; }); };
      var full = body.filter(isFull);

      var named = cols.length > 0 && cols.every(function (c) { return c !== ''; });
      if (named) {
        var seen = {};
        cols.forEach(function (c) { if (seen[c]) { named = false; } seen[c] = 1; });
      }

      var size = cols.length >= 4 && full.length >= 8;

      var first = -1, last = -1;
      body.forEach(function (r, i) { if (isFull(r)) { if (first < 0) { first = i; } last = i; } });
      var gap = true;
      for (var i2 = first; i2 >= 0 && i2 <= last; i2++) { if (!isFull(body[i2])) { gap = false; } }

      // คอลัมน์ใดเป็นวันที่ ต้องเป็นวันที่จริงทุกเซลล์ที่กรอก
      var dateCol = [];
      cols.forEach(function (c, j) {
        var vals = full.map(function (r) { return r[j] || ''; }).filter(function (v) { return v !== ''; });
        if (vals.length && vals.every(function (v) { return DATE_RE.test(v); })) { dateCol.push(j); }
      });
      var date = dateCol.length > 0;

      // คอลัมน์ที่ตั้งใจเก็บตัวเลข ต้องไม่มีหน่วยหรือข้อความปน
      var pure = true;
      cols.forEach(function (c, j) {
        if (dateCol.indexOf(j) >= 0) { return; }
        var vals = full.map(function (r) { return r[j] || ''; }).filter(function (v) { return v !== ''; });
        if (!vals.length) { return; }
        var digitish = vals.filter(function (v) { return /^-?\\d/.test(v); }).length;
        if (digitish / vals.length >= 0.6 && !vals.every(function (v) { return NUM_RE.test(v); })) {
          pure = false;
        }
      });
      if (!full.length) { pure = false; date = false; }

      return { named: named, size: size, gap: gap && first >= 0, date: date, pure: pure };
    }

    function bldRun() {
      var m = bldRead(), res = bldRules(m), pass = 0, out = [];
      [].forEach.call(document.querySelectorAll('#bldChk li'), function (li) {
        var k = li.getAttribute('data-k'), ok = !!res[k];
        li.classList.remove('pass', 'fail');
        li.classList.add(ok ? 'pass' : 'fail');
        if (ok) { pass++; }
        out.push({ text: li.textContent.trim(), ok: ok });
      });
      var sc = document.getElementById('bldScore');
      sc.textContent = 'ผ่าน ' + pass + ' จาก 5 ข้อ'
        + (pass === 5 ? ' · ตารางนี้พร้อมนำไปวิเคราะห์ต่อได้จริง' : ' · ดูรายการที่ยังเป็นกากบาทด้านล่าง');
      sc.className = 'labscore ' + (pass === 5 ? 'ok' : 'no');
      return { pass: pass, items: out, model: m };
    }
    document.getElementById('bldCheck').addEventListener('click', bldRun);
    WORK.push({
      title: 'ตารางที่สร้างในแบบฝึกหัด 1.1',
      run: function () {
        var b = bldRun();
        return { score: 'ตรวจโครงสร้างผ่าน ' + b.pass + ' จาก 5 ข้อ', items: b.items,
                 cols: b.model.cols, rows: b.model.rows };
      }
    });
  }

  /* ---------- ห้องฝึกทำความสะอาดข้อมูลในหน้าเว็บ ---------- */
  var labTable = document.getElementById('labGrid');   // ต้องเจาะจงด้วย id เพราะกระดานสร้างตารางใช้คลาสเดียวกัน
  if (labTable) {
    var cells = [].slice.call(labTable.querySelectorAll('input'));
    // ใช้ defaultValue เพราะค่าปัจจุบันอาจถูกแทนด้วยงานที่บันทึกไว้ไปแล้ว
    var seed = cells.map(function (i) { return i.defaultValue; });
    cells.forEach(function (i) { i.setAttribute('data-save', 'lab:' + i.getAttribute('data-r') + ':' + i.getAttribute('data-c')); });
    restore();
    var nRow = 0, nCol = 0;
    cells.forEach(function (i) {
      nRow = Math.max(nRow, +i.getAttribute('data-r') + 1);
      nCol = Math.max(nCol, +i.getAttribute('data-c') + 1);
    });
    var heads = [].slice.call(labTable.querySelectorAll('thead th')).slice(1)
      .map(function (t) { return t.textContent; });

    function grid() {
      var g = [];
      for (var r = 0; r < nRow; r++) { g.push(new Array(nCol).fill('')); }
      cells.forEach(function (i) { g[+i.getAttribute('data-r')][+i.getAttribute('data-c')] = i.value; });
      return g;
    }
    function filled(row) { return row.some(function (v) { return v.trim() !== ''; }); }

    /* คืนค่า true เมื่อแก้ปัญหานั้นเรียบร้อยแล้ว */
    var rules = {
      space: function (g) {
        return g.every(function (row) { return row.every(function (v) { return v === v.trim(); }); });
      },
      dup: function (g) {
        var seen = {}, dup = false;
        g.forEach(function (row) {
          if (!filled(row)) return;
          var k = row.map(function (v) { return v.trim(); }).join('\\u0001');
          if (seen[k]) dup = true; else seen[k] = 1;
        });
        return !dup;
      },
      sentinel: function (g) {
        return g.every(function (row) {
          var v = (row[2] || '').trim();
          return v !== '-999' && v !== '0' && v !== '-999.0';
        });
      },
      textnum: function (g) {
        return g.every(function (row) {
          var v = (row[2] || '').trim();
          return v === '' || /^-?\\d+(\\.\\d+)?$/.test(v);
        });
      },
      zero: function (g) {
        return g.every(function (row) {
          var v = (row[3] || '').trim();
          return v === '' || (/^-?\\d+(\\.\\d+)?$/.test(v) && parseFloat(v) > 0);
        });
      }
    };

    var lastResult = null;
    function runCheck() {
      var g = grid(), pass = 0, out = [];
      [].forEach.call(document.querySelectorAll('#labChk li'), function (li) {
        var k = li.getAttribute('data-k');
        var ok = rules[k] ? rules[k](g) : false;
        li.classList.remove('pass', 'fail');
        li.classList.add(ok ? 'pass' : 'fail');
        if (ok) pass++;
        out.push({ k: k, text: li.textContent.trim(), ok: ok });
      });
      var sc = document.getElementById('labScore');
      sc.textContent = 'แก้ได้ ' + pass + ' จาก 5 ข้อ'
        + (pass === 5 ? ' · ตารางนี้พร้อมนำไปคำนวณแล้ว' : ' · ลองอ่านรายการที่ยังเป็นกากบาทอีกครั้ง');
      sc.className = 'labscore ' + (pass === 5 ? 'ok' : 'no');
      lastResult = { pass: pass, items: out, grid: g };
      return lastResult;
    }

    document.getElementById('labCheck').addEventListener('click', runCheck);
    document.getElementById('labReset').addEventListener('click', function () {
      cells.forEach(function (i, k) { i.value = seed[k]; });
      [].forEach.call(document.querySelectorAll('#labChk li'), function (li) {
        li.classList.remove('pass', 'fail');
      });
      var sc = document.getElementById('labScore');
      sc.textContent = ''; sc.className = 'labscore';
      lastResult = null;
      scheduleSave(null);
    });

    WORK.push({
      title: 'ห้องฝึกทำความสะอาดข้อมูล',
      run: function () {
        var r = runCheck();   // ตรวจสดทุกครั้ง กันการส่งออกผลเก่าที่ไม่ตรงกับตารางปัจจุบัน
        return { score: 'แก้ได้ ' + r.pass + ' จาก 5 ข้อ', items: r.items,
                 cols: heads, rows: r.grid.filter(filled),
                 notes: [['บันทึกการตัดสินใจ', document.getElementById('labNote').value]] };
      }
    });
  }

  [].forEach.call(document.querySelectorAll('.qz'), function (qz) {
    var correct = parseInt(qz.getAttribute('data-a'), 10);
    var btns = [].slice.call(qz.querySelectorAll('.qo button'));
    btns.forEach(function (b) {
      b.addEventListener('click', function () {
        if (qz.classList.contains('done')) return;
        var picked = parseInt(b.getAttribute('data-i'), 10);
        btns.forEach(function (x) {
          var i = parseInt(x.getAttribute('data-i'), 10);
          if (i === correct) x.classList.add('right');
          else if (i === picked) x.classList.add('wrong');
          x.disabled = true;
        });
        qz.classList.add('done');
        quizDone++;
        if (picked === correct) quizRight++;
      });
    });
  });


  /* ================= เครื่องมือตรวจสูตร ใช้ร่วมกันหลายห้องฝึก ================= */
  function fxNorm(v) { return String(v || '').replace(/\\s+/g, '').toUpperCase(); }
  function fxHasExactLookup(v) {
    var u = fxNorm(v);
    if (u.indexOf('XLOOKUP') >= 0) { return true; }
    if (u.indexOf('VLOOKUP') >= 0) { return /FALSE/.test(u) || /,0[),]/.test(u); }
    if (u.indexOf('INDEX') >= 0 && u.indexOf('MATCH') >= 0) { return /MATCH\([^)]*,0\)/.test(u); }
    return false;
  }
  function fxHandlesMissing(v) {
    var u = fxNorm(v);
    if (/IFERROR\(/.test(u)) {
      // ต้องคืนเป็นข้อความ ไม่ใช่ศูนย์หรือค่าว่าง
      return !/,0\)$/.test(u) && !/,""\)$/.test(u) && !/,\)$/.test(u);
    }
    // XLOOKUP มีอาร์กิวเมนต์ที่สี่ไว้ใส่ค่าที่จะคืนเมื่อไม่พบ
    if (u.indexOf('XLOOKUP') >= 0) {
      var inner = u.slice(u.indexOf('XLOOKUP') + 8);
      return (inner.match(/,/g) || []).length >= 3 && /"[^"]+"/.test(String(v));
    }
    return false;
  }
  function commaCount(v) { return (fxNorm(v).match(/,/g) || []).length; }

  /* ================= กระดาน PivotTable จำลอง ================= */
  var pvBox = document.getElementById('pvBoard');
  if (pvBox) {
    var PV = JSON.parse(document.getElementById('pvData').textContent);
    var pvSeen = document.getElementById('pvSeen');
    var elRow = document.getElementById('pvRow'), elCol = document.getElementById('pvCol');
    var elVal = document.getElementById('pvVal'), elAgg = document.getElementById('pvAgg');
    var elFf = document.getElementById('pvFf'), elFv = document.getElementById('pvFv');
    var elOut = document.getElementById('pvOut');

    function uniq(field) {
      var seen = {}, out = [];
      PV.rows.forEach(function (r) { if (!seen[r[field]]) { seen[r[field]] = 1; out.push(r[field]); } });
      return out;   // เรียงตามลำดับที่พบในข้อมูล เดือนจึงเรียงถูกโดยไม่ต้องแปลงชื่อ
    }
    function agg(list, how, grand) {
      if (how === 'count') { return list.length; }
      var sum = list.reduce(function (a, b) { return a + b; }, 0);
      if (how === 'sum') { return Math.round(sum * 10) / 10; }
      if (how === 'pct') { return grand ? (Math.round(sum / grand * 1000) / 10) + '%%' : ''; }
      return list.length ? Math.round(sum / list.length * 10) / 10 : '';
    }

    /* เติมรายการค่าที่กรองได้ตามฟิลด์ที่เลือก */
    function fillFilterValues() {
      var f = elFf.value;
      elFv.innerHTML = '<option value="">ทั้งหมด</option>';
      if (!f) { elFv.disabled = true; return; }
      elFv.disabled = false;
      uniq(f).forEach(function (v) {
        var o = document.createElement('option');
        o.value = v; o.textContent = v;
        elFv.appendChild(o);
      });
    }

    /* แถวที่เหลือหลังผ่านพื้นที่ตัวกรอง */
    function filtered() {
      var f = elFf.value, v = elFv.value;
      if (!f || !v) { return PV.rows; }
      return PV.rows.filter(function (r) { return String(r[f]) === v; });
    }
    function pvRender() {
      var rf = elRow.value, cf = elCol.value, vf = elVal.value, how = elAgg.value;
      if (!rf && !cf) {
        elOut.innerHTML = '<div class="pvempty">เลือกอย่างน้อยหนึ่งพื้นที่ก่อน '
          + 'ลองเริ่มจากใส่ สถานี ไว้ที่แถว แล้วใส่ ค่าPM25 ไว้ที่ค่า</div>';
        return null;
      }
      if (!vf && how !== 'count') {
        elOut.innerHTML = '<div class="pvempty">ยังไม่ได้เลือกฟิลด์ในพื้นที่ค่า '
          + 'หรือจะเปลี่ยนวิธีสรุปเป็นจำนวนนับก็ได้</div>';
        return null;
      }
      var src = filtered();
      if (!src.length) {
        elOut.innerHTML = '<div class="pvempty">ตัวกรองที่เลือกไม่มีข้อมูลเหลืออยู่เลย '
          + 'ข้อนี้เป็นสิ่งที่ต้องสังเกต เพราะตารางว่างไม่ได้แปลว่าค่าเป็นศูนย์</div>';
        return null;
      }
      var rows = [], cols = [];
      (function () {
        var s1 = {}, s2 = {};
        src.forEach(function (r) {
          if (rf && !s1[r[rf]]) { s1[r[rf]] = 1; rows.push(r[rf]); }
          if (cf && !s2[r[cf]]) { s2[r[cf]] = 1; cols.push(r[cf]); }
        });
      })();
      if (!rf) { rows = ['']; }
      if (!cf) { cols = ['']; }
      var grand = 0;
      if (vf) { src.forEach(function (r) { grand += Number(r[vf]) || 0; }); }
      var cell = {};
      src.forEach(function (r) {
        var k = (rf ? r[rf] : '') + '' + (cf ? r[cf] : '');
        (cell[k] = cell[k] || []).push(vf ? Number(r[vf]) : 1);
      });
      var h = '<table><thead><tr><th class="cl">' + (rf || '') + '</th>';
      cols.forEach(function (c) { h += '<th class="cl">' + (c || (vf || 'ค่า')) + '</th>'; });
      if (cf) { h += '<th class="cl">รวมทุกคอลัมน์</th>'; }
      h += '</tr></thead><tbody>';
      rows.forEach(function (r) {
        h += '<tr><td class="h">' + r + '</td>';
        var all = [];
        cols.forEach(function (c) {
          var v = cell[r + '' + c] || [];
          all = all.concat(v);
          h += '<td class="n">' + (v.length ? agg(v, how, grand) : '') + '</td>';
        });
        if (cf) { h += '<td class="n">' + (all.length ? agg(all, how, grand) : '') + '</td>'; }
        h += '</tr>';
      });
      elOut.innerHTML = h + '</tbody></table>';

      // จำไว้ว่าเคยตั้งค่าถูกต้องแล้วหรือยัง เพราะต้องสลับไปดูจำนวนนับด้วยจึงจะตอบครบ
      var flags = (pvSeen.value || '').split(',');
      function mark(f) { if (flags.indexOf(f) < 0) { flags.push(f); pvSeen.value = flags.join(','); scheduleSave(null); } }
      if (rf === 'สถานี' && vf === 'ค่าPM25' && how === 'avg') { mark('avg'); }
      if (rf === 'สถานี' && how === 'count') { mark('count'); }
      if (elFf.value && elFv.value) { mark('filter'); }
      return true;
    }
    [elRow, elCol, elVal, elAgg, elFv].forEach(function (el) { el.addEventListener('change', pvRender); });
    elFf.addEventListener('change', function () { fillFilterValues(); pvRender(); });
    fillFilterValues();
    pvRender();

    function pvRun() {
      var flags = (pvSeen.value || '').split(',');
      var a1 = (document.getElementById('pvA1').value || '').replace(/\\s/g, '');
      var want = PV.topAvg.replace(/\\s/g, '');
      var a4 = (document.getElementById('pvA4').value || '').replace(/\\s/g, '');
      var wantJan = PV.topJan.replace(/\\s/g, '');
      function same(got, exp) {
        return got !== '' && (got === exp || (exp.indexOf(got) >= 0 && got.length >= 6));
      }
      var res = {
        cfg: flags.indexOf('avg') >= 0,
        a1: same(a1, want),
        a2: parseInt(document.getElementById('pvA2').value, 10) === PV.topN,
        a4: flags.indexOf('filter') >= 0 && same(a4, wantJan),
        a3: document.getElementById('pvA3').value === 'b'
      };
      var pass = 0, items = [];
      [].forEach.call(document.querySelectorAll('#pvChk li'), function (li) {
        var ok = !!res[li.getAttribute('data-k')];
        li.classList.remove('pass', 'fail');
        li.classList.add(ok ? 'pass' : 'fail');
        if (ok) { pass++; }
        items.push({ text: li.textContent.trim(), ok: ok });
      });
      var sc = document.getElementById('pvScore');
      var n = document.querySelectorAll('#pvChk li').length;
      sc.textContent = 'ผ่าน ' + pass + ' จาก ' + n + ' ข้อ'
        + (pass === n ? ' · อ่านตารางสรุปเป็นแล้ว' : ' · ลองเปลี่ยนวิธีสรุปหรือใช้ตัวกรองดูอีกครั้ง');
      sc.className = 'labscore ' + (pass === n ? 'ok' : 'no');
      return { pass: pass, items: items };
    }
    document.getElementById('pvCheck').addEventListener('click', pvRun);
    WORK.push({
      title: 'กระดาน PivotTable แบบฝึกหัด 2.3',
      run: function () {
        var r = pvRun();
        return { score: 'ผ่าน ' + r.pass + ' จาก ' + r.items.length + ' ข้อ', items: r.items,
                 notes: [['การตั้งค่ากระดาน', 'แถว ' + (elRow.value || 'ไม่ใช้')
                          + ' · คอลัมน์ ' + (elCol.value || 'ไม่ใช้')
                          + ' · ค่า ' + (elVal.value || 'ไม่ใช้')
                          + ' · วิธีสรุป ' + elAgg.options[elAgg.selectedIndex].text
                          + ' · ตัวกรอง ' + (elFf.value ? elFf.value + ' = ' + (elFv.value || 'ทั้งหมด') : 'ไม่ใช้')],
                         ['สถานีที่ตอบ', document.getElementById('pvA1').value],
                         ['จำนวนเดือนที่ตอบ', document.getElementById('pvA2').value],
                         ['สถานีที่ตอบเมื่อกรองเฉพาะมกราคม', document.getElementById('pvA4').value],
                         ['ข้อสรุปที่เลือก', document.getElementById('pvA3').selectedIndex > 0
                          ? document.getElementById('pvA3').options[document.getElementById('pvA3').selectedIndex].text
                          : '']] };
      }
    });
  }

  /* ================= ห้องฝึกเชื่อมตาราง ================= */
  var lkGrid = document.getElementById('lkGrid');
  if (lkGrid) {
    var LK = JSON.parse(document.getElementById('lkData').textContent);
    var lkIn = [].slice.call(lkGrid.querySelectorAll('input[data-lk]'));
    lkIn.forEach(function (el) { el.setAttribute('data-save', 'lk:' + el.getAttribute('data-lk')); });
    restore();

    function lkRules() {
      var vals = lkIn.map(function (el) { return el.value.trim(); });
      var match = true, space = true, miss = true, filledAny = false;
      LK.main.forEach(function (m, i) {
        var code = String(m.code).trim();
        var want = LK.ref[code];
        var got = vals[i];
        if (got !== '') { filledAny = true; }
        if (want) {
          if (got !== want) {
            match = false;
            if (String(m.code) !== code) { space = false; }
          }
        } else {
          // แถวที่หาไม่พบ ต้องเขียนบอกสาเหตุ ไม่ใช่ปล่อยว่างหรือใส่ศูนย์
          if (got === '' || got === '0' || /^\d+$/.test(got)) { miss = false; }
        }
      });
      if (!filledAny) { match = false; space = false; miss = false; }
      var fx = document.getElementById('lkFx').value;
      return {
        match: match, space: space, miss: miss,
        count: parseInt(document.getElementById('lkCount').value, 10) === LK.missing,
        fx: fx.trim() !== '' && fxHasExactLookup(fx) && fxHandlesMissing(fx)
      };
    }
    function lkRun() {
      var res = lkRules(), pass = 0, items = [];
      [].forEach.call(document.querySelectorAll('#lkChk li'), function (li) {
        var ok = !!res[li.getAttribute('data-k')];
        li.classList.remove('pass', 'fail');
        li.classList.add(ok ? 'pass' : 'fail');
        if (ok) { pass++; }
        items.push({ text: li.textContent.trim(), ok: ok });
      });
      var sc = document.getElementById('lkScore');
      sc.textContent = 'ผ่าน ' + pass + ' จาก 5 ข้อ'
        + (pass === 5 ? ' · เชื่อมตารางได้ถูกต้องและตรวจสอบย้อนกลับได้' : ' · ดูรายการที่ยังเป็นกากบาท');
      sc.className = 'labscore ' + (pass === 5 ? 'ok' : 'no');
      return { pass: pass, items: items, vals: lkIn.map(function (e2) { return e2.value; }) };
    }
    document.getElementById('lkCheck').addEventListener('click', lkRun);
    document.getElementById('lkReset').addEventListener('click', function () {
      lkIn.forEach(function (el) { el.value = ''; });
      [].forEach.call(document.querySelectorAll('#lkChk li'), function (li) {
        li.classList.remove('pass', 'fail');
      });
      var sc = document.getElementById('lkScore'); sc.textContent = ''; sc.className = 'labscore';
      scheduleSave(null);
    });
    WORK.push({
      title: 'ห้องฝึกเชื่อมตาราง แบบฝึกหัด 2.2',
      run: function () {
        var r = lkRun();
        return { score: 'ผ่าน ' + r.pass + ' จาก 5 ข้อ', items: r.items,
                 cols: ['เลขที่', 'รหัสจังหวัด', 'ชื่อจังหวัดที่ตอบ'],
                 rows: LK.main.map(function (m, i) { return [m.id, m.code, r.vals[i]]; }),
                 notes: [['จำนวนแถวที่หาไม่พบที่ตอบ', document.getElementById('lkCount').value],
                         ['สูตรที่เขียน', document.getElementById('lkFx').value]] };
      }
    });
  }

  /* ================= ห้องฝึกเขียนสูตร ================= */
  var fxWrap = document.getElementById('fx');
  if (fxWrap) {
    var FXR = {
      f1: function (v) {
        if (fxNorm(v).indexOf('SUMIFS') < 0) { return 'ยังไม่ได้ใช้ SUMIFS ซึ่งเป็นฟังก์ชันที่รวมค่าตามเงื่อนไขหลายข้อ'; }
        if (commaCount(v) < 4) { return 'ใช้ SUMIFS ถูกแล้ว แต่ยังใส่เงื่อนไขไม่ครบสองข้อ ต้องมีทั้งตำบลและปี'; }
        return '';
      },
      f2: function (v) {
        if (fxNorm(v).indexOf('COUNTIFS') < 0) { return 'ยังไม่ได้ใช้ COUNTIFS ระวังสับสนกับ COUNTIF ที่รับเงื่อนไขได้ข้อเดียว'; }
        if (commaCount(v) < 3) { return 'ใช้ COUNTIFS ถูกแล้ว แต่ยังใส่เงื่อนไขไม่ครบสองข้อ'; }
        return '';
      },
      f3: function (v) {
        if (!/VLOOKUP|XLOOKUP|INDEX/i.test(v)) { return 'ยังไม่ได้ใช้ฟังก์ชันค้นหา ลองเลือกจากสามตัวในตารางเปรียบเทียบ'; }
        if (!fxHasExactLookup(v)) { return 'ใช้ฟังก์ชันค้นหาแล้ว แต่ยังไม่ได้ระบุการจับคู่แบบตรงกันเป๊ะ นี่คือกับดักที่ทำให้ได้คำตอบผิดโดยไม่มีการแจ้งเตือน'; }
        return '';
      },
      f4: function (v) {
        if (v.indexOf('*') < 0) { return 'โจทย์ให้คูณ จึงต้องมีเครื่องหมายคูณในสูตร'; }
        if (!/\$[A-Za-z]+\$?\d+|\$?[A-Za-z]+\$\d+/.test(v)) { return 'ยังไม่ได้ตรึงเซลล์ราคาต่อหน่วยด้วยเครื่องหมายดอลลาร์ ลากสูตรลงแล้วจะเลื่อนไปอ้างเซลล์ว่าง'; }
        if (!/\$[A-Za-z]+\$\d+/.test(v)) { return 'ตรึงไว้ด้านเดียว ข้อนี้ต้องตรึงทั้งคอลัมน์และแถว จึงต้องมีดอลลาร์สองตัว'; }
        return '';
      },
      f5: function (v) {
        if (!/IFERROR/i.test(v) && !/XLOOKUP/i.test(v)) { return 'ยังไม่ได้จัดการกรณีหาไม่พบ ใช้ IFERROR ครอบ หรือใช้อาร์กิวเมนต์ที่สี่ของ XLOOKUP'; }
        if (!fxHasExactLookup(v)) { return 'อย่าลืมว่าสูตรข้างในต้องเป็นการค้นหาแบบตรงกันเป๊ะด้วย'; }
        if (!fxHandlesMissing(v)) { return 'จัดการกรณีหาไม่พบแล้ว แต่ต้องคืนเป็นข้อความที่บอกสาเหตุ ไม่ใช่ศูนย์หรือค่าว่าง มิฉะนั้นจะกลบปัญหาแทนที่จะรายงาน'; }
        return '';
      }
    };
    function fxRun() {
      var pass = 0, items = [], notes = [];
      [].forEach.call(fxWrap.querySelectorAll('input[data-fx]'), function (el) {
        var k = el.getAttribute('data-fx');
        var q = el.parentNode.querySelector('.fxq').textContent.trim();
        var out = fxWrap.querySelector('[data-fxr="' + k + '"]');
        var msg = el.value.trim() === '' ? 'ยังไม่ได้เขียนสูตรข้อนี้' : FXR[k](el.value);
        out.className = 'fxr on ' + (msg ? 'no' : 'ok');
        out.textContent = msg || 'ผ่าน โครงสร้างสูตรครบตามที่โจทย์ต้องการ';
        if (!msg) { pass++; }
        items.push({ text: q, ok: !msg });
        notes.push([q, el.value]);
      });
      var sc = document.getElementById('fxScore');
      sc.textContent = 'ผ่าน ' + pass + ' จาก 5 ข้อ';
      sc.className = 'labscore ' + (pass === 5 ? 'ok' : 'no');
      return { pass: pass, items: items, notes: notes };
    }
    document.getElementById('fxCheck').addEventListener('click', fxRun);
    document.getElementById('fxClear').addEventListener('click', function () {
      [].forEach.call(fxWrap.querySelectorAll('input[data-fx]'), function (el) { el.value = ''; });
      [].forEach.call(fxWrap.querySelectorAll('.fxr'), function (el) { el.className = 'fxr'; });
      document.getElementById('fxScore').textContent = '';
      scheduleSave(null);
    });
    WORK.push({
      title: 'ห้องฝึกเขียนสูตร',
      run: function () {
        var r = fxRun();
        return { score: 'ผ่าน ' + r.pass + ' จาก 5 ข้อ', items: r.items, notes: r.notes };
      }
    });
  }


  /* ================= กระดานล่าสูตรผิด และห้องฝึกตัดสินใจเรื่องข้อมูล =================
     ทั้งสองห้องใช้โครงเดียวกัน คือเลือกคำตอบจากรายการ แล้วเทียบกับเฉลยพร้อมคำอธิบาย */
  function makePicker(cfg) {
    var wrap = document.getElementById(cfg.box);
    if (!wrap) { return; }
    var KEY = JSON.parse(document.getElementById(cfg.key).textContent);
    var sels = [].slice.call(wrap.querySelectorAll('[' + cfg.attr + ']'));

    function run() {
      var pass = 0, items = [], notes = [];
      sels.forEach(function (sel) {
        var id = sel.getAttribute(cfg.attr);
        var k = KEY[id];
        var q = sel.parentNode.querySelector('.fxq').textContent.trim();
        var out = wrap.querySelector('[' + cfg.rattr + '="' + id + '"]');
        var ok = sel.value !== '' && sel.value === k.ans;
        if (ok) { pass++; }
        out.className = 'fxr on ' + (ok ? 'ok' : 'no');
        out.textContent = (sel.value === '' ? 'ยังไม่ได้ตอบข้อนี้ · ' : (ok ? 'ตอบถูก · ' : 'ยังไม่ใช่ · '))
          + k.why;
        items.push({ text: q, ok: ok });
        notes.push([q, sel.value === '' ? 'ยังไม่ได้ตอบ'
          : sel.options[sel.selectedIndex].text]);
      });
      var sc = document.getElementById(cfg.score);
      sc.textContent = 'ตอบถูก ' + pass + ' จาก ' + sels.length + ' ข้อ'
        + (pass >= cfg.passAt ? ' · ' + cfg.okMsg : ' · ' + cfg.noMsg);
      sc.className = 'labscore ' + (pass >= cfg.passAt ? 'ok' : 'no');
      return { pass: pass, items: items, notes: notes };
    }
    document.getElementById(cfg.check).addEventListener('click', run);
    document.getElementById(cfg.clear).addEventListener('click', function () {
      sels.forEach(function (s2) { s2.value = ''; });
      [].forEach.call(wrap.querySelectorAll('.fxr'), function (el) { el.className = 'fxr'; });
      document.getElementById(cfg.score).textContent = '';
      scheduleSave(null);
    });
    WORK.push({
      title: cfg.title,
      run: function () {
        var r = run();
        return { score: 'ตอบถูก ' + r.pass + ' จาก ' + sels.length + ' ข้อ',
                 items: r.items, notes: r.notes };
      }
    });
  }


  /* ================= เครื่องประกอบคำสั่งสี่ส่วน ================= */
  var pmBox = document.getElementById('pmBoard');
  if (pmBox) {
    var pmF = {};
    [].forEach.call(pmBox.querySelectorAll('[data-pm]'), function (el) {
      pmF[el.getAttribute('data-pm')] = el;
    });
    var PM_LABEL = { ctx: 'บริบท', data: 'โครงสร้างข้อมูล', want: 'สิ่งที่ต้องการ', fmt: 'รูปแบบผลลัพธ์และกรณีขอบ' };

    function pmText() {
      var out = [];
      ['ctx', 'data', 'want', 'fmt'].forEach(function (k) {
        var v = (pmF[k].value || '').trim();
        if (v) { out.push(PM_LABEL[k] + ': ' + v); }
      });
      return out.join('\\n');
    }
    function pmDraw() {
      var t = pmText();
      document.getElementById('pmOut').textContent = t || 'ยังไม่ได้กรอกส่วนใดเลย';
    }
    pmBox.addEventListener('input', pmDraw);
    pmDraw();

    function pmRun() {
      var all = pmText();
      var d = (pmF.data.value || '');
      // นับชื่อคอลัมน์ที่ระบุไว้ ยอมรับทั้งการเขียนว่าคอลัมน์ ตามด้วยชื่อ หรือใส่ไว้ในเครื่องหมายคำพูด
      var colHits = (d.match(/คอลัมน์|ช่อง|column/gi) || []).length
        + (d.match(/[A-Z]\s*(?:ถึง|-|:)\s*[A-Z]/g) || []).length;
      var res = {
        ctx: /excel|sheets|ชีต|สเปรดชีต|กูเกิล|google/i.test(pmF.ctx.value + ' ' + d),
        cols: colHits >= 2 || (d.match(/,/g) || []).length >= 2,
        size: /\d/.test(d),
        want: (pmF.want.value || '').trim().length >= 25,
        edge: /ว่าง|ไม่พบ|ซ้ำ|error|#n\/a|ผิดพลาด|ติดลบ|ข้าม/i.test(all)
      };
      var pass = 0, items = [];
      [].forEach.call(document.querySelectorAll('#pmChk li'), function (li) {
        var ok = !!res[li.getAttribute('data-k')];
        li.classList.remove('pass', 'fail');
        li.classList.add(ok ? 'pass' : 'fail');
        if (ok) { pass++; }
        items.push({ text: li.textContent.trim(), ok: ok });
      });
      var sc = document.getElementById('pmScore');
      sc.textContent = 'ครบ ' + pass + ' จาก 5 ข้อ'
        + (pass === 5 ? ' · คำสั่งนี้ใช้ได้จริงแล้ว' : ' · เติมส่วนที่ยังเป็นกากบาท');
      sc.className = 'labscore ' + (pass === 5 ? 'ok' : 'no');
      return { pass: pass, items: items };
    }
    document.getElementById('pmCheck').addEventListener('click', pmRun);
    document.getElementById('pmCopy').addEventListener('click', function () {
      var t = pmText();
      var sc = document.getElementById('pmScore');
      if (!t) { sc.textContent = 'ยังไม่มีข้อความให้คัดลอก'; sc.className = 'labscore no'; return; }
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(t).then(function () {
          sc.textContent = 'คัดลอกคำสั่งแล้ว นำไปวางในเครื่องมือ AI ได้เลย';
          sc.className = 'labscore ok';
        });
      } else {
        sc.textContent = 'เบราว์เซอร์นี้ไม่รองรับการคัดลอก ให้เลือกข้อความในกรอบแล้วคัดลอกเอง';
        sc.className = 'labscore no';
      }
    });
    WORK.push({
      title: 'คำสั่งสี่ส่วน แบบฝึกหัด 3.1',
      run: function () {
        var r = pmRun();
        return { score: 'ครบ ' + r.pass + ' จาก 5 ข้อ', items: r.items,
                 notes: [['คำสั่งเต็มที่เขียนไว้', pmText()]] };
      }
    });
  }

  makePicker({
    box: 'bhBoard', key: 'bhKey', attr: 'data-bh', rattr: 'data-bhr',
    check: 'bhCheck', clear: 'bhClear', score: 'bhScore', passAt: 5,
    title: 'กระดานล่าสูตรผิด แบบฝึกหัด 3.2',
    okMsg: 'ผ่านเกณฑ์ ตรวจสอบเป็นแล้ว', noMsg: 'อ่านคำอธิบายใต้ข้อที่ยังไม่ใช่ แล้วลองใหม่ได้'
  });
  makePicker({
    box: 'ethics', key: 'etKey', attr: 'data-et', rattr: 'data-etr',
    check: 'etCheck', clear: 'etClear', score: 'etScore', passAt: 5,
    title: 'ห้องฝึกตัดสินใจเรื่องข้อมูลกับ AI',
    okMsg: 'แยกกรณีได้ดี', noMsg: 'ลองอ่านหลักสามคำถามท้ายหน้าอีกครั้ง'
  });


  /* ================= ห้องฝึกจดหมายเวียน ================= */
  makePicker({
    box: 'mgBoard', key: 'mgKey', attr: 'data-mg', rattr: 'data-mgr',
    check: 'mgCheck', clear: 'mgClear', score: 'mgScore', passAt: 5,
    title: 'ห้องฝึกแก้ปัญหาจดหมายเวียน แบบฝึกหัด 2.5',
    okMsg: 'แยกสาเหตุได้ถูก', noMsg: 'อ่านคำอธิบายใต้ข้อที่ยังไม่ใช่ แล้วลองใหม่ได้'
  });

  var mgBox = document.getElementById('mgBoard');
  if (mgBox) {
    /* ตรวจรหัสรูปแบบของเขตข้อมูล ตรวจที่โครงสร้าง ไม่ได้ตรวจว่าพิมพ์เหมือนเฉลยทุกตัวอักษร */
    var MGFX = {
      d1: function (v) {
        var u = v.replace(/\\s+/g, '');
        if (u === '') { return 'ยังไม่ได้เขียน'; }
        if (u.indexOf('\\@') < 0) { return 'รหัสรูปแบบวันที่ต้องขึ้นต้นด้วยแบ็กสแลชตามด้วยเครื่องหมายแอท'; }
        if (!/"[^"]*"/.test(v)) { return 'ใส่รูปแบบไว้ในเครื่องหมายคำพูดคู่ด้วย'; }
        var pat = v.match(/"([^"]*)"/)[1];
        if (pat.indexOf('MMMM') < 0) { return 'ต้องการชื่อเดือนแบบเต็ม จึงต้องใช้ MMMM สี่ตัว ถ้าใช้ MM จะได้เลขเดือน'; }
        if (!/yyyy/.test(pat)) { return 'ต้องการปีสี่หลัก จึงต้องใช้ yyyy สี่ตัว'; }
        if (!/(^|[^d])d([^d]|$)/.test(pat)) { return 'ต้องการวันแบบไม่เติมศูนย์นำหน้า จึงใช้ d ตัวเดียว'; }
        return '';
      },
      n1: function (v) {
        var u = v.replace(/\\s+/g, '');
        if (u === '') { return 'ยังไม่ได้เขียน'; }
        if (u.indexOf('\\#') < 0) { return 'รหัสรูปแบบตัวเลขต้องขึ้นต้นด้วยแบ็กสแลชตามด้วยเครื่องหมายสี่เหลี่ยม'; }
        if (!/"[^"]*"/.test(v)) { return 'ใส่รูปแบบไว้ในเครื่องหมายคำพูดคู่ด้วย'; }
        var pat = v.match(/"([^"]*)"/)[1].replace(/\\s/g, '');
        if (pat.indexOf(',') < 0) { return 'ยังไม่มีเครื่องหมายคั่นหลักพัน ต้องเขียนแบบมีจุลภาคอยู่ในรูปแบบ'; }
        if (!/\\.00/.test(pat)) { return 'ต้องการทศนิยมสองตำแหน่งเสมอ จึงต้องเขียนจุดตามด้วยศูนย์สองตัว ไม่ใช่เครื่องหมายสี่เหลี่ยม'; }
        return '';
      }
    };
    function mgFxRun() {
      var pass = 0, items = [], notes = [];
      [].forEach.call(mgBox.querySelectorAll('input[data-mgfx]'), function (el) {
        var k = el.getAttribute('data-mgfx');
        var q = el.parentNode.querySelector('.fxq').textContent.trim();
        var out = mgBox.querySelector('[data-mgfxr="' + k + '"]');
        var msg = MGFX[k](el.value);
        out.className = 'fxr on ' + (msg ? 'no' : 'ok');
        out.textContent = msg || 'ผ่าน รหัสรูปแบบนี้ใช้ได้จริงใน Word';
        if (!msg) { pass++; }
        items.push({ text: q, ok: !msg });
        notes.push([q, el.value]);
      });
      var sc = document.getElementById('mgFxScore');
      sc.textContent = 'ผ่าน ' + pass + ' จาก 2 ข้อ';
      sc.className = 'labscore ' + (pass === 2 ? 'ok' : 'no');
      return { pass: pass, items: items, notes: notes };
    }
    document.getElementById('mgFxCheck').addEventListener('click', mgFxRun);
    WORK.push({
      title: 'รหัสรูปแบบเขตข้อมูล แบบฝึกหัด 2.5',
      run: function () {
        var r = mgFxRun();
        return { score: 'ผ่าน ' + r.pass + ' จาก 2 ข้อ', items: r.items, notes: r.notes };
      }
    });
  }

  /* ---------- ส่งออกไฟล์งานสำหรับอัปโหลดเข้า Google Classroom ---------- */
  (function () {
    if (!document.getElementById('expCsv')) { return; }
    function q(id) { var el = document.getElementById(id); return el ? el.value.trim() : ''; }
    function esc(v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; }
    function row() { return [].slice.call(arguments).map(esc).join(','); }
    function nonEmpty(r) { return r.some(function (v) { return String(v).trim() !== ''; }); }

    function buildRows() {
      var d = new Date();
      var stamp = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-'
        + ('0' + d.getDate()).slice(-2) + ' ' + ('0' + d.getHours()).slice(-2) + ':'
        + ('0' + d.getMinutes()).slice(-2);
      var L = [];
      L.push(row('ใบงานระดับ %(levelname)s', 'ชุดการเรียนรู้ Excel และ Google Sheets'));
      L.push(row('ชื่อ-นามสกุล', q('stName') || 'ยังไม่ได้กรอก'));
      L.push(row('รหัสนักศึกษา', q('stId') || 'ยังไม่ได้กรอก'));
      L.push(row('สาขา', q('stProg') || 'ยังไม่ได้กรอก'));
      L.push(row('เวลาที่ส่งออกไฟล์', stamp));
      L.push('');

      WORK.forEach(function (w) {
        var r = w.run() || {};
        L.push(row(w.title, r.score || ''));
        (r.items || []).forEach(function (it) {
          L.push(row(it.text, it.ok ? 'ผ่าน' : 'ยังไม่ผ่าน'));
        });
        if (r.cols && r.cols.length) { L.push(r.cols.map(esc).join(',')); }
        (r.rows || []).forEach(function (rw) { if (nonEmpty(rw)) { L.push(rw.map(esc).join(',')); } });
        (r.notes || []).forEach(function (n) {
          L.push(row(n[0], String(n[1] || '').trim() || 'ยังไม่ได้เขียน'));
        });
        L.push('');
      });

      L.push(row('คำตอบแบบฝึกหัดย่อย'));
      [].forEach.call(document.querySelectorAll('.ws'), function (ws) {
        L.push(row(ws.querySelector('.wsc').textContent, ws.querySelector('.wsh b').textContent));
        [].forEach.call(ws.querySelectorAll('.wsans .fld'), function (f) {
          L.push(row(f.querySelector('label').textContent,
                     f.querySelector('textarea').value.trim() || 'ยังไม่ได้ตอบ'));
        });
        L.push('');
      });
      L.push(row('แบบตรวจความเข้าใจ',
        quizDone === 0 ? 'ยังไม่ได้ทำ' : 'ตอบถูก ' + quizRight + ' จาก ' + quizDone + ' ข้อที่ทำ'));
      return L;
    }

    function say(msg, good) {
      var m = document.getElementById('expMsg');
      m.textContent = msg;
      m.className = 'labscore ' + (good ? 'ok' : 'no');
    }

    document.getElementById('expCsv').addEventListener('click', function () {
      var text = buildRows().join('\\r\\n');
      // เครื่องหมาย BOM ข้างหน้าทำให้ Excel รู้ว่าเป็น UTF-8 ภาษาไทยจึงไม่เพี้ยน
      var blob = new Blob(['\\ufeff' + text], { type: 'text/csv;charset=utf-8;' });
      var name = 'ใบงาน-%(levelname)s-' + (q('stId') || q('stName') || 'ไม่ระบุ') + '.csv';
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
      say('ดาวน์โหลดแล้ว ชื่อไฟล์ ' + name + ' · นำไปอัปโหลดในงานที่อาจารย์มอบหมายได้เลย', true);
    });

    document.getElementById('expCopy').addEventListener('click', function () {
      var txt = buildRows().join('\\n').replace(/"/g, '');
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(function () {
          say('คัดลอกแล้ว วางในช่องความคิดเห็นส่วนตัวของงานใน Classroom ได้', true);
        }, function () { say('คัดลอกไม่สำเร็จ กรุณาใช้ปุ่มดาวน์โหลดไฟล์แทน', false); });
      } else {
        say('เบราว์เซอร์นี้ไม่รองรับการคัดลอก กรุณาใช้ปุ่มดาวน์โหลดไฟล์แทน', false);
      }
    });
  })();
})();
</script>
</body>
</html>
""" % dict(C, nav=nav_html, modules=modules_html, quiz=quiz_html, check=check_html,
           hero=L.HERO, start=L.START, labs="\n\n".join(L.LABS),
           explist="".join("<li>{}</li>".format(e(x)) for x in L.EXPORT_LIST),
           navextra=nav_extra, storekey=L.LEVEL["store"], levelname=L.LEVEL["name"],
           noindex=('<meta name="robots" content="noindex, nofollow">\n' if NOINDEX else ""))

open(OUT, "w", encoding="utf-8").write(HTML)
print("wrote", OUT, len(HTML), "bytes")
