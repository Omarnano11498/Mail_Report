#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
نظام أتمتة متابعة بريد Zoho Mail
- يجلب الرسائل عبر IMAP (قراءة فقط، دون تغيير حالة "مقروء")
- يصنّفها وفق قواعد قابلة للتعديل في config.json
- يرتّبها حسب المرسل (الأكثر مراسلةً أولًا)
- يصدّر تقريرًا منسقًا بالعربية إلى ملف .docx

الاستخدام:
    python zoho_mail_report.py                 # آخر N يوم (من الإعدادات)
    python zoho_mail_report.py --days 3        # آخر 3 أيام
    python zoho_mail_report.py --new-only      # الرسائل الجديدة منذ آخر تشغيل فقط
    python zoho_mail_report.py --unseen        # غير المقروءة فقط
    python zoho_mail_report.py --demo          # تجربة ببيانات وهمية دون اتصال
"""
import argparse
import datetime as dt
import email
import imaplib
import json
import os
import re
import sys
from collections import defaultdict, Counter
from email.header import decode_header, make_header
from email.utils import parseaddr, parsedate_to_datetime

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

DEFAULT_CATEGORY = "أخرى"


# ───────────────────────────── الإعدادات ─────────────────────────────
def load_config(path):
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    # كلمة المرور من متغير بيئة أولًا (أكثر أمانًا من حفظها في الملف)
    cfg["password"] = os.environ.get("ZOHO_PASSWORD") or cfg.get("password", "")
    cfg["email"] = os.environ.get("ZOHO_EMAIL") or cfg.get("email", "")
    return cfg


# ───────────────────────────── الجلب عبر IMAP ─────────────────────────────
def decode_str(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def extract_snippet(msg, limit=220):
    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and "attachment" not in str(part.get("Content-Disposition", "")):
                text = _payload(part)
                break
        if not text:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    text = re.sub(r"<[^>]+>", " ", _payload(part))
                    break
    else:
        text = _payload(msg)
        if msg.get_content_type() == "text/html":
            text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit] + ("…" if len(text) > limit else "")


def _payload(part):
    raw = part.get_payload(decode=True) or b""
    charset = part.get_content_charset() or "utf-8"
    try:
        return raw.decode(charset, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


def has_attachments(msg):
    return any(
        "attachment" in str(p.get("Content-Disposition", "")).lower()
        for p in msg.walk()
    ) if msg.is_multipart() else False


def fetch_messages(cfg, days, folders, unseen_only, skip_uids):
    imap = imaplib.IMAP4_SSL(cfg["imap_host"], int(cfg.get("imap_port", 993)))
    imap.login(cfg["email"], cfg["password"])
    since = (dt.date.today() - dt.timedelta(days=days)).strftime("%d-%b-%Y")
    results = []
    max_n = int(cfg.get("max_messages_per_folder", 500))

    for folder in folders:
        typ, _ = imap.select(f'"{folder}"', readonly=True)
        if typ != "OK":
            print(f"⚠️ تعذّر فتح المجلد: {folder}", file=sys.stderr)
            continue
        criteria = ["SINCE", since] + (["UNSEEN"] if unseen_only else [])
        typ, data = imap.uid("search", None, *criteria)
        uids = data[0].split()[-max_n:] if typ == "OK" else []

        for uid in uids:
            key = f"{folder}:{uid.decode()}"
            if key in skip_uids:
                continue
            typ, msg_data = imap.uid("fetch", uid, "(BODY.PEEK[] FLAGS)")
            if typ != "OK" or not msg_data or msg_data[0] is None:
                continue
            raw = msg_data[0][1]
            # بعض الخوادم ترسل FLAGS بعد جسم الرسالة في عنصر منفصل، فنجمع كل أجزاء الرد
            flags = " ".join(
                (part[0] if isinstance(part, tuple) else part).decode(errors="ignore")
                for part in msg_data if part
            )
            msg = email.message_from_bytes(raw)
            name, addr = parseaddr(decode_str(msg.get("From")))
            try:
                date = parsedate_to_datetime(msg.get("Date"))
                # التواريخ بلا منطقة زمنية (-0000) تُعاد naive وتسبب خطأ عند الترتيب
                if date.tzinfo is None:
                    date = date.replace(tzinfo=dt.timezone.utc)
            except Exception:
                date = None
            results.append({
                "key": key,
                "folder": folder,
                "from_name": name or addr,
                "from_email": addr.lower(),
                "subject": decode_str(msg.get("Subject")) or "(بدون عنوان)",
                "date": date,
                "snippet": extract_snippet(msg),
                "seen": "\\Seen" in flags,
                "attachments": has_attachments(msg),
            })
    imap.logout()
    return results


# ───────────────────────────── التصنيف ─────────────────────────────
def classify(msg, rules):
    """يعيد (التصنيف، هل هي عاجلة)."""
    haystack = f"{msg['subject']} {msg['snippet']}".lower()
    sender = msg["from_email"]
    best, best_score = DEFAULT_CATEGORY, 0
    for cat in rules["categories"]:
        score = 0
        score += sum(3 for p in cat.get("sender_patterns", []) if p.lower() in sender)
        score += sum(1 for k in cat.get("keywords", []) if k.lower() in haystack)
        if score > best_score:
            best, best_score = cat["name"], score
    urgent = any(k.lower() in haystack for k in rules.get("urgent_keywords", []))
    return best, urgent


# ───────────────────────────── تنسيق Word بالعربية ─────────────────────────────
# Word يرفض الملف ("محتوى غير قابل للقراءة") إذا لم تكن العناصر بالترتيب الذي يحدده مخطط OOXML،
# لذا نُدرج كل عنصر قبل أول عنصر لاحق له بدل إلحاقه في النهاية.
_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
              "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
              "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
              "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
              "specVanish", "oMath"]
_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
              "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
              "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
              "rPr", "sectPr", "pPrChange"]
_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
                "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
                "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
_SECTPR_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
                 "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
                 "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
                 "sectPrChange"]


def _set_child(parent, tag, order, val=None):
    """يضيف (أو يعيد) العنصر w:<tag> في موضعه الصحيح وفق ترتيب المخطط."""
    el = parent.find(qn(f"w:{tag}"))
    if el is None:
        el = OxmlElement(f"w:{tag}")
        later = set(order[order.index(tag) + 1:])
        nxt = next((c for c in parent if c.tag.split("}")[-1] in later), None)
        if nxt is not None:
            nxt.addprevious(el)
        else:
            parent.append(el)
    if val is not None:
        el.set(qn("w:val"), val)
    return el


def _rtl_paragraph(p, align=None):
    _set_child(p._p.get_or_add_pPr(), "bidi", _PPR_ORDER, "1")
    # في الفقرات العربية (bidi) يكون الاتجاه الطبيعي يمينًا؛ نحدد المحاذاة فقط عند التوسيط
    if align is not None:
        p.alignment = align


def _rtl_run(run, size=None, bold=False, color=None):
    rPr = run._r.get_or_add_rPr()
    _set_child(rPr, "rtl", _RPR_ORDER, "1")
    fonts = _set_child(rPr, "rFonts", _RPR_ORDER)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        fonts.set(qn(attr), "Arial")
    run.bold = bold
    if bold:
        _set_child(rPr, "bCs", _RPR_ORDER)
    if size:
        run.font.size = Pt(size)
        _set_child(rPr, "szCs", _RPR_ORDER, str(int(size * 2)))
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_text(doc_or_cell, text, size=11, bold=False, color=None, align=None, space_after=4):
    p = doc_or_cell.add_paragraph()
    _rtl_paragraph(p, align)
    p.paragraph_format.space_after = Pt(space_after)
    _rtl_run(p.add_run(text), size=size, bold=bold, color=color)
    return p


def add_heading(doc, text, level=1):
    sizes = {0: 22, 1: 16, 2: 13}
    colors = {0: "1F3864", 1: "2E74B5", 2: "404040"}
    return add_text(doc, text, size=sizes[level], bold=True, color=colors[level], space_after=8)


def _shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_child(table._tbl.tblPr, "bidiVisual", _TBLPR_ORDER)

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
        add_text(cell, h, size=10, bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
        _shade(cell, "2E74B5")

    for r_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cell = cells[i]
            cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
            add_text(cell, str(val), size=9, space_after=0)
            if r_idx % 2:
                _shade(cell, "EAF1FB")

    if widths:
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return table


def fmt_date(d):
    return d.astimezone().strftime("%Y-%m-%d %H:%M") if d else "—"


def build_report(messages, out_path, account, period_label):
    doc = Document()
    sec = doc.sections[0]
    sec.right_margin = sec.left_margin = Cm(1.8)
    _set_child(sec._sectPr, "bidi", _SECTPR_ORDER)

    # الغلاف والملخص
    add_heading(doc, "تقرير متابعة البريد الإلكتروني — Zoho Mail", 0)
    add_text(doc, f"الحساب: {account}", size=11, color="595959")
    add_text(doc, f"الفترة: {period_label}", size=11, color="595959")
    add_text(doc, f"تاريخ الإنشاء: {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}", size=11, color="595959", space_after=12)

    by_sender = defaultdict(list)
    for m in messages:
        by_sender[m["from_email"]].append(m)
    cat_count = Counter(m["category"] for m in messages)
    urgent = [m for m in messages if m["urgent"]]
    unread = sum(1 for m in messages if not m["seen"])

    add_heading(doc, "الملخص التنفيذي", 1)
    add_table(doc, ["المؤشر", "القيمة"], [
        ["إجمالي الرسائل", len(messages)],
        ["عدد المرسلين", len(by_sender)],
        ["غير المقروءة", unread],
        ["رسائل عاجلة", len(urgent)],
    ], widths=[8, 4])

    add_heading(doc, "التوزيع حسب التصنيف", 1)
    add_table(doc, ["التصنيف", "عدد الرسائل", "النسبة"],
              [[c, n, f"{n * 100 / max(len(messages), 1):.0f}%"] for c, n in cat_count.most_common()],
              widths=[7, 3, 3])

    if urgent:
        add_heading(doc, "⚠ رسائل عاجلة تحتاج متابعة", 1)
        add_table(doc, ["التاريخ", "المرسل", "الموضوع"],
                  [[fmt_date(m["date"]), m["from_name"], m["subject"]]
                   for m in sorted(urgent, key=lambda x: x["date"] or dt.datetime.min.replace(tzinfo=dt.timezone.utc), reverse=True)],
                  widths=[3.5, 4.5, 9])

    # فهرس المرسلين
    senders_sorted = sorted(by_sender.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    add_heading(doc, "ترتيب المرسلين حسب حجم المراسلة", 1)
    add_table(doc, ["#", "المرسل", "البريد", "الرسائل", "التصنيف الغالب"],
              [[i + 1, msgs[0]["from_name"], addr, len(msgs),
                Counter(m["category"] for m in msgs).most_common(1)[0][0]]
               for i, (addr, msgs) in enumerate(senders_sorted)],
              widths=[1, 4, 5.5, 2, 4])

    # التفاصيل لكل مرسل
    add_heading(doc, "تفاصيل الرسائل حسب المرسل", 1)
    for addr, msgs in senders_sorted:
        add_heading(doc, f"{msgs[0]['from_name']} — {len(msgs)} رسالة  ({addr})", 2)
        msgs.sort(key=lambda x: x["date"] or dt.datetime.min.replace(tzinfo=dt.timezone.utc), reverse=True)
        add_table(doc, ["التاريخ", "الموضوع", "التصنيف", "الحالة", "مقتطف"],
                  [[fmt_date(m["date"]), m["subject"], m["category"] + (" • عاجل" if m["urgent"] else ""),
                    ("مقروءة" if m["seen"] else "غير مقروءة") + (" • بمرفقات" if m["attachments"] else ""),
                    m["snippet"]] for m in msgs],
                  widths=[2.8, 4.2, 2.6, 2.2, 5.6])

    doc.save(out_path)


# ───────────────────────────── حفظ الحالة (للمتابعة المستمرة) ─────────────────────────────
def load_state(path):
    try:
        with open(path, encoding="utf-8") as f:
            return set(json.load(f).get("seen_keys", []))
    except FileNotFoundError:
        return set()


def save_state(path, keys):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"seen_keys": sorted(keys), "updated": dt.datetime.now().isoformat()}, f, ensure_ascii=False)


# ───────────────────────────── بيانات تجريبية ─────────────────────────────
def demo_messages():
    now = dt.datetime.now(dt.timezone.utc)
    s = [
        ("أحمد سمير", "ahmed@client-co.com", "طلب عرض سعر للمشروع الجديد", "نرجو إرسال عرض السعر قبل نهاية الأسبوع", 1),
        ("أحمد سمير", "ahmed@client-co.com", "عاجل: تعديل موعد الاجتماع", "نحتاج تأجيل الاجتماع إلى الخميس", 0),
        ("أحمد سمير", "ahmed@client-co.com", "ملاحظات على العقد", "مرفق ملاحظات الإدارة القانونية", 2),
        ("Zoho Billing", "billing@zoho.com", "Invoice #INV-2291", "Your invoice for October is ready. Amount due", 1),
        ("المحاسبة", "finance@mycompany.com", "فاتورة شهر أكتوبر", "مرفق الفاتورة والإيصال للمراجعة", 3),
        ("المحاسبة", "finance@mycompany.com", "تذكير بسداد مستحقات", "يرجى السداد في أقرب وقت", 0),
        ("LinkedIn", "news@linkedin.com", "Weekly newsletter", "Top stories for you. Unsubscribe", 2),
        ("Noon Offers", "offers@noon.com", "خصم 50% لفترة محدودة", "عرض حصري لك - إلغاء الاشتراك", 4),
        ("فريق الدعم", "support@hosting.com", "Ticket #4471 - Server down", "We are investigating the issue urgently", 0),
        ("سارة علي", "sara@partner.org", "دعوة لورشة عمل التدريب", "يسعدنا دعوتكم لحضور ورشة الذكاء الاصطناعي meeting", 5),
        ("سارة علي", "sara@partner.org", "متابعة الاتفاقية", "هل تمت مراجعة المسودة؟", 6),
        ("Google Calendar", "calendar-notification@google.com", "Invitation: Weekly sync", "Join with Google Meet", 1),
    ]
    out = []
    for i, (n, e, subj, snip, d) in enumerate(s):
        out.append({"key": f"DEMO:{i}", "folder": "INBOX", "from_name": n, "from_email": e,
                    "subject": subj, "date": now - dt.timedelta(days=d, hours=i), "snippet": snip,
                    "seen": i % 3 != 0, "attachments": "مرفق" in snip})
    return out


# ───────────────────────────── التشغيل ─────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="أتمتة متابعة بريد Zoho Mail وتصديره إلى Word")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--days", type=int)
    ap.add_argument("--unseen", action="store_true", help="غير المقروءة فقط")
    ap.add_argument("--new-only", action="store_true", help="الرسائل الجديدة منذ آخر تشغيل فقط")
    ap.add_argument("--out", help="مسار ملف docx الناتج")
    ap.add_argument("--demo", action="store_true", help="تشغيل تجريبي ببيانات وهمية")
    args = ap.parse_args()

    cfg = load_config(args.config)
    days = args.days or int(cfg.get("days_back", 7))
    state_path = cfg.get("state_file", "mail_state.json")
    skip = load_state(state_path) if args.new_only else set()

    if args.demo:
        messages, account = demo_messages(), "demo@example.com"
    else:
        if not cfg["email"] or not cfg["password"]:
            sys.exit("❌ أدخل البريد وكلمة مرور التطبيق في config.json أو متغيري البيئة ZOHO_EMAIL / ZOHO_PASSWORD")
        messages = fetch_messages(cfg, days, cfg.get("folders", ["INBOX"]), args.unseen, skip)
        account = cfg["email"]

    for m in messages:
        m["category"], m["urgent"] = classify(m, cfg["rules"])

    out_dir = cfg.get("output_dir", "reports")
    os.makedirs(out_dir, exist_ok=True)
    out = args.out or os.path.join(out_dir, f"mail_report_{dt.datetime.now():%Y-%m-%d_%H%M}.docx")
    label = "منذ آخر تشغيل" if args.new_only else f"آخر {days} " + ("يوم" if days == 1 else "أيام")

    if not messages:
        print("ℹ️ لا توجد رسائل جديدة في الفترة المحددة.")
        return
    build_report(messages, out, account, label)

    if not args.demo:
        save_state(state_path, load_state(state_path) | {m["key"] for m in messages})
    print(f"✅ تم تصدير {len(messages)} رسالة من {len({m['from_email'] for m in messages})} مرسل إلى: {out}")


if __name__ == "__main__":
    main()
