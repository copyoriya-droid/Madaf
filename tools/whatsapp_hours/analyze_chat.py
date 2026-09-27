#!/usr/bin/env python3
"""חישוב שעות עבודה מייצוא צ'אט וואטסאפ.

שלב 1 (ניתוח):
    python analyze_chat.py chat.zip --me "אוריה" --month 9 --year 2026 --cap-min 15 --deliv-value-h 0
מדפיס את הזמן שנמדד אוטומטית ורשימות ממוספרות של מועמדים לאישור
(תסריטים, הפצות, שיחות, פגישות), ושומר analysis.json.

שלב 2 (סיכום): אחרי שמאשרים את המועמדים, כותבים decisions.json ומריצים שוב עם
    --decisions decisions.json [--xlsx hours.xlsx]

בלי --me הסקריפט מדפיס את רשימת השולחים בצ'אט.
"""
import argparse
import calendar
import io
import json
import re
import sys
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta

# תווי כיווניות ורווחים מיוחדים שוואטסאפ משתיל בייצוא
INVISIBLE = dict.fromkeys(map(ord, "‎‏‪‫‬‭‮⁦⁧⁨⁩﻿"), None)

# iOS:     [27.9.2026, 14:03:22] שם: טקסט
# Android: 27.9.2026, 14:03 - שם: טקסט   (גם 27/09/26 ו-AM/PM)
LINE_RE = re.compile(
    r"^\[?(?P<d>\d{1,2})[./-](?P<m>\d{1,2})[./-](?P<y>\d{2,4}),?\s+"
    r"(?P<H>\d{1,2}):(?P<M>\d{2})(?::(?P<S>\d{2}))?\s*(?P<ampm>[AaPp]\.?[Mm]\.?)?\]?"
    r"\s*(?:-\s*)?(?P<rest>.*)$"
)

MEDIA_RE = re.compile(r"(הושמט|הושמטה|omitted|\.mp4|\.jpg|\.jpeg|\.png|\.mov|<מצורף|<attached)", re.I)
VIDEO_RE = re.compile(r"(\.mp4|\.mov|סרטון הווידאו הושמט|video omitted)", re.I)
UPLOAD_RE = re.compile(
    r"https?://(?:www\.)?(?:instagram\.com/(?:reel|reels|p|tv)/|facebook\.com/|fb\.watch/|m\.facebook\.com/|"
    r"tiktok\.com/|vm\.tiktok\.com/)\S+",
    re.I,
)
DOC_RE = re.compile(r"https?://docs\.google\.com/document/\S+", re.I)
DISTRIB_RE = re.compile(r"(להפיץ|הפצה|הפצות|הופץ|הופצה|הופצו|מפיץ|מפיצה|מפיצים|תפיץ|תפיצי|נפיץ)")
CALL_RE = re.compile(r"(אני מתקשר|אני מתקשרת|אתקשר|נתקשר|מחייג|מחייגת|אחייג|מתקשר אלי|מתקשרת אלי)")
MEETING_RE = re.compile(r"(נקבע פגישה|לקבוע פגישה|ניפגש|נפגש ב|נפגשים|פגישה ב|פגישת עבודה)")
INVOICE_RE = re.compile(r"(דרישת תשלום|חשבונית|קבלה|receipt|invoice|tax_invoice)", re.I)
# סימן לעבודה בתוך רצף הודעות. רצף בלי אף סימן כזה נחשב שיחת חולין ולא נספר.
RE_WORK = re.compile(
    r"(תסריט|סרטון|וידאו|ויד|עריכה|לערוך|עורך|עורכת|הפצה|להפיץ|פוסט|העלאה|להעלות|עלה|קאפשן|כיתוב|כתוביות|"
    r"טקסט|גרסה|גירסה|תיקון|תיקונים|לתקן|לקוח|לקוחה|קמפיין|ריל|רילס|סטורי|קרוסלה|טיוטה|אישור|מאושר|"
    r"לאשר|קישור|לינק|כותרת|סקריפט|תוכן|דדליין|מחר ב|הבריף|בריף|הקלטה|צילום|לצלם|פרסום|לפרסם|"
    r"https?://|הושמט|\.mp4|script|video|edit|draft|post|reel)",
    re.I,
)


@dataclass
class Msg:
    ts: datetime
    sender: str
    text: str


def read_chat_text(path):
    if path.lower().endswith(".zip"):
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".txt")]
            if not names:
                sys.exit("לא נמצא קובץ txt בתוך ה-zip")
            name = next((n for n in names if "chat" in n.lower()), names[0])
            raw = z.read(name)
    else:
        with open(path, "rb") as f:
            raw = f.read()
    return io.TextIOWrapper(io.BytesIO(raw), encoding="utf-8-sig", errors="replace").read()


def parse_chat(text):
    msgs = []
    for line in text.splitlines():
        line = line.translate(INVISIBLE).replace(" ", " ").replace("\xa0", " ")
        m = LINE_RE.match(line)
        if m and ": " in m.group("rest"):
            y = int(m.group("y"))
            y += 2000 if y < 100 else 0
            H = int(m.group("H"))
            ampm = (m.group("ampm") or "").lower().replace(".", "")
            if ampm == "pm" and H < 12:
                H += 12
            elif ampm == "am" and H == 12:
                H = 0
            try:
                ts = datetime(y, int(m.group("m")), int(m.group("d")), H, int(m.group("M")), int(m.group("S") or 0))
            except ValueError:
                ts = None
            if ts:
                sender, _, body = m.group("rest").partition(": ")
                msgs.append(Msg(ts, sender.strip(), body.strip()))
                continue
        if msgs and line.strip():
            msgs[-1].text += "\n" + line.strip()
    msgs.sort(key=lambda x: x.ts)
    return msgs


def is_smalltalk_burst(burst):
    return not any(RE_WORK.search(m.text) for m in burst)


def bursts_of(msgs, cap_min):
    out, cur = [], []
    for m in msgs:
        if cur and (m.ts - cur[-1].ts) > timedelta(minutes=cap_min):
            out.append(cur)
            cur = []
        cur.append(m)
    if cur:
        out.append(cur)
    return out


def context(msgs, i, width=1):
    lo, hi = max(0, i - width), min(len(msgs), i + width + 1)
    return " | ".join(f"{m.sender}: {m.text[:80]}" for m in msgs[lo:hi]).replace("\n", " ")


def analyze(msgs, me, month, year, a):
    in_month = [m for m in msgs if m.ts.month == month and m.ts.year == year]
    day = lambda m: m.ts.date().isoformat()
    daily = defaultdict(lambda: defaultdict(float))  # day -> category -> hours
    notes = []

    # 1) זמן פעיל ברצפי הודעות
    excluded = 0
    for b in bursts_of(in_month, a.cap_min):
        if not a.keep_smalltalk and is_smalltalk_burst(b):
            excluded += 1
            continue
        mins = max((b[-1].ts - b[0].ts).total_seconds() / 60, a.burst_floor_min)
        daily[day(b[0])]["active"] += mins / 60
    notes.append(f"{excluded} רצפים של שיחת חולין לא נספרו (מסנן מילות עבודה)")

    # 2) סרטונים: וידאו אחרי הפסקה של break-min. ברירת המחדל בשיטה היא 0 שעות (הצלבה בלבד).
    videos, last_video = [], None
    for i, m in enumerate(in_month):
        if not VIDEO_RE.search(m.text):
            continue
        gap = (m.ts - in_month[i - 1].ts) if i else timedelta(days=1)
        if gap < timedelta(minutes=a.break_min):
            continue
        if last_video and (m.ts - last_video) <= timedelta(minutes=a.deliv_merge_min):
            continue
        last_video = m.ts
        videos.append({"ts": m.ts.isoformat(sep=" ", timespec="minutes"), "sender": m.sender})
        daily[day(m)]["videos"] += a.deliv_value_h

    # 3) העלאות: קישור לפוסט שפורסם, מהמשתמשת, אחרי שהופיעה מדיה באותו יום
    uploads, last_up = [], None
    media_days = set()
    for m in in_month:
        if MEDIA_RE.search(m.text):
            media_days.add(day(m))
        if m.sender == me and UPLOAD_RE.search(m.text) and day(m) in media_days:
            if last_up and (m.ts - last_up) <= timedelta(minutes=a.upload_merge_min):
                continue
            last_up = m.ts
            uploads.append({"ts": m.ts.isoformat(sep=" ", timespec="minutes"), "url": UPLOAD_RE.search(m.text).group(0)})
            daily[day(m)]["uploads"] += a.upload_value_min / 60

    # 4) מועמדים לאישור ידני
    def collect(pred):
        out = []
        for i, m in enumerate(in_month):
            if pred(m):
                out.append({
                    "n": len(out) + 1,
                    "ts": m.ts.isoformat(sep=" ", timespec="minutes"),
                    "day": day(m),
                    "sender": m.sender,
                    "text": m.text[:300],
                    "context": context(in_month, i),
                })
        return out

    scripts = collect(lambda m: m.sender == me and DOC_RE.search(m.text))
    for s in scripts:
        s["url"] = DOC_RE.search(s["text"]).group(0) if DOC_RE.search(s["text"]) else ""
        s["looks_like_invoice"] = bool(INVOICE_RE.search(s["context"]))
    distribution = collect(lambda m: DISTRIB_RE.search(m.text))
    calls = collect(lambda m: m.sender != me and CALL_RE.search(m.text))
    meetings = collect(lambda m: m.sender != me and MEETING_RE.search(m.text))

    return {
        "month": month, "year": year, "me": me,
        "settings": vars(a) | {"chat": None, "decisions": None, "xlsx": None},
        "daily": {d: dict(v) for d, v in sorted(daily.items())},
        "videos": videos, "uploads": uploads,
        "candidates": {"scripts": scripts, "distribution": distribution, "calls": calls, "meetings": meetings},
        "notes": notes,
    }


def apply_decisions(res, dec):
    daily = defaultdict(lambda: defaultdict(float), {d: defaultdict(float, v) for d, v in res["daily"].items()})
    details = []
    for cat in ("scripts", "distribution", "calls", "meetings"):
        by_n = {str(c["n"]): c for c in res["candidates"][cat]}
        for n, hours in (dec.get(cat) or {}).items():
            c = by_n.get(str(n))
            if c is None:
                print(f"אזהרה: אין מועמד מספר {n} בקטגוריה {cat}", file=sys.stderr)
                continue
            daily[c["day"]][cat] += float(hours)
            details.append({"cat": cat, "n": c["n"], "ts": c["ts"], "hours": float(hours), "text": c["text"][:120]})
    for extra in dec.get("extra") or []:  # {"day": "2026-09-03", "cat": "meetings", "hours": 0.5, "text": "..."}
        daily[extra["day"]][extra["cat"]] += float(extra["hours"])
        details.append({"cat": extra["cat"], "n": "ידני", "ts": extra["day"], "hours": float(extra["hours"]),
                        "text": extra.get("text", "")})
    res["daily"] = {d: dict(v) for d, v in sorted(daily.items())}
    res["details"] = details
    weekly = float(dec.get("inbox_hours_per_week") or 0)
    days = calendar.monthrange(res["year"], res["month"])[1]
    res["inbox_hours"] = round(weekly * days / 7, 2)
    return res


CATS = [("active", "זמן פעיל בצ'אט"), ("uploads", "העלאות"), ("videos", "סרטונים"), ("scripts", "תסריטים"),
        ("distribution", "הפצה"), ("calls", "שיחות"), ("meetings", "פגישות")]


def totals(res):
    t = {k: round(sum(v.get(k, 0) for v in res["daily"].values()), 2) for k, _ in CATS}
    t["inbox"] = res.get("inbox_hours", 0)
    t["total"] = round(sum(t.values()), 2)
    return t


def print_report(res, final):
    print(f"\n=== {res['month']}/{res['year']} · {res['me']} ===\n")
    print("יום         | " + " | ".join(label for _, label in CATS) + " | סה\"כ")
    for d, v in res["daily"].items():
        row = [v.get(k, 0) for k, _ in CATS]
        print(f"{d}  | " + " | ".join(f"{x:.2f}" for x in row) + f" | {sum(row):.2f}")
    t = totals(res)
    print("\nסיכום לפי קטגוריה:")
    for k, label in CATS:
        print(f"  {label}: {t[k]:.2f}")
    if t["inbox"]:
        print(f"  תיבת הודעות ותגובות (קבוע שבועי): {t['inbox']:.2f}")
    print(f"  סה\"כ: {t['total']:.2f} שעות")
    for n in res["notes"]:
        print("  ·", n)
    if final:
        return
    labels = {"scripts": "קישורי Google Docs ששלחת: לסמן מה מהם תסריט",
              "distribution": "אזכורי הפצה: לסמן מה קרה בפועל וכמה זמן",
              "calls": "הצד השני אמר שהוא מתקשר (ברירת מחדל 10 דק')",
              "meetings": "הצד השני הציע להיפגש (ברירת מחדל 30 דק')"}
    for cat, label in labels.items():
        items = res["candidates"][cat]
        print(f"\n--- {label} ({len(items)}) ---")
        for c in items:
            flag = " ⚠️ נראה כמו חשבונית/מסמך" if c.get("looks_like_invoice") else ""
            print(f"[{c['n']}] {c['ts']} · {c['context'][:200]}{flag}")
            if c.get("url"):
                print(f"     {c['url']}")
    print(f"\nסרטונים שזוהו להצלבה מול טבלת המעקב: {len(res['videos'])}")
    for v in res["videos"]:
        print(f"  · {v['ts']} ({v['sender']})")


def write_xlsx(res, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "סיכום יומי"
    ws.sheet_view.rightToLeft = True
    head = ["יום"] + [label for _, label in CATS] + ["סה\"כ"]
    ws.append(head)
    for d, v in res["daily"].items():
        r = ws.max_row + 1
        ws.append([d] + [round(v.get(k, 0), 2) for k, _ in CATS] + [f"=SUM(B{r}:{chr(65 + len(CATS))}{r})"])
    last = ws.max_row
    ws.append(["סה\"כ"] + [f"=SUM({chr(66 + i)}2:{chr(66 + i)}{last})" for i in range(len(CATS) + 1)])
    inbox = res.get("inbox_hours", 0)
    if inbox:
        ws.append(["תיבת הודעות ותגובות", inbox])
        ws.append(["סה\"כ לחיוב", f"={chr(66 + len(CATS))}{last + 1}+B{last + 2}"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
    for row in ws.iter_rows(min_row=last + 1):
        for c in row:
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 22

    wd = wb.create_sheet("פירוט")
    wd.sheet_view.rightToLeft = True
    wd.append(["קטגוריה", "מס'", "מתי", "שעות", "תיאור"])
    names = dict(CATS)
    for d in res.get("details", []):
        wd.append([names.get(d["cat"], d["cat"]), d["n"], d["ts"], d["hours"], d["text"]])
    wd.column_dimensions["E"].width = 80
    wb.save(path)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("chat", help="קובץ ‎.zip‏ או ‎_chat.txt‏")
    p.add_argument("--me", help="השם שלך כפי שהוא מופיע בצ'אט")
    p.add_argument("--month", type=int)
    p.add_argument("--year", type=int)
    p.add_argument("--cap-min", type=float, default=15, help="מרווח מקסימלי בין הודעות באותו רצף (15)")
    p.add_argument("--burst-floor-min", type=float, default=0, help="מינימום דקות לרצף אחד (0)")
    p.add_argument("--keep-smalltalk", action="store_true", help="לספור גם רצפים בלי מילות עבודה")
    p.add_argument("--break-min", type=float, default=45)
    p.add_argument("--deliv-value-h", type=float, default=0, help="שעות לכל סרטון (0 = הצלבה בלבד)")
    p.add_argument("--deliv-merge-min", type=float, default=60)
    p.add_argument("--upload-value-min", type=float, default=20)
    p.add_argument("--upload-merge-min", type=float, default=30)
    p.add_argument("--decisions", help="קובץ decisions.json")
    p.add_argument("--xlsx", help="נתיב לקובץ אקסל לחיוב")
    p.add_argument("--out", default="analysis.json")
    a = p.parse_args()

    msgs = parse_chat(read_chat_text(a.chat))
    if not msgs:
        sys.exit("לא זוהו הודעות. האם זה ייצוא וואטסאפ?")
    if not a.me:
        counts = defaultdict(int)
        for m in msgs:
            counts[m.sender] += 1
        print("שולחים בצ'אט (להריץ שוב עם --me):")
        for s, n in sorted(counts.items(), key=lambda x: -x[1]):
            print(f"  {s}: {n} הודעות")
        return
    if not (a.month and a.year):
        last = msgs[-1].ts
        a.month, a.year = a.month or last.month, a.year or last.year

    res = analyze(msgs, a.me, a.month, a.year, a)
    if a.decisions:
        with open(a.decisions, encoding="utf-8") as f:
            res = apply_decisions(res, json.load(f))
    res["totals"] = totals(res)
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=2, default=str)
    print_report(res, final=bool(a.decisions))
    if a.xlsx:
        write_xlsx(res, a.xlsx)
        print(f"\nנשמר: {a.xlsx}")


if __name__ == "__main__":
    main()
