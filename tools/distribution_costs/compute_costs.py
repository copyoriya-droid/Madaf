#!/usr/bin/env python3
"""חישוב עלויות הפצה מטבלת מעקב סרטונים.

    python compute_costs.py tracker.xlsx --config prices.json [--sheet "ספטמבר"] [--csv out.csv]

tracker: ייצוא של טבלת המעקב (xlsx עם גיליון לכל חודש, או csv של גיליון אחד).
כל שורה היא סרטון, כל עמודה היא ערוץ הפצה, ותא עם "כן" = הופץ בערוץ.

כללי הספירה והמחירים נמצאים ב-config (ראו prices.example.json).
המחירים האמיתיים לא נשמרים בריפו: שומרים אותם ב-prices.local.json, שמוחרג ב-.gitignore.
"""
import argparse
import csv
import json
import sys
from datetime import date, datetime


def norm(s):
    return " ".join(str(s or "").replace("‏", "").replace("‎", "").split())


def load_sheets(path, only=None):
    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig", newline="") as f:
            return {only or "csv": [list(r) for r in csv.reader(f)]}
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    names = [only] if only else wb.sheetnames
    missing = [n for n in names if n not in wb.sheetnames]
    if missing:
        sys.exit(f"לא נמצא גיליון: {missing}. קיימים: {wb.sheetnames}")
    return {n: [list(r) for r in wb[n].iter_rows(values_only=True)] for n in names}


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = norm(v)
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d.%m.%y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def price_on(ch, when):
    """המחיר התקף ביום ההפצה. בלי תאריך בשורה, או לפני כל התקופות: המחיר הבסיסי."""
    price = ch["price"]
    for period in sorted(ch.get("price_from", []), key=lambda p: p["from"]):
        if when and when >= date.fromisoformat(period["from"]):
            price = period["price"]
    return price


def count_cell(cell, ch, cfg):
    """כמה הפצות לחייב על תא אחד, לפי הכללים שבקונפיג."""
    s = norm(cell)
    if not s:
        return 0
    if any(p in s for p in cfg.get("not_billed_phrases", [])):
        return 0
    if s in cfg.get("no_values", []) or not any(s.startswith(y) for y in cfg.get("yes_prefixes", ["כן"])):
        return 0
    n = 1
    if ch.get("count") == "per_marker":
        n += sum(1 for mk in ch.get("extra_markers", []) if mk in s)
    return n


def find_header(rows, channels):
    for i, r in enumerate(rows[:15]):
        cells = [norm(c) for c in r]
        if sum(1 for c in cells if c in channels) >= 1:
            return i, cells
    return None, None


def compute(rows, cfg):
    channels = {norm(k): v for k, v in cfg["channels"].items()}
    hi, header = find_header(rows, channels)
    if hi is None:
        return None
    cols = {name: j for j, name in enumerate(header) if name in channels}
    date_col = next((j for j, name in enumerate(header) if name == norm(cfg.get("date_column", ""))), None)
    title_col = next((j for j, name in enumerate(header) if name == norm(cfg.get("title_column", ""))), 0)
    stop = [norm(s) for s in cfg.get("stop_at_row_starting_with", [])]

    res = {name: {"count": 0, "cost": 0.0} for name in cols}
    distributed = not_distributed = 0
    unknown_cells = []
    for r in rows[hi + 1:]:
        title = norm(r[title_col]) if title_col < len(r) else ""
        if not title:
            continue
        if any(title.startswith(s) for s in stop):
            break  # למשל מחירון ישן בתחתית הגיליון
        when = to_date(r[date_col]) if date_col is not None and date_col < len(r) else None
        row_total = 0
        for name, j in cols.items():
            cell = r[j] if j < len(r) else None
            n = count_cell(cell, channels[name], cfg)
            s = norm(cell)
            if s and n == 0 and s not in cfg.get("no_values", []) and not any(
                p in s for p in cfg.get("not_billed_phrases", [])
            ):
                unknown_cells.append(f"{title[:30]} / {name}: {s}")
            res[name]["count"] += n
            res[name]["cost"] += n * price_on(channels[name], when)
            row_total += n
        if row_total:
            distributed += 1
        else:
            not_distributed += 1
    return {"channels": res, "videos_distributed": distributed, "videos_not_distributed": not_distributed,
            "unknown_cells": unknown_cells}


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("tracker")
    p.add_argument("--config", required=True)
    p.add_argument("--sheet", help="גיליון אחד בלבד (ברירת מחדל: כולם)")
    p.add_argument("--csv", help="קובץ CSV לייבוא ל-Google Sheets, עם נוסחאות SUM")
    a = p.parse_args()

    with open(a.config, encoding="utf-8") as f:
        cfg = json.load(f)
    months = {}
    for name, rows in load_sheets(a.tracker, a.sheet).items():
        r = compute(rows, cfg)
        if r is None:
            print(f"· דילגתי על '{name}': לא נמצאה שורת כותרות עם שמות ערוצים")
            continue
        months[name] = r

    if not months:
        sys.exit("לא חושב אף חודש.")
    chans = list(next(iter(months.values()))["channels"])
    for m in months.values():
        chans += [c for c in m["channels"] if c not in chans]

    for name, r in months.items():
        total_n = sum(v["count"] for v in r["channels"].values())
        total_c = sum(v["cost"] for v in r["channels"].values())
        print(f"\n=== {name} ===")
        print(f"סרטונים שהופצו: {r['videos_distributed']} · הופקו ולא הופצו: {r['videos_not_distributed']}")
        for c in chans:
            v = r["channels"].get(c)
            if v and v["count"]:
                print(f"  {c}: {v['count']} הפצות · ₪{v['cost']:,.2f}")
        print(f"  סה\"כ: {total_n} הפצות · ₪{total_c:,.2f}")
        if r["unknown_cells"]:
            print("  ⚠️ תאים שלא זוהו כ'כן' או 'לא' (לא נספרו, כדאי לבדוק):")
            for u in r["unknown_cells"]:
                print("    -", u)

    if a.csv:
        names = list(months)
        with open(a.csv, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["ערוץ", "מחיר להפצה"] + [f"הפצות {n}" for n in names] + [f"עלות {n}" for n in names])
            for i, c in enumerate(chans, start=2):
                price = cfg["channels"].get(c, {}).get("price", "")
                counts = [months[n]["channels"].get(c, {}).get("count", 0) for n in names]
                costs = [round(months[n]["channels"].get(c, {}).get("cost", 0), 2) for n in names]
                w.writerow([c, price] + counts + costs)
            last = len(chans) + 1
            cols = [chr(ord("C") + k) for k in range(2 * len(names))]
            w.writerow(["סה\"כ", ""] + [f"=SUM({c}2:{c}{last})" for c in cols])
        print(f"\nנשמר: {a.csv} (לייבוא ל-Google Sheets. בלי שורת sep=)")


if __name__ == "__main__":
    main()
