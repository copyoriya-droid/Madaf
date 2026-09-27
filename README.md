# מדף: כלים לעבודה חוזרת

סקריפטים שרצים בלי מודל, כלומר בחינם ובאותה תוצאה בכל פעם. אפשר לבקש מ-Claude להריץ אותם, או להריץ לבד.

> ⚠️ **הריפו ציבורי.** צ'אטים, טבלאות ומחירים לא נשמרים כאן. ‎`.gitignore`‏ חוסם אותם.

דרישות: Python 3.9 ומעלה, ‏`pip install openpyxl` (רק לקבצי xlsx).

## ‏`tools/whatsapp_hours`: שעות עבודה מצ'אט וואטסאפ

```bash
# 1. מי השולחים בצ'אט
python tools/whatsapp_hours/analyze_chat.py chat.zip
# 2. ניתוח: זמן אוטומטי + רשימות ממוספרות לאישור
python tools/whatsapp_hours/analyze_chat.py chat.zip --me "אוריה" --month 9 --year 2026
# 3. אחרי שמסמנים מה נכון, מסכמים (ואפשר גם להפיק אקסל)
python tools/whatsapp_hours/analyze_chat.py chat.zip --me "אוריה" --month 9 --year 2026 \
    --decisions decisions.json --xlsx hours.xlsx
```

ברירות המחדל לפי השיטה שנקבעה ביוני 2026: רצף הודעות עם פער של עד 15 דק', וסרטונים לא מחויבים בנפרד (הם נספרים רק להצלבה). כל העלאה היא 20 דק'.

‏`decisions.json` (המספרים הם מספרי המועמדים מהשלב הקודם, והערכים הם שעות):
```json
{
  "scripts": {"1": 1, "3": 1},
  "distribution": {"2": 0.25},
  "calls": {"1": 0.1667},
  "meetings": {"1": 0.5},
  "extra": [{"day": "2026-09-03", "cat": "meetings", "hours": 0.5, "text": "פגישה בבית קפה"}],
  "inbox_hours_per_week": 2
}
```

## ‏`tools/distribution_costs`: עלויות הפצה מטבלת המעקב

```bash
cp tools/distribution_costs/prices.example.json prices.local.json   # ממלאים ערוצים ומחירים אמיתיים
python tools/distribution_costs/compute_costs.py "מעקב סרטונים.xlsx" --config prices.local.json --csv costs.csv
```

- ‏xlsx: מחשב כל גיליון (חודש) בנפרד. ‏`--sheet` מחשב חודש אחד בלבד.
- כללי הספירה מוגדרים בקונפיג:
  - ‏`once`: תא אחד = הפצה אחת, גם אם כתובים בו כמה תתי-ערוצים.
  - ‏`per_marker`: כל סמן נוסף, למשל "וכדאי לדעת", נספר כהפצה נוספת.
  - ‏`not_billed_phrases`: למשל "בלי קשר אלינו", לא מחויב.
  - ‏`price_from`: מחיר חדש מתאריך מסוים. דורש עמודת תאריך.
- תא שלא מזוהה כ"כן" או כ"לא" לא נספר, ומופיע ברשימת אזהרות לבדיקה.
- ‏`costs.csv` נועד לייבוא ל-Google Sheets, עם נוסחאות ‏`SUM` חיות.
