#!/usr/bin/env python3
"""
מפיק דוח HTML יומי מקבצי Excel של גדודים.

שימוש:
    python run.py <תיקייה>                          # היום
    python run.py <תיקייה> --date 2026-05-27         # תאריך ספציפי
    python run.py <תיקייה> --date 2026-05-27 --out custom.html

מבנה תיקיית הקלט:
    הסקריפט סורק את התיקייה ומחפש קבצים בפורמט:
        <מספר>_export*.xlsx         (לדוגמה: 240_export.xlsx, 241_export_v2.xlsx)
        <שם-עברי>_export*.xlsx      (לדוגמה: מפקדה_מחוז_export.xlsx)
"""

import argparse
import json
import re
import sys
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import Counter, defaultdict
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# ============== טעינת קונפיגורציה ==============

SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
TEMPLATE_PATH = SCRIPT_DIR / "template.html"
OUTPUT_DIR = SCRIPT_DIR / "reports"


def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


# ============== ניקוי נתונים ==============

def normalize_name(name):
    """מנקה רווחים מיותרים משמות אנשי צוות"""
    if pd.isna(name):
        return None
    cleaned = re.sub(r"\s+", " ", str(name).strip())
    return cleaned if cleaned else None


def parse_date(val):
    """ממיר ערכי תאריך מסוגים שונים ל-date object"""
    if pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def format_unit_label(unit):
    """תווית תצוגה ליחידה.

    - מספר (כמו "240")        → "גדוד 240"
    - שם בעברית ("מפקדה_מחוז") → "מפקדה מחוז"  (אנדרסקור→רווח)
    """
    s = str(unit)
    if not s:
        return s
    return f"גדוד {s}" if s[0].isdigit() else s.replace("_", " ")


# ============== קריאת הקבצים ==============

def discover_files(folder):
    """מוצא את כל הקבצים בתיקייה ומחלץ שם יחידה משם הקובץ.

    תומך בשני פורמטים:
    - מספר גדוד:  240_export.xlsx          → "240"
    - שם בעברית:  מפקדה_מחוז_export.xlsx   → "מפקדה_מחוז"
    """
    folder = Path(folder)
    if not folder.is_dir():
        raise FileNotFoundError(f"תיקייה לא קיימת: {folder}")

    # ֐-׿ = טווח Unicode של אותיות עבריות
    pattern = re.compile(
        r"^(\d+|[֐-׿]+(?:_[֐-׿]+)*).*?_export",
        re.IGNORECASE,
    )
    files = []
    for f in folder.glob("*.xlsx"):
        if f.name.startswith("~$"):  # קבצי lock של Excel
            continue
        m = pattern.search(f.name)
        if m:
            files.append((m.group(1), f))

    if not files:
        raise FileNotFoundError(
            f"לא נמצאו קבצים בפורמט <מספר>_export*.xlsx או <שם-עברי>_export*.xlsx "
            f"בתיקייה: {folder}"
        )

    return sorted(files, key=lambda x: x[0])


def load_battalion(unit_id, filepath):
    """טוען קובץ אחד ומחזיר DataFrame מנוקה"""
    df = pd.read_excel(filepath)
    df["_unit"] = unit_id
    df["_source_file"] = filepath.name
    
    # נרמול שם איש הצוות
    if "מי יצרה קשר" in df.columns:
        df["_staff"] = df["מי יצרה קשר"].apply(normalize_name)
    else:
        df["_staff"] = None
    
    # נרמול תאריך
    if "תאריך" in df.columns:
        df["_date"] = df["תאריך"].apply(parse_date)
    else:
        df["_date"] = None
    
    return df


# ============== חישובי הדוח ==============

def filter_by_date(df, target_date):
    return df[df["_date"] == target_date].copy()


def filter_by_date_range(df, start_date, end_date):
    return df[(df["_date"] >= start_date) & (df["_date"] <= end_date)].copy()


def compute_staff_breakdown(df, status_mapping):
    """מחשב פילוח לפי איש צוות"""
    df = df.copy()
    df["_status_norm"] = df["סטטוס פנייה"].map(status_mapping).fillna("אחר")
    
    result = []
    for (staff, unit), group in df.groupby(["_staff", "_unit"], dropna=False):
        if pd.isna(staff) or not staff:
            continue
        status_counts = group["_status_norm"].value_counts().to_dict()
        result.append({
            "name": staff,
            "unit": unit,
            "total": len(group),
            "handled": status_counts.get("טופל", 0),
            "no_answer": status_counts.get("לא ענה", 0),
            "pending": status_counts.get("ממתין", 0),
            "other": status_counts.get("אחר", 0),
        })
    
    return sorted(result, key=lambda r: r["total"], reverse=True)


def compute_topic_breakdown(df, topics_config, active_values, status_mapping):
    """מחשב כמה פעמים כל נושא הופיע — גלובלי, לפי איש צוות, לפי סטטוס ולפי גדוד"""
    global_counts = Counter()
    by_staff = defaultdict(Counter)
    by_category = Counter()
    by_topic_status = defaultdict(Counter)
    by_topic_unit = defaultdict(Counter)
    units_seen = set()

    active_set = set(active_values)

    df = df.copy()
    df["_status_norm"] = df["סטטוס פנייה"].map(status_mapping).fillna("אחר")

    for col_key, meta in topics_config.items():
        if col_key not in df.columns:
            continue
        for _, row in df.iterrows():
            val = row[col_key]
            if pd.isna(val):
                continue
            if str(val).strip() in active_set:
                label = meta["label"]
                global_counts[label] += 1
                by_category[meta["category"]] += 1
                staff = row.get("_staff")
                if staff:
                    by_staff[staff][label] += 1
                status = row.get("_status_norm", "אחר")
                by_topic_status[label][status] += 1
                unit = row.get("_unit")
                if unit:
                    by_topic_unit[label][unit] += 1
                    units_seen.add(unit)

    return {
        "global":           dict(global_counts.most_common()),
        "by_category":      dict(by_category.most_common()),
        "by_staff":         {k: dict(v.most_common()) for k, v in by_staff.items()},
        "by_topic_status":  {k: dict(v) for k, v in by_topic_status.items()},
        "by_topic_unit":    {k: dict(v) for k, v in by_topic_unit.items()},
        "units":            sorted(units_seen),
    }


def compute_unit_summary(df, status_mapping):
    """סיכום לפי גדוד"""
    df = df.copy()
    df["_status_norm"] = df["סטטוס פנייה"].map(status_mapping).fillna("אחר")

    result = []
    for unit, group in df.groupby("_unit"):
        status_counts = group["_status_norm"].value_counts().to_dict()
        result.append({
            "unit": unit,
            "total": len(group),
            "handled": status_counts.get("טופל", 0),
            "no_answer": status_counts.get("לא ענה", 0),
            "pending": status_counts.get("ממתין", 0),
            "other": status_counts.get("אחר", 0),
        })
    return sorted(result, key=lambda r: r["unit"])


HEBREW_MONTHS = [
    "ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
    "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר",
]


def compute_time_series(df, start, end):
    """מחשב פילוח פניות לאורך הזמן.

    בוחר אוטומטית את גודל הדלי:
        טווח עד 14 ימים   → דלי יומי
        טווח עד 120 ימים  → דלי שבועי (החל מיום ראשון)
        טווח מעל 120 ימים → דלי חודשי

    מחזיר (buckets, bucket_kind) כאשר buckets = רשימת (label, count)
    ו-bucket_kind ∈ {"day", "week", "month"}.
    """
    days_span = (end - start).days + 1
    dates = [d for d in df["_date"].tolist() if d is not None]

    if days_span <= 14:
        counts = Counter(dates)
        out = []
        cur = start
        while cur <= end:
            out.append((cur.strftime("%d/%m"), counts.get(cur, 0)))
            cur += timedelta(days=1)
        return out, "day"

    if days_span <= 120:
        # יום ראשון = תחילת השבוע (בישראל). weekday(): שני=0..ראשון=6
        def week_start_of(d):
            return d - timedelta(days=(d.weekday() + 1) % 7)

        counts = Counter(week_start_of(d) for d in dates)
        out = []
        cur = week_start_of(start)
        while cur <= end:
            week_end = cur + timedelta(days=6)
            label = f"{cur.strftime('%d/%m')}–{week_end.strftime('%d/%m')}"
            out.append((label, counts.get(cur, 0)))
            cur += timedelta(days=7)
        return out, "week"

    counts = Counter((d.year, d.month) for d in dates)
    out = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        label = f"{HEBREW_MONTHS[m - 1]} {y}"
        out.append((label, counts.get((y, m), 0)))
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out, "month"


# ============== ייצור ה-HTML ==============

def format_hebrew_date(d):
    """מעצב תאריך לעברית: יום בשבוע + תאריך"""
    days = ["שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת", "ראשון"]
    months = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
              "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]
    return f"יום {days[d.weekday()]}, {d.day} ב{months[d.month-1]} {d.year}"


def build_staff_stacked(staff_breakdown):
    """בונה bars אופקיים מוערמים — חלוקת סטטוסים לכל איש צוות"""
    if not staff_breakdown:
        return '<div class="empty">אין נתוני אנשי צוות להצגה</div>'

    rows = []
    for s in staff_breakdown:
        total = s["total"]
        if total == 0:
            continue

        segs = []
        for key, cls, label in [
            ("handled",   "good",    "טופל"),
            ("no_answer", "bad",     "לא ענה"),
            ("pending",   "neutral", "ממתין"),
            ("other",     "other",   "אחר"),
        ]:
            count = s[key]
            if not count:
                continue
            pct = (count / total) * 100
            segs.append(
                f'<div class="seg {cls}" style="width:{pct:.2f}%" '
                f'title="{label}: {count}"><span>{count}</span></div>'
            )

        rows.append(f"""
        <div class="stack-row">
          <div class="stack-name">{s['name']}</div>
          <div class="stack-bar">{''.join(segs)}</div>
          <div class="stack-total">{total}</div>
        </div>""")

    return "".join(rows)


def build_topic_status_stacks(by_topic_status):
    """לכל נושא — סרגל מוערם של חלוקת סטטוסים (אותו דפוס כמו staff_stacks)"""
    if not by_topic_status:
        return '<div class="empty">אין נושאים פעילים להצגה</div>'

    # מיון לפי סה"כ פניות יורד
    ordered = sorted(
        by_topic_status.items(),
        key=lambda kv: sum(kv[1].values()),
        reverse=True,
    )

    rows = []
    for label, status_counts in ordered:
        total = sum(status_counts.values())
        if total == 0:
            continue
        segs = []
        for status_key, cls in [
            ("טופל",   "good"),
            ("לא ענה", "bad"),
            ("ממתין",  "neutral"),
            ("אחר",    "other"),
        ]:
            c = status_counts.get(status_key, 0)
            if not c:
                continue
            pct = (c / total) * 100
            segs.append(
                f'<div class="seg {cls}" style="width:{pct:.2f}%" '
                f'title="{status_key}: {c}"><span>{c}</span></div>'
            )
        rows.append(f"""
        <div class="stack-row">
          <div class="stack-name">{label}</div>
          <div class="stack-bar">{''.join(segs)}</div>
          <div class="stack-total">{total}</div>
        </div>""")

    return "".join(rows)


def build_topic_unit_table(units, by_topic_unit):
    """לכל נושא — חלוקה לפי גדוד, כטבלה"""
    if not by_topic_unit or not units:
        return '<div class="empty">אין נתונים להצגה</div>'

    # מיון לפי סה"כ פניות בנושא יורד
    ordered = sorted(
        by_topic_unit.items(),
        key=lambda kv: sum(kv[1].values()),
        reverse=True,
    )

    header_cells = "".join(
        f'<th class="num">{format_unit_label(u)}</th>' for u in units
    )

    body_rows = []
    for label, per_unit in ordered:
        total = sum(per_unit.values())
        cells = "".join(
            f'<td class="num" data-label="{format_unit_label(u)}">'
            f'{per_unit.get(u, 0) if per_unit.get(u, 0) > 0 else "—"}</td>'
            for u in units
        )
        body_rows.append(
            f'<tr>'
            f'<td class="topic-name" data-label="נושא">{label}</td>'
            f'{cells}'
            f'<td class="num rate" data-label="סה&quot;כ">{total}</td>'
            f'</tr>'
        )

    return f"""
    <table class="data">
      <thead>
        <tr>
          <th>נושא</th>
          {header_cells}
          <th class="num">סה"כ</th>
        </tr>
      </thead>
      <tbody>{''.join(body_rows)}</tbody>
    </table>"""


def build_time_series_section(df, start, end):
    """בונה את כל הסקשן של 'פניות לאורך זמן'.

    מוחזר '' אם הטווח קטן מ-8 ימים — אז ציר זמן לא מוסיף ערך,
    והקופסה הקיימת של KPI כבר מציגה את הסה"כ.
    """
    if (end - start).days + 1 <= 7:
        return ""

    series, kind = compute_time_series(df, start, end)
    if not series or all(c == 0 for _, c in series):
        return ""

    max_val = max(c for _, c in series) or 1
    bars = []
    for label, count in series:
        pct = (count / max_val) * 100
        bars.append(f"""
        <div class="bar-row">
          <div class="bar-label">{label}</div>
          <div class="bar-track"><div class="bar-fill" style="width:{pct:.2f}%"></div></div>
          <div class="bar-value">{count}</div>
        </div>""")

    subtitle_map = {
        "day":   "פניות בכל יום בטווח הדוח",
        "week":  "פניות בכל שבוע (יום ראשון–שבת)",
        "month": "פניות בכל חודש",
    }
    return f"""
  <section>
    <div class="section-header">
      <span class="section-number">§ 06</span>
      <h2>פניות לאורך זמן</h2>
    </div>
    <p class="section-subtitle">{subtitle_map[kind]}</p>
    <div class="topic-bars">{''.join(bars)}</div>
  </section>"""


def build_status_legend(status_mapping):
    """בונה את ה-HTML של מקרא הסטטוסים מתוך מיפוי ה-config"""
    grouped = defaultdict(list)
    for raw, canonical in status_mapping.items():
        grouped[canonical].append(raw)

    descriptions = {
        "טופל":   ("פנייה שטופלה והושלמה",        "good"),
        "לא ענה": ("לא ניתן היה ליצור קשר",       "bad"),
        "ממתין":  ("דורש המשך טיפול בהמשך",       "neutral"),
    }

    items = []
    for canonical in ["טופל", "לא ענה", "ממתין"]:
        if canonical not in grouped:
            continue
        desc, css_class = descriptions.get(canonical, ("", ""))
        sources = " · ".join(f'"{s}"' for s in grouped[canonical])
        items.append(f"""
        <div class="legend-item">
          <div class="legend-label {css_class}">{canonical}</div>
          <div class="legend-desc">{desc}</div>
          <div class="legend-source">{sources}</div>
        </div>""")

    items.append("""
        <div class="legend-item">
          <div class="legend-label">אחר</div>
          <div class="legend-desc">כל סטטוס שלא הוגדר במיפוי</div>
          <div class="legend-source">ערכים שאינם ברשימה מימין</div>
        </div>""")

    return "".join(items)


def render_html(report_period, unit_summary, staff_breakdown, topic_data, total_calls, status_mapping, day_data=None):
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        template = f.read()
    
    # קוביות KPI
    total_handled = sum(u["handled"] for u in unit_summary)
    total_no_answer = sum(u["no_answer"] for u in unit_summary)
    total_pending = sum(u["pending"] for u in unit_summary)
    handle_rate = (total_handled / total_calls * 100) if total_calls else 0
    
    # טבלת גדודים
    unit_rows = ""
    for u in unit_summary:
        rate = (u["handled"] / u["total"] * 100) if u["total"] else 0
        unit_rows += f"""
        <tr>
          <td class="unit-name" data-label="יחידה">{format_unit_label(u['unit'])}</td>
          <td class="num" data-label="סה&quot;כ">{u['total']}</td>
          <td class="num good" data-label="טופל">{u['handled']}</td>
          <td class="num bad" data-label="לא ענה">{u['no_answer']}</td>
          <td class="num neutral" data-label="ממתין">{u['pending']}</td>
          <td class="num" data-label="אחר">{u['other']}</td>
          <td class="num rate" data-label="אחוז טיפול">{rate:.0f}%</td>
        </tr>"""

    # טבלת אנשי צוות
    staff_rows = ""
    for s in staff_breakdown:
        rate = (s["handled"] / s["total"] * 100) if s["total"] else 0
        staff_rows += f"""
        <tr>
          <td class="staff-name" data-label="שם">{s['name']}</td>
          <td class="unit-tag" data-label="יחידה">{s['unit']}</td>
          <td class="num" data-label="סה&quot;כ">{s['total']}</td>
          <td class="num good" data-label="טופל">{s['handled']}</td>
          <td class="num bad" data-label="לא ענה">{s['no_answer']}</td>
          <td class="num neutral" data-label="ממתין">{s['pending']}</td>
          <td class="num" data-label="אחר">{s['other']}</td>
          <td class="num rate" data-label="אחוז טיפול">{rate:.0f}%</td>
        </tr>"""
    
    # פילוח נושאים - גרף עמודות
    topics_global = topic_data["global"]
    if topics_global:
        max_topic = max(topics_global.values())
        topic_bars = ""
        for label, count in list(topics_global.items())[:15]:
            pct = (count / max_topic) * 100
            topic_bars += f"""
            <div class="bar-row">
              <div class="bar-label">{label}</div>
              <div class="bar-track"><div class="bar-fill" style="width:{pct}%"></div></div>
              <div class="bar-value">{count}</div>
            </div>"""
    else:
        topic_bars = '<div class="empty">לא תועדו נושאים פעילים ביום זה</div>'
    
    # קטגוריות
    cat_data = topic_data["by_category"]
    if cat_data:
        cat_total = sum(cat_data.values())
        cat_rows = ""
        for cat, count in cat_data.items():
            pct = (count / cat_total) * 100
            cat_rows += f"""
            <div class="cat-row">
              <div class="cat-name">{cat}</div>
              <div class="cat-bar"><div class="cat-fill" style="width:{pct}%"></div></div>
              <div class="cat-value">{count} <span class="cat-pct">({pct:.0f}%)</span></div>
            </div>"""
    else:
        cat_rows = '<div class="empty">—</div>'
    
    # מטריצת איש צוות × נושא
    staff_topics = topic_data["by_staff"]
    matrix_html = ""
    if staff_topics and topics_global:
        top_topics = list(topics_global.keys())[:8]
        header_cells = "".join(f'<th class="rot">{t}</th>' for t in top_topics)
        matrix_rows = ""
        for staff_name in sorted(staff_topics.keys()):
            row_data = staff_topics[staff_name]
            cells = ""
            for t in top_topics:
                v = row_data.get(t, 0)
                cls = "heat" if v > 0 else "empty-cell"
                intensity = min(v / 3, 1.0) if v else 0
                style = f"--intensity:{intensity}" if v > 0 else ""
                cells += f'<td class="{cls}" style="{style}">{v if v else ""}</td>'
            matrix_rows += f'<tr><td class="staff-name">{staff_name}</td>{cells}</tr>'
        
        matrix_html = f"""
        <table class="matrix">
          <thead>
            <tr><th>איש צוות</th>{header_cells}</tr>
          </thead>
          <tbody>{matrix_rows}</tbody>
        </table>"""
    else:
        matrix_html = '<div class="empty">אין נתונים להצגת מטריצה</div>'
    
    # הסתגלות לתאריך בודד או טווח תאריכים
    if isinstance(report_period, tuple):
        start_date, end_date = report_period
    else:
        start_date = end_date = report_period

    title_date = (
        format_hebrew_date(start_date)
        if start_date == end_date
        else f"{format_hebrew_date(start_date)} עד {format_hebrew_date(end_date)}"
    )
    iso_date = (
        start_date.isoformat()
        if start_date == end_date
        else f"{start_date.isoformat()}_to_{end_date.isoformat()}"
    )

    # החלפות בתבנית
    html = template
    html = html.replace("{{TITLE_DATE}}", title_date)
    html = html.replace("{{ISO_DATE}}", iso_date)
    html = html.replace("{{GENERATED_AT}}", datetime.now().strftime("%d/%m/%Y %H:%M"))
    units_list = " · ".join(u["unit"] for u in unit_summary) if unit_summary else "—"
    html = html.replace("{{UNITS_LIST}}", units_list)
    html = html.replace("{{TOTAL_CALLS}}", str(total_calls))
    html = html.replace("{{TOTAL_HANDLED}}", str(total_handled))
    html = html.replace("{{TOTAL_NO_ANSWER}}", str(total_no_answer))
    html = html.replace("{{TOTAL_PENDING}}", str(total_pending))
    html = html.replace("{{HANDLE_RATE}}", f"{handle_rate:.0f}")
    html = html.replace("{{UNIT_ROWS}}", unit_rows)
    html = html.replace("{{STAFF_ROWS}}", staff_rows)
    html = html.replace("{{TOPIC_BARS}}", topic_bars)
    html = html.replace("{{CATEGORY_ROWS}}", cat_rows)
    html = html.replace("{{STAFF_TOPIC_MATRIX}}", matrix_html)
    html = html.replace("{{STATUS_LEGEND}}", build_status_legend(status_mapping))
    html = html.replace("{{STAFF_STACKED}}", build_staff_stacked(staff_breakdown))
    html = html.replace(
        "{{TOPIC_STATUS_STACKS}}",
        build_topic_status_stacks(topic_data.get("by_topic_status", {}))
    )
    html = html.replace(
        "{{TOPIC_UNIT_TABLE}}",
        build_topic_unit_table(
            topic_data.get("units", []),
            topic_data.get("by_topic_unit", {})
        )
    )
    time_series_html = (
        build_time_series_section(day_data, start_date, end_date)
        if day_data is not None and len(day_data) > 0
        else ""
    )
    html = html.replace("{{TIME_SERIES_SECTION}}", time_series_html)

    return html


# ============== main ==============

def main():
    parser = argparse.ArgumentParser(description="הפקת דוח יומי מקבצי גדודים")
    parser.add_argument("folder", help="תיקייה עם קבצי האקסל")
    parser.add_argument("--date", help="תאריך הדוח (YYYY-MM-DD). ברירת מחדל: היום")
    parser.add_argument("--out", help="שם קובץ פלט (ברירת מחדל: reports/report_<תאריך>.html)")
    args = parser.parse_args()
    
    # תאריך יעד
    if args.date:
        try:
            target = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print(f"שגיאה: תאריך לא תקין: {args.date}. השתמש בפורמט YYYY-MM-DD")
            sys.exit(1)
    else:
        target = date.today()
    
    print(f"מפיק דוח עבור: {target}")
    
    config = load_config()
    
    # טעינת קבצים
    files = discover_files(args.folder)
    print(f"נמצאו {len(files)} קבצים: {', '.join(u for u, _ in files)}")
    
    dfs = []
    for unit_id, filepath in files:
        df = load_battalion(unit_id, filepath)
        dfs.append(df)
        print(f"  {format_unit_label(unit_id)}: {len(df)} רשומות")
    
    all_data = pd.concat(dfs, ignore_index=True)
    
    # סינון לפי תאריך
    day_data = filter_by_date(all_data, target)
    print(f"פניות ב-{target}: {len(day_data)}")
    
    if len(day_data) == 0:
        print(f"\nאזהרה: אין פניות בתאריך {target}. הדוח יופק עם נתונים ריקים.")
    
    # חישובים
    unit_summary = compute_unit_summary(day_data, config["status_mapping"])
    staff_breakdown = compute_staff_breakdown(day_data, config["status_mapping"])
    topic_data = compute_topic_breakdown(
        day_data, config["topics"], config["active_values"], config["status_mapping"]
    )
    
    # ייצור HTML
    html = render_html(
        target, unit_summary, staff_breakdown, topic_data, len(day_data),
        config["status_mapping"], day_data=day_data,
    )
    
    # שמירה
    OUTPUT_DIR.mkdir(exist_ok=True)
    if args.out:
        out_path = Path(args.out)
    else:
        out_path = OUTPUT_DIR / f"report_{target.isoformat()}.html"
    
    out_path.write_text(html, encoding="utf-8")
    print(f"\n✓ הדוח נשמר: {out_path.absolute()}")


if __name__ == "__main__":
    main()
