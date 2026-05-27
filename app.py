"""
Streamlit wrapper for the daily reports tool.

Upload Excel files, pick a date, get the same HTML report you'd get from the CLI
(run.py) — but accessible from any browser.

Run locally:  streamlit run app.py
Deploy:       push to a public GitHub repo, connect at share.streamlit.io
"""

import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from run import (
    OUTPUT_DIR,
    compute_staff_breakdown,
    compute_topic_breakdown,
    compute_unit_summary,
    discover_files,
    filter_by_date,
    load_battalion,
    load_config,
    render_html,
)

# ============ עיצוב הדף ============

st.set_page_config(
    page_title="דוח פעילות יומית — מרכז העורף",
    page_icon="📊",
    layout="centered",
)

# RTL לכל ממשק Streamlit
st.markdown(
    """
    <style>
      .stApp { direction: rtl; }
      .stApp h1, .stApp h2, .stApp h3, .stApp p, .stApp label,
      .stApp .stMarkdown, .stApp .stCaption { text-align: right; }
      .stApp [data-testid="stFileUploaderDropzone"] { direction: ltr; }
      .stApp [data-testid="stFileUploaderFileName"] { direction: ltr; }
      div[data-testid="stAlert"] { direction: rtl; text-align: right; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("דוח פעילות יומית — מרכז העורף")
st.caption("העלה את קבצי האקסל של היום, בחר תאריך, וצור את הדוח HTML להורדה ולשיתוף")

# ============ קלט ============

target_date = st.date_input(
    "תאריך הדוח",
    value=date.today(),
    format="DD/MM/YYYY",
)

uploaded = st.file_uploader(
    "קבצי גדודים (xlsx) — הפורמט המצופה: <מספר>_export*.xlsx",
    type=["xlsx"],
    accept_multiple_files=True,
    help="לדוגמה: 240_export.xlsx, 241_export.xlsx. אפשר להעלות כמה שצריך בבת אחת.",
)

if not uploaded:
    st.info("לא הועלו קבצים. בחר לפחות קובץ אקסל אחד כדי להמשיך.")
    st.stop()

# סינון קבצי lock של Excel (~$*.xlsx) — נוצרים כשהקובץ פתוח ב-Excel ולא ניתן להעלות אותם
lock_files = [f.name for f in uploaded if f.name.startswith("~$")]
uploaded = [f for f in uploaded if not f.name.startswith("~$")]

if lock_files:
    st.warning(
        f"דולגים על {len(lock_files)} קבצי lock זמניים של Excel: "
        + ", ".join(lock_files)
        + ". (נוצרים כשהקובץ פתוח ב-Excel — אפשר להתעלם.)"
    )

if not uploaded:
    st.error("כל הקבצים שהועלו הם קבצי lock זמניים. סגור את Excel ובחר את הקבצים האמיתיים.")
    st.stop()

st.success(f"מוכן לעיבוד · {len(uploaded)} קבצים: " + ", ".join(f.name for f in uploaded))

# ============ הפקה ============

if not st.button("צור דוח", type="primary", use_container_width=True):
    st.stop()

config = load_config()

with tempfile.TemporaryDirectory() as tmpdir:
    tmppath = Path(tmpdir)
    for f in uploaded:
        (tmppath / f.name).write_bytes(f.getvalue())

    try:
        files = discover_files(tmppath)
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()

    dfs = []
    with st.spinner("קורא קבצים..."):
        for unit_id, filepath in files:
            df = load_battalion(unit_id, filepath)
            dfs.append(df)

    all_data = pd.concat(dfs, ignore_index=True)
    day_data = filter_by_date(all_data, target_date)

    if len(day_data) == 0:
        st.warning(
            f"לא נמצאו פניות בתאריך {target_date.isoformat()}. "
            "הדוח ייוצר עם נתונים ריקים."
        )

    with st.spinner("מחשב סטטיסטיקות..."):
        unit_summary = compute_unit_summary(day_data, config["status_mapping"])
        staff_breakdown = compute_staff_breakdown(day_data, config["status_mapping"])
        topic_data = compute_topic_breakdown(
            day_data,
            config["topics"],
            config["active_values"],
            config["status_mapping"],
        )

    with st.spinner("מייצר HTML..."):
        html = render_html(
            target_date,
            unit_summary,
            staff_breakdown,
            topic_data,
            len(day_data),
            config["status_mapping"],
        )

# ============ פלט ============

# שמירת הדוח לדיסק תחת reports/ — בדיוק כמו ה-CLI (run.py)
OUTPUT_DIR.mkdir(exist_ok=True)
out_path = OUTPUT_DIR / f"report_{target_date.isoformat()}.html"
out_path.write_text(html, encoding="utf-8")

st.success(
    f"הדוח מוכן · {len(day_data)} פניות · "
    f"{len(files)} יחידות ({', '.join(u for u, _ in files)})"
)
st.caption(f"נשמר אוטומטית: `{out_path}`")

st.download_button(
    label="הורד את הדוח (HTML)",
    data=html.encode("utf-8"),
    file_name=f"report_{target_date.isoformat()}.html",
    mime="text/html",
    use_container_width=True,
)

st.markdown(
    "**שיתוף בווטסאפ:** הורד את הקובץ, ושלח אותו כקובץ מצורף בשיחה. "
    "הוא יוצג כעמוד אינטרנט שאפשר לפתוח בכל מכשיר."
)

with st.expander("תצוגה מקדימה של הדוח"):
    st.components.v1.html(html, height=900, scrolling=True)
