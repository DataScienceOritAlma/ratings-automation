"""
אוטומצית רייטינג - אתר Streamlit לעדכון לשונית 'לוח' מקובץ יומי.
מריץ את הכללים שסיכמנו והופך את הפלט לניתן לעריכה בממשק ולהורדה.
"""
import sys
import os
import io
import shutil
import tempfile
from datetime import datetime, timedelta, date
import streamlit as st
import openpyxl
import pandas as pd

# ייבוא הלוגיקה מהסקריפט הראשי
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(PROJECT_ROOT, "scripts")
sys.path.insert(0, SCRIPTS_DIR)
import update_luach as U  # noqa: E402
import asrun as A  # noqa: E402


# --- הגדרות ---
st.set_page_config(
    page_title="אוטומצית רייטינג",
    page_icon="📺",
    layout="wide",
    initial_sidebar_state="expanded",
)

# עיצוב מותאם RTL + עברית
st.markdown("""
    <style>
    body, .stMarkdown, .stText, .stButton, .stSelectbox, .stFileUploader, .stDataFrame {
        direction: rtl;
        text-align: right;
    }
    .main-header {
        color: #1E3A8A;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        border-radius: 12px;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

FEEDBACK_FILE = os.path.join(PROJECT_ROOT, "webapp", "feedback.md")
DOWNLOADS_DIR = os.path.join(PROJECT_ROOT, "results")
TEMPLATE_TARGET = os.path.join(PROJECT_ROOT, "webapp", "template.xlsx")
CHANNELS = {
    11: ("KAN 11", (17, 18, 19)),
    12: ("קשת 12", (24, 25, 26)),
    13: ("רשת 13", (31, 32, 33)),
    14: ("ערוץ 14", (38, 39, 40)),
}


# --- פונקציות עזר ---
def td_to_hhmm(td):
    """timedelta -> 'HH:MM' string."""
    if td is None:
        return ""
    if isinstance(td, timedelta):
        s = int(td.total_seconds())
        h = s // 3600
        m = (s % 3600) // 60
        return f"{h:02d}:{m:02d}"
    return str(td)


def hhmm_to_td(s):
    """'HH:MM' -> timedelta. מחזיר None לריק."""
    if not s or (isinstance(s, str) and not s.strip()):
        return None
    if isinstance(s, timedelta):
        return s
    if isinstance(s, str):
        parts = s.strip().split(":")
        try:
            if len(parts) == 2:
                return timedelta(hours=int(parts[0]), minutes=int(parts[1]))
            if len(parts) == 3:
                return timedelta(hours=int(parts[0]), minutes=int(parts[1]), seconds=int(parts[2]))
        except ValueError:
            return None
    return None


def process_files(source_path, target_path):
    """מריץ את הפייפליין של update_luach ומחזיר את הנתונים כמילון של DataFrames."""
    programs_by_channel = U.read_programs_from_source(source_path)
    # כלל "התחלה אחרי ברייק פתיחה" - חל רק על ברייקים בתוך 3 דק' מתחילת התוכנית
    breaks_by_channel = U.read_breaks_from_source(source_path)
    for chan in programs_by_channel:
        programs_by_channel[chan] = U.adjust_start_by_opening_break(
            programs_by_channel[chan],
            breaks_by_channel.get(chan, []),
        )
    # בונה DataFrames לתצוגה
    dfs = {}
    for chan, programs in programs_by_channel.items():
        rows = []
        for start, end, name in programs:
            rows.append({
                "התחלה": td_to_hhmm(start),
                "סיום": td_to_hhmm(end),
                "שם התוכנית": name if name else "",
            })
        dfs[chan] = pd.DataFrame(rows)
    return dfs


def apply_edits_to_target(target_path, edited_dfs, output_path,
                          breaks_by_channel=None, asrun_programs=None, asrun_gaps=None,
                          asrun_short=None):
    """מקבל את הקובץ המקורי, מיישם עליו את העריכות מהממשק, ושומר לפלט.
    אם breaks_by_channel סופק - גם מעדכן את לשונית ברייקים.
    אם asrun_programs/gaps סופקו - גם ממלא את i24 (לוח A-C + ברייקים 'ערוץ שלנו').
    אם asrun_short סופק - גם ממלא את הלוח המקוצר של i24 (עמודות J-K-L)."""
    shutil.copy2(target_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = U.find_target_sheet(wb)

    for chan, (chan_name, cols) in CHANNELS.items():
        if chan not in edited_dfs:
            continue
        df = edited_dfs[chan]
        programs = []
        for _, row in df.iterrows():
            s = hhmm_to_td(row.get("התחלה", ""))
            e = hhmm_to_td(row.get("סיום", ""))
            n = row.get("שם התוכנית", "")
            if s is None and e is None and not n:
                continue
            programs.append((s, e, n))
        # ניקוי + כתיבה
        U.clear_channel_columns(ws, cols, max(ws.max_row, 60))
        U.write_programs(ws, cols, programs)

    # ברייקים (אם יש)
    if breaks_by_channel:
        ws_breaks = U.find_breaks_sheet(wb)
        if ws_breaks is not None:
            max_break_row = max(ws_breaks.max_row, 100)
            for chan, cols in U.TARGET_BREAKS_COLS_BY_CHANNEL.items():
                breaks = breaks_by_channel.get(chan, [])
                U.clear_breaks_channel_columns(ws_breaks, cols, max_break_row)
                U.write_breaks(ws_breaks, cols, breaks)

    # AsRun של i24 (אם יש)
    if asrun_programs:
        U.clear_asrun_luach_columns(ws, max(ws.max_row, 100))
        U.write_asrun_programs(ws, asrun_programs)
    if asrun_short:
        U.clear_asrun_luach_short_columns(ws, max(ws.max_row, 100))
        U.write_asrun_short_programs(ws, asrun_short)
    if asrun_gaps:
        ws_breaks = U.find_breaks_sheet(wb)
        if ws_breaks is not None:
            max_break_row = max(ws_breaks.max_row, 100)
            U.clear_breaks_channel_columns(ws_breaks, U.TARGET_ASRUN_BREAKS_COLS, max_break_row)
            U.write_breaks(ws_breaks, U.TARGET_ASRUN_BREAKS_COLS, asrun_gaps)

    wb.save(output_path)


def save_feedback(text):
    """שומר משוב לקובץ עם חותמת זמן."""
    os.makedirs(os.path.dirname(FEEDBACK_FILE), exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(FEEDBACK_FILE, "a", encoding="utf-8") as f:
        f.write(f"\n\n## {ts}\n{text}\n")


def read_feedback():
    if not os.path.exists(FEEDBACK_FILE):
        return "אין עדיין משובים."
    with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
        return f.read()


# --- ממשק המשתמש ---
st.markdown("""
    <div class="main-header">
        <h1>📺 אוטומצית רייטינג</h1>
        <p>עדכון אוטומטי של לשונית "לוח" - העלאה, עריכה, הורדה</p>
    </div>
""", unsafe_allow_html=True)

# --- Sidebar ---
with st.sidebar:
    st.header("⚙️ הגדרות")
    data_date_str = (date.today() - timedelta(days=1)).strftime("%d.%m.%Y")
    st.markdown(f"**תאריך הנתונים (אתמול):** {data_date_str}")
    st.markdown("---")
    st.markdown("### 📖 הוראות שימוש")
    st.markdown("""
    1. העלי את קובץ המקור (PreliminaryProgramsReport)
    2. לחצי על "עבד קבצים"
    3. ערכי את התוצאה בטבלה למטה (אופציונלי)
    4. הורידי את הקובץ המעודכן
    """)
    st.markdown("---")
    st.markdown("### 💬 יש הצעה?")
    st.markdown('כתבי בכרטיסיית "משוב"')


# --- Tabs ---
tab_upload, tab_edit, tab_feedback, tab_rules, tab_downloads = st.tabs([
    "📤 העלאה ועיבוד",
    "✏️ עריכה והורדה",
    "💬 משוב",
    "📋 כללים ומילונים",
    "📁 קבצים שהורדו",
])


# --- Tab 1: העלאה ועיבוד ---
with tab_upload:
    st.subheader("שלב 1: העלאת קבצי מקור")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**קובץ Preliminary (KAN/קשת/רשת/ערוץ 14)**")
        source_file = st.file_uploader(
            "בחרי קובץ PreliminaryProgramsReport",
            type=["xlsx", "xls"],
            key="source_upload",
            help="הקובץ עם הגיליונות KAN / Keshet 12 / Reshet 13 / Arutz14",
        )
    with col2:
        st.markdown("**קובץ AsRun של i24 (ערוץ שלנו) - אופציונלי**")
        asrun_file = st.file_uploader(
            "בחרי קובץ AsRun (TXT)",
            type=["txt"],
            key="asrun_upload",
            help="קובץ ה-log של i24 - יעדכן את עמודות A-C בלוח ואת בלוק 'ערוץ שלנו' בברייקים",
        )

    with st.expander("⚙️ קובץ יעד (אופציונלי - יש תבנית שמורה)", expanded=False):
        st.markdown(
            "כברירת מחדל האפליקציה משתמשת בתבנית שמורה של קובץ הלוח. "
            "אם את רוצה להשתמש בקובץ יעד ספציפי (למשל קובץ יומי אחר), העלי אותו כאן:"
        )
        target_file = st.file_uploader(
            "קובץ יעד עם לשונית 'לוח'",
            type=["xlsx"],
            key="target_upload",
        )

    st.markdown("---")

    if st.button("🚀 עבד קבצים", type="primary", disabled=not source_file):
        with st.spinner("מעבד..."):
            try:
                # שמירה זמנית של הקבצים
                tmpdir = tempfile.mkdtemp()
                src_path = os.path.join(tmpdir, source_file.name)
                target_filename = target_file.name if target_file else "לוח.xlsx"
                tgt_path = os.path.join(tmpdir, target_filename)

                with open(src_path, "wb") as f:
                    f.write(source_file.getbuffer())
                if target_file:
                    with open(tgt_path, "wb") as f:
                        f.write(target_file.getbuffer())
                else:
                    # אין קובץ יעד - משתמשים בתבנית השמורה
                    if not os.path.exists(TEMPLATE_TARGET):
                        st.error(f"לא נמצאה תבנית שמורה: {TEMPLATE_TARGET}")
                        st.stop()
                    shutil.copy2(TEMPLATE_TARGET, tgt_path)

                # המרת .xls ל-.xlsx אם צריך
                if src_path.lower().endswith(".xls"):
                    try:
                        import xlrd
                    except ImportError:
                        st.error("נדרש xlrd לקבצי .xls: `pip install xlrd==1.2.0`")
                        st.stop()
                    xls_wb = xlrd.open_workbook(src_path)
                    new_wb = openpyxl.Workbook()
                    new_wb.remove(new_wb.active)
                    for sn in xls_wb.sheet_names():
                        old_ws = xls_wb.sheet_by_name(sn)
                        new_ws = new_wb.create_sheet(sn)
                        for r in range(old_ws.nrows):
                            for c in range(old_ws.ncols):
                                cell = old_ws.cell(r, c)
                                val = cell.value
                                if cell.ctype == 3:  # date/time
                                    y, mo, d, h, mi, s = xlrd.xldate_as_tuple(val, xls_wb.datemode)
                                    if y == 0 and mo == 0 and d == 0:
                                        val = f"{h:02d}:{mi:02d}:{s:02d}"
                                new_ws.cell(row=r+1, column=c+1).value = val
                    src_path = src_path.replace(".xls", ".xlsx")
                    new_wb.save(src_path)

                # עיבוד
                dfs = process_files(src_path, tgt_path)
                breaks = U.read_breaks_from_source(src_path)

                # AsRun (i24) - אם הועלה
                asrun_programs = []
                asrun_gaps = []
                asrun_short = []
                if asrun_file is not None:
                    asrun_path = os.path.join(tmpdir, asrun_file.name)
                    with open(asrun_path, "wb") as f:
                        f.write(asrun_file.getbuffer())
                    asrun_programs, asrun_gaps = A.parse_asrun_file(asrun_path)
                    asrun_short = A.consolidate_asrun_for_luach(asrun_programs)

                st.session_state["dfs"] = dfs
                st.session_state["breaks"] = breaks
                st.session_state["asrun_programs"] = asrun_programs
                st.session_state["asrun_gaps"] = asrun_gaps
                st.session_state["asrun_short"] = asrun_short
                st.session_state["target_path"] = tgt_path
                st.session_state["source_filename"] = source_file.name
                st.session_state["target_filename"] = target_filename
                st.session_state["tmpdir"] = tmpdir

                st.success(f"✅ עובד בהצלחה! מעבירה אותך ללשונית 'עריכה והורדה'")
                st.balloons()

                # תצוגה מקדימה מהירה
                st.markdown("### תצוגה מקדימה:")
                for chan, (chan_name, _) in CHANNELS.items():
                    st.markdown(f"**{chan_name}** — {len(dfs[chan])} שורות בלוח, {len(breaks.get(chan, []))} ברייקים")
                if asrun_programs or asrun_gaps:
                    st.markdown(f"**ערוץ שלנו (i24)** — {len(asrun_programs)} תוכניות, {len(asrun_gaps)} פערים")

            except Exception as e:
                st.error(f"אירעה שגיאה: {e}")
                import traceback
                st.code(traceback.format_exc())


# --- Tab 2: עריכה והורדה ---
with tab_edit:
    if "dfs" not in st.session_state:
        st.info("👆 בבקשה תעלי ותעבדי את הקבצים בלשונית הראשונה")
    else:
        st.subheader("עריכת הלוח")
        st.markdown("אפשר לערוך ישירות בטבלה. הזמנים בפורמט `HH:MM` (כמו `07:08` או `25:59`).")

        edited_dfs = {}
        for chan, (chan_name, _) in CHANNELS.items():
            with st.expander(f"📺 {chan_name}  ({len(st.session_state['dfs'][chan])} שורות)", expanded=True):
                edited = st.data_editor(
                    st.session_state["dfs"][chan],
                    key=f"editor_{chan}",
                    use_container_width=True,
                    num_rows="dynamic",
                    hide_index=False,
                )
                edited_dfs[chan] = edited

        st.markdown("---")
        st.subheader("הורדה")

        # שם קובץ הפלט - ברירת מחדל: תאריך הנתונים (אתמול) + סיומת V
        # V מתקדם אוטומטית לפי כמה קבצים כבר נשמרו בסשן הנוכחי
        data_date = (date.today() - timedelta(days=1)).strftime("%d.%m.%Y")
        v_num = len(st.session_state.get("downloads", [])) + 1
        default_output = f"{data_date}_V{v_num}.xlsx"
        output_name = st.text_input("שם קובץ הפלט:", value=default_output)

        if st.button("💾 שמור והורד", type="primary"):
            try:
                output_path = os.path.join(st.session_state["tmpdir"], output_name)
                apply_edits_to_target(
                    st.session_state["target_path"],
                    edited_dfs,
                    output_path,
                    breaks_by_channel=st.session_state.get("breaks"),
                    asrun_programs=st.session_state.get("asrun_programs"),
                    asrun_gaps=st.session_state.get("asrun_gaps"),
                    asrun_short=st.session_state.get("asrun_short"),
                )
                with open(output_path, "rb") as f:
                    file_bytes = f.read()

                # שמירה בזיכרון של הסשן הנוכחי (לתצוגה בלשונית "קבצים שהורדו")
                if "downloads" not in st.session_state:
                    st.session_state["downloads"] = []
                st.session_state["downloads"].insert(0, {
                    "name": output_name,
                    "bytes": file_bytes,
                    "when": datetime.now(),
                    "size": len(file_bytes),
                })

                st.download_button(
                    label=f"⬇️ הורד: {output_name}",
                    data=file_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("✅ הקובץ מוכן להורדה! (זמין גם בלשונית 'קבצים שהורדו')")
            except Exception as e:
                st.error(f"שגיאה בשמירה: {e}")


# --- Tab 3: משוב ---
with tab_feedback:
    st.subheader("💬 משוב וכללים חדשים")
    st.markdown("""
    כתבי כאן דברים לשנות/להוסיף — הודעות עוברות למפתחת (Claude).
    לדוגמה:
    - "התוכנית X בערוץ 12 - השם צריך להיות Y"
    - "בערוץ 13 יש קוד חדש: XYZ_1 = שם התוכנית"
    - "בשעה 23:00 יש עוד תוכנית שצריך לזהות כשידור חוזר"
    """)
    feedback_text = st.text_area(
        "המשוב שלך:",
        height=150,
        placeholder="כתבי כאן..."
    )
    col_a, col_b = st.columns([1, 4])
    with col_a:
        if st.button("📝 שמור משוב"):
            if feedback_text.strip():
                save_feedback(feedback_text)
                st.success("✅ נשמר! פתחי צ'אט עם Claude כדי לבצע את השינוי")
            else:
                st.warning("המשוב ריק")

    st.markdown("---")
    st.markdown("### 📚 היסטוריית משובים:")
    st.markdown(read_feedback())


# --- Tab 4: כללים ומילונים ---
with tab_rules:
    st.subheader("📋 המילונים שבסקריפט")
    st.markdown("""
    כאן אפשר לראות (רק לצפייה) את המיפויים שהסקריפט מכיר.
    לשינויים - כתבי בכרטיסיית **משוב**.
    """)

    with st.expander("🎯 תרגומי רשת 13 (מאנגלית לעברית)"):
        df = pd.DataFrame(
            [(k, v) for k, v in U.RESHET_TRANSLATIONS.items()],
            columns=["מקור (אנגלית)", "תרגום (עברית)"],
        )
        st.dataframe(df, use_container_width=True)

    with st.expander("✂️ קיצורי שמות"):
        df = pd.DataFrame(
            [(k, v) for k, v in U.NAME_SHORTCUTS.items()],
            columns=["מקור", "מקוצר"],
        )
        st.dataframe(df, use_container_width=True)

    with st.expander("🔤 קידומות קוד -> שם תוכנית"):
        df = pd.DataFrame(
            [(k, v) for k, v in U.CODED_SHOW_PREFIXES.items()],
            columns=["קידומת", "שם תוכנית"],
        )
        st.dataframe(df, use_container_width=True)

    with st.expander("🛡️ תוכניות מוגנות (לא מתאחדות)"):
        st.write(sorted(U.PROTECTED_SHOWS))


# --- Tab 5: קבצים שהורדו ---
with tab_downloads:
    st.subheader("📁 קבצים שהורדו")
    st.markdown(
        "כל קובץ שתשמרי בלשונית 'עריכה והורדה' יופיע כאן. "
        "הקבצים זמינים כל עוד את בתוך אותו סשן בדפדפן."
    )

    downloads = st.session_state.get("downloads", [])
    if not downloads:
        st.info("עדיין לא נשמרו קבצים. שמרי קובץ בלשונית 'עריכה והורדה' והוא יופיע כאן.")
    else:
        st.caption(f"{len(downloads)} קבצים בסשן הנוכחי")
        for i, item in enumerate(downloads):
            when = item["when"].strftime("%d/%m/%Y %H:%M")
            size_kb = item["size"] / 1024
            with st.container():
                c1, c2, c3, c4 = st.columns([4, 2, 1, 2])
                c1.markdown(f"**{item['name']}**")
                c2.markdown(f"🕐 {when}")
                c3.markdown(f"{size_kb:.0f} KB")
                c4.download_button(
                    label="⬇️ הורד",
                    data=item["bytes"],
                    file_name=item["name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{i}_{item['name']}",
                )
                st.markdown("---")


# Footer
st.markdown("---")
st.markdown(
    '<div style="text-align: center; color: #888; font-size: 0.85em;">'
    'אוטומצית רייטינג • גרסה 1.0 • רץ מקומית במחשב שלך'
    '</div>',
    unsafe_allow_html=True,
)
