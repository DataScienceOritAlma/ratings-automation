"""
עיבוד קובץ AsRun של i24 (ערוץ שלנו) — מחקה את המאקרו ImportAndOrganizeAsRun.
מקבל קובץ TXT ומפיק שתי טבלאות:
  * תוכניות: (id, start, end, title) — עם איחוד Media/Switch pairs וכל שורה עוקבת עם אותו ID+כותרת
  * פערים: (prev_end, cur_start) עבור פערים של 1-240 דקות בין אירועים עוקבים

בנוסף: קונסולידציה של התוכניות ללוח (עמודות J-L) — כמו של הערוצים האחרים,
עם שמות מקוצרים ומיזוגים לפי כללי i24.
"""
import re
from datetime import timedelta

# מיקום שדות בשורה (0-indexed, מקביל ל-VBA Mid שהוא 1-indexed)
POS_START     = (17, 25)   # HH:MM:SS - 8 chars from position 18
POS_END       = (40, 48)   # HH:MM:SS - 8 chars from position 41
POS_ID        = (52, 84)   # 32 chars from position 53
POS_TITLE     = (105, 137) # 32 chars from position 106

# תיקרת זמן: 26:00 (=1 יום + 2 שעות). מעל זה עוצרים או קוטמים.
LIMIT = timedelta(hours=26)
CLIP_END = timedelta(hours=25, minutes=59, seconds=59)
MAX_END_LUACH = timedelta(hours=25, minutes=59)  # תיקרה לעמודות K (סיום מעוגל)

# פערים "רלוונטיים": בין דקה ל-4 שעות
GAP_MIN = timedelta(seconds=60)
GAP_MAX = timedelta(hours=4)

# --- כללי קונסולידציה של i24 ללוח (עמודות J-L) ---

# משך מינימלי ללוח (< 2 דקות = פילטר, כמו בערוצים האחרים)
MIN_DURATION_LUACH_SECONDS = 120

# כל תוכנית שמתחילה לפני 07:00 = שידור חוזר (ש.ח)
EARLY_RERUN_CUTOFF_I24 = timedelta(hours=7)

# חלון המהדורה המרכזית (זהה לערוצים האחרים)
MAIN_NEWS_START = timedelta(hours=19, minutes=30)
MAIN_NEWS_END   = timedelta(hours=20, minutes=15)
MAIN_NEWS_LABEL = "מרכזית"

# מיפוי שם גולמי (אחרי ניקוי) → שם מקוצר ללוח של i24.
# המפתחות הם השם אחרי הסרת מספרי פרק/עונה/תאריך.
I24_NAME_SHORTCUTS = {
    "הישראלים":                     "הישראלים",
    "לקט המגזין השבת":              "הסיפורים הגדולים",
    "לקט הסיפורים הגדולים":         "הסיפורים הגדולים",
    "מדברים חדשות":                 "מדברים חדשות",
    "מהדורת 15":                    "מהדורת 15",
    "מהדורת חמש":                   "מהדורת חמש",
    "מדד גלבוע":                    "מדד גלבוע",
    "7 שרון גל":                    "שרון גל",
    "שרון גל":                      "שרון גל",
    "המהדורה המרכזית":              "מרכזית",
    "הערביסטים חמישי":              "הערביסטים",
    "הערביסטים ראשון":              "הערביסטים",
    "הערביסטים שני":                "הערביסטים",
    "הערביסטים שלישי":              "הערביסטים",
    "הערביסטים רביעי":              "הערביסטים",
    "הערביסטים שישי":               "הערביסטים",
    "הערביסטים שבת":                "הערביסטים",
    "ערביסטים חמישי":               "הערביסטים",
    "ערביסטים ראשון":               "הערביסטים",
    "ערביסטים שני":                 "הערביסטים",
    "ערביסטים שלישי":               "הערביסטים",
    "ערביסטים רביעי":               "הערביסטים",
    "ערביסטים שישי":                "הערביסטים",
    "ערביסטים שבת":                 "הערביסטים",
    "הערביסטים":                    "הערביסטים",
    "ערביסטים":                     "הערביסטים",
    "חדר החדשות יום ראשון":         "חדר החדשות",
    "חדר החדשות יום שני":           "חדר החדשות",
    "חדר החדשות יום שלישי":         "חדר החדשות",
    "חדר החדשות יום רביעי":         "חדר החדשות",
    "חדר החדשות יום חמישי":         "חדר החדשות",
    "חדר החדשות יום שישי":          "חדר החדשות",
    "חדר החדשות יום שבת":           "חדר החדשות",
    "חדר החדשות":                   "חדר החדשות",
    "HE FLASH":                     "HE FLASH",
    "על הצלחת":                     "על הצלחת",
    "חוק וצדק":                     "חוק וצדק",
    "סרט הטבע הישראלי":             "סרט הטבע הישראלי",
}

# תוכניות שנחשבות "חצאי שעות" - כשהן ברצף בחלון הבוקר-צהריים
HALF_HOUR_NAMES = {"חדר החדשות", "HE FLASH"}
HALF_HOUR_MIN_COUNT = 3          # צריך לפחות 3 בלוקים ברצף כדי להגדיר "חצאי שעות"
HALF_HOUR_MAX_GAP = timedelta(minutes=10)


def parse_hhmmss(s):
    """'HH:MM:SS' -> timedelta או None אם לא חוקי."""
    s = s.strip()
    if not s or len(s) < 8:
        return None
    parts = s.split(":")
    if len(parts) != 3:
        return None
    try:
        h, m, sec = int(parts[0]), int(parts[1]), int(parts[2])
        return timedelta(hours=h, minutes=m, seconds=sec)
    except ValueError:
        return None


def parse_asrun_file(path):
    """
    קורא קובץ AsRun ומחזיר (programs, gaps).
    - programs: רשימת (id, start_td, end_td, title) - עם איחוד רצפים של אותו id+title
    - gaps: רשימת (prev_end_td, cur_start_td) - רק פערים בין דקה ל-4 שעות
    """
    programs = []
    gaps = []

    last_start = timedelta(0)  # מעקב אחרי חצות
    day_offset = timedelta(0)
    prev_end_td = None  # לחישוב פערים - סוף השורה הקודמת שעובדה

    # לצורך המיזוג: פרטי השורה האחרונה שהוספנו לתוכניות
    prev_id = None
    prev_title = None

    with open(path, encoding="utf-8") as f:
        for line in f:
            if "HES" not in line:
                continue
            if "Switch Event" not in line and "Media Event" not in line:
                continue

            st_str = line[POS_START[0]:POS_START[1]]
            en_str = line[POS_END[0]:POS_END[1]]
            st = parse_hhmmss(st_str)
            en = parse_hhmmss(en_str)
            if st is None or en is None:
                continue

            # לוגיקת מעבר חצות: אם השעה קטנה מהשעה הקודמת שראינו, עברנו יום
            candidate = st + day_offset
            if candidate < last_start:
                day_offset += timedelta(hours=24)
                candidate = st + day_offset
            st_td = candidate

            en_td = en + day_offset
            # אם הסיום קטן מההתחלה (התוכנית חוצה חצות בעצמה)
            if en_td < st_td:
                en_td += timedelta(hours=24)

            last_start = st_td

            # תנאי עצירה: הגענו ל-26:00
            if st_td >= LIMIT:
                break

            # קיטום סיום ל-25:59:59 אם חורג מ-26:00
            if en_td > LIMIT:
                en_td = CLIP_END

            mid = line[POS_ID[0]:POS_ID[1]].strip()
            ti  = line[POS_TITLE[0]:POS_TITLE[1]].strip()

            # פערים: מסתמכים על סוף השורה הקודמת (לפני המיזוג)
            if prev_end_td is not None:
                gap = st_td - prev_end_td
                if GAP_MIN < gap < GAP_MAX:
                    gaps.append((prev_end_td, st_td))
            prev_end_td = en_td

            # מיזוג: אם id וגם title זהים לקודמים - רק מרחיבים סיום
            if mid == prev_id and ti == prev_title and programs:
                pid, pst, _pen, pti = programs[-1]
                programs[-1] = (pid, pst, en_td, pti)
            else:
                programs.append((mid, st_td, en_td, ti))
                prev_id = mid
                prev_title = ti

    return programs, gaps


# --- קונסולידציה של תוכניות i24 ללוח (עמודות J-L) ---

def _has_date_suffix(title):
    """מזהה אם שם התוכנית מסתיים בתאריך (DDMMYY, 6 ספרות) - מסמן שידור חוזר."""
    if not title:
        return False
    return bool(re.search(r'\s+\d{6}\s*$', str(title).strip()))


def _clean_asrun_title(title):
    """מנקה שם תוכנית מ-AsRun:
      - מסיר תאריך בסוף (6 ספרות DDMMYY)
      - מסיר מספר פרק בסוף (3+ ספרות)
      - אבל שומר מספרים קצרים (1-2 ספרות) שיכולים להיות חלק מהשם
        למשל 'מהדורת 15' - 15 חלק מהשם, לא מספר פרק.
    """
    if not title:
        return ""
    s = str(title).strip()
    # תאריך בסוף: DDMMYY (6 ספרות) - להריץ פעם אחת
    s = re.sub(r'\s+\d{6}\s*$', '', s).strip()
    # מספר פרק בסוף: 3 ספרות ומעלה (למשל 410, 587). לא נוגעים ב-1-2 ספרות.
    s = re.sub(r'\s+\d{3,}\s*$', '', s).strip()
    # ניקוי סופי של פסיקים/מקפים
    s = re.sub(r'[\s,\-–]+$', '', s).strip()
    return s


def _shorten_i24_name(title):
    """מקצר שם תוכנית של i24 לפי המילון.
    אם השם לא במילון - מחזיר את השם המנוקה כמו שהוא."""
    clean = _clean_asrun_title(title)
    # התאמה מדויקת
    if clean in I24_NAME_SHORTCUTS:
        return I24_NAME_SHORTCUTS[clean]
    # התאמה חלקית: אולי השם המקוצר מופיע כתחילת השם המנוקה
    # (למשל "חדר החדשות יום חמישי" -> חדר החדשות)
    for key in sorted(I24_NAME_SHORTCUTS.keys(), key=len, reverse=True):
        if clean.startswith(key):
            return I24_NAME_SHORTCUTS[key]
    return clean


def _is_main_news_time(start):
    """האם התוכנית מתחילה בחלון המהדורה המרכזית (19:30-20:15)."""
    if start is None:
        return False
    seconds_in_day = start.total_seconds() % 86400
    start_of_day = timedelta(seconds=seconds_in_day)
    return MAIN_NEWS_START <= start_of_day <= MAIN_NEWS_END


def _is_rerun_time(start):
    """
    תוכניות שמתחילות לפני 07:00 (בבוקר מוקדם) או אחרי חצות (>= 24:00)
    נחשבות שידור חוזר.
    """
    if start is None:
        return False
    if start >= timedelta(hours=24):
        return True
    if start < EARLY_RERUN_CUTOFF_I24:
        return True
    return False


def _round_to_minute(td):
    """מעגל timedelta לדקה הקרובה."""
    if td is None:
        return None
    total_s = td.total_seconds()
    minutes = round(total_s / 60)
    return timedelta(minutes=minutes)


def _floor_end_minute(td):
    """מקצץ סיום: אם השניות ≥ 30 → הדקה הנוכחית, אם < 30 → הדקה הקודמת."""
    if td is None:
        return None
    total_s = int(td.total_seconds()) - 30
    if total_s < 0:
        return td
    return timedelta(minutes=total_s // 60)


def consolidate_asrun_for_luach(programs):
    """
    מקבל רשימת תוכניות גולמיות מ-AsRun ומחזיר רשימת (start, end, name) מקוצרת ללוח.

    שלבים:
      1. סינון תוכניות קצרות מ-2 דקות (מרקרים).
      2. מיזוג: כל 'פילר X' רצוף -> 'פילרים' אחד.
      3. זיהוי רצפי 'חצאי שעות' (3+ בלוקים של חדר החדשות/HE FLASH ברצף) -> 'חצאי שעות'.
      4. קיצור שמות לפי I24_NAME_SHORTCUTS.
      5. מיזוג שורות סמוכות עם שם זהה.
      6. שם המהדורה המרכזית -> 'מרכזית' (לפי חלון 19:30-20:15).
      7. הוספת סיומת 'ש.ח' לתוכניות שמתחילות לפני 07:00 או אחרי חצות.
      8. עיגול זמנים לדקה.
      9. קיטום ל-25:59.
    """
    # שלב 1: סינון פחות מ-2 דקות. (רשומה: (s, e, name, is_rerun) - is_rerun מתעדכן בהמשך)
    filtered = []
    for id_, s, e, t in programs:
        if s is not None and e is not None:
            dur = (e - s).total_seconds()
            if dur < MIN_DURATION_LUACH_SECONDS:
                continue
        filtered.append((s, e, t, False))

    # שלב 2: מיזוג "פילר X" ל"פילרים"
    step2 = []
    for s, e, t, is_rr in filtered:
        name = str(t).strip() if t else ""
        if name.startswith("פילר"):
            new_name = "פילרים"
            if step2 and step2[-1][2] == new_name:
                ps, _pe, pn, prr = step2[-1]
                step2[-1] = (ps, e, pn, prr or is_rr)
                continue
            step2.append((s, e, new_name, is_rr))
        else:
            step2.append((s, e, name, is_rr))

    # שלב 3: זיהוי "חצאי שעות" - רצף של 3+ חדר החדשות/HE FLASH
    step3 = []
    i = 0
    while i < len(step2):
        s, e, name, is_rr = step2[i]
        short = _shorten_i24_name(name)
        if short in HALF_HOUR_NAMES:
            group = [(s, e, name, is_rr)]
            j = i + 1
            while j < len(step2):
                s2, e2, n2, rr2 = step2[j]
                short2 = _shorten_i24_name(n2)
                if short2 not in HALF_HOUR_NAMES:
                    break
                if group[-1][1] is not None and s2 is not None:
                    gap = s2 - group[-1][1]
                    if gap > HALF_HOUR_MAX_GAP:
                        break
                group.append((s2, e2, n2, rr2))
                j += 1
            if len(group) >= HALF_HOUR_MIN_COUNT:
                merged_start = group[0][0]
                merged_end = group[-1][1]
                any_rr = any(g[3] for g in group)
                step3.append((merged_start, merged_end, "חצאי שעות", any_rr))
                i = j
                continue
        step3.append((s, e, name, is_rr))
        i += 1

    # שלב 4: קיצור שמות
    step4 = []
    for s, e, name, is_rr in step3:
        if name == "חצאי שעות" or name == "פילרים":
            step4.append((s, e, name, is_rr))
        else:
            step4.append((s, e, _shorten_i24_name(name), is_rr))

    # שלב 5: מיזוג שורות סמוכות עם שם זהה (עד 5 דקות רווח)
    step5 = []
    max_gap = timedelta(minutes=5)
    for s, e, name, is_rr in step4:
        if step5 and step5[-1][2] == name:
            ps, pe, pn, prr = step5[-1]
            if pe is not None and s is not None and (s - pe) <= max_gap:
                step5[-1] = (ps, e, pn, prr or is_rr)
                continue
        step5.append((s, e, name, is_rr))

    # שלב 6: המהדורה המרכזית -> "מרכזית"
    step6 = []
    for s, e, name, is_rr in step5:
        if _is_main_news_time(s):
            step6.append((s, e, MAIN_NEWS_LABEL, False))  # מרכזית אף פעם לא ש.ח
        else:
            step6.append((s, e, name, is_rr))

    # שלב 7: הוספת ש.ח לפי שלושת הכללים:
    #   (א) התוכנית מתחילה לפני 07:00, או
    #   (ב) אחרי חצות (>= 24:00), או
    #   (ג) זו הופעה שנייה+ של אותה תוכנית באותו יום.
    seen_names = set()
    step7 = []
    for s, e, name, _is_rr in step6:
        # שם בסיס (בלי סיומת ש.ח קיימת)
        base_name = name
        is_second_occurrence = base_name in seen_names
        is_rerun = _is_rerun_time(s) or is_second_occurrence
        seen_names.add(base_name)
        if is_rerun and name != MAIN_NEWS_LABEL and not name.endswith("ש.ח"):
            name = f"{name} ש.ח"
        step7.append((s, e, name))

    # שלב 8: עיגול זמנים לדקה
    step8 = []
    for s, e, name in step7:
        step8.append((_round_to_minute(s), _floor_end_minute(e), name))

    # שלב 9: קיטום ל-25:59
    step9 = []
    for s, e, name in step8:
        if s is not None and s >= timedelta(hours=26):
            continue
        if e is not None and e > MAX_END_LUACH:
            e = MAX_END_LUACH
        step9.append((s, e, name))

    return step9
