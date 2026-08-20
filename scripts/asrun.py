"""
עיבוד קובץ AsRun של i24 (ערוץ שלנו) — מחקה את המאקרו ImportAndOrganizeAsRun.
מקבל קובץ TXT ומפיק שתי טבלאות:
  * תוכניות: (id, start, end, title) — עם איחוד Media/Switch pairs וכל שורה עוקבת עם אותו ID+כותרת
  * פערים: (prev_end, cur_start) עבור פערים של 1-240 דקות בין אירועים עוקבים
"""
from datetime import timedelta

# מיקום שדות בשורה (0-indexed, מקביל ל-VBA Mid שהוא 1-indexed)
POS_START     = (17, 25)   # HH:MM:SS - 8 chars from position 18
POS_END       = (40, 48)   # HH:MM:SS - 8 chars from position 41
POS_ID        = (52, 84)   # 32 chars from position 53
POS_TITLE     = (105, 137) # 32 chars from position 106

# תיקרת זמן: 26:00 (=1 יום + 2 שעות). מעל זה עוצרים או קוטמים.
LIMIT = timedelta(hours=26)
CLIP_END = timedelta(hours=25, minutes=59, seconds=59)

# פערים "רלוונטיים": בין דקה ל-4 שעות
GAP_MIN = timedelta(seconds=60)
GAP_MAX = timedelta(hours=4)


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
