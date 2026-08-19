"""
עדכון לשונית 'לוח' בקובץ הרייטינג היומי מתוך קובץ Preliminary Programs Report.

שימוש:
    python update_luach.py <source_file.xlsx> <target_file.xlsx>

או ללא ארגומנטים - הסקריפט יחפש בתיקייה של הקובץ הזה את הקובץ העדכני ביותר של כל סוג.
"""
import sys
import os
import re
import glob
import shutil
import io
from datetime import datetime, timedelta
import openpyxl
from copy import copy
from openpyxl.styles import Font, Alignment

# הבטחת פלט UTF-8 בטרמינל של Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# --- מיפויים ---

# שם גיליון בקובץ המקור (של כל ערוץ) => מספר הערוץ בקובץ היעד
SOURCE_SHEET_TO_CHANNEL = {
    "KAN":       11,
    "Keshet 12": 12,
    "Reshet 13": 13,
    "Arutz14":   14,
}

# עמודות בלשונית 'לוח' של קובץ היעד לכל ערוץ: (start_col, end_col, name_col)
# מבנה: לכל ערוץ בלוק של 7 עמודות: התחלה, סיום, שם, TVR, 000s, ריק, אמצע
TARGET_COLS_BY_CHANNEL = {
    11: (17, 18, 19),
    12: (24, 25, 26),
    13: (31, 32, 33),
    14: (38, 39, 40),
}

TARGET_SHEET_NAME = "לוח"
FIRST_DATA_ROW = 2  # שורה 1 היא כותרות

# חלון זמן שבו מתחילה המהדורה המרכזית של הערב (לזיהוי אוטומטי + הדגשה בבולד)
MAIN_NEWS_START = timedelta(hours=19, minutes=30)
MAIN_NEWS_END   = timedelta(hours=20, minutes=15)

# תוכניות שהתחילו לפני השעה הזאת מוגדרות כשידור חוזר (בבוקר מוקדם)
EARLY_RERUN_CUTOFF = timedelta(hours=6, minutes=0)

# סימנים שמזהים תוכנית כשידור חוזר
RERUN_MARKERS = ["ש.ח", "RERUN", "רירן"]

RERUN_LABEL = "שידורים חוזרים"
MAIN_NEWS_LABEL = "מרכזית"

# תוכניות "מוכרות" שלעולם לא מתאחדות ל"שידורים חוזרים" - שמות אחרי כל כללי הניקוי
PROTECTED_SHOWS = {
    # ערוץ 11
    "הבוקר הזה", "קלמן-ליברמן", "קלמן ליברמן", "קובנוביק",
    "גליק ותמר", "העולם היום", "כאן בשש", "חדשות +", "הכוורת",
    "בואו לאכול איתי", "חדשות הלילה",
    # ערוץ 12
    "כותרות הבוקר", "חדשות הבוקר",
    "משדר צהריים", "מהדורת היום",
    "חמש עם", "שש עם", "שבע עם",
    # (מאסטר שף וסברי מרנן לא מוגנים - הגרסה הרירן שלהם באמצע היום מתאחדת לשידורים חוזרים.
    #  הגרסה של מאסטר שף בערב לא נכנסת לחלון האיחוד כי היא אחרי 15:00.)
    "גיא פינס",
    "אברי ושרקי", "אברי ושרקי +", "מאחורי הקלעים", "אולפן שישי",
    # ערוץ 13
    "רגע לפני", "העולם הבוקר", "פותחים יום",
    'הדו"ח היומי', "חדשות היום", "שיחת היום",
    "מוריה וברקו", "אזור בחירה",
    "הצינור", "היום שהיה",
    # ערוץ 14
    "ישראל הבוקר", "ישראל כותרות",
    "מהדורת בוקר", "חדשת צהריים", "יומן צהריים",
    "אולפן פתוח", "חמש", "ריקלין ושות'", "ברדוגו ושלזינגר",
    "הפטריוטים", "משדר מיוחד", "גירסת יוש", "פתחי את שי",
    # כל הערוצים
    MAIN_NEWS_LABEL, RERUN_LABEL,
}

# חלון זמן שבו רצפים של תוכניות לא-מוכרות מתאחדות לשידורים חוזרים
MIDDAY_START = timedelta(hours=11, minutes=0)
MIDDAY_END = timedelta(hours=15, minutes=0)
# ערוצים שבהם גם הלילה המאוחר מתאחד (23:30+)
LATE_NIGHT_START = timedelta(hours=23, minutes=30)
LATE_NIGHT_CONSOLIDATE_CHANNELS = {13, 14}

# קידומות של קודים פנימיים של הערוץ שמייצגים תוכניות אמיתיות.
# כל שם שמתחיל בקידומת (למשל 'A187850_1', 'A187850_2'...) יומר לשם התוכנית.
# אחרי ההמרה - האיחוד יתפוס אותם כאותה תוכנית וישלב לשורה אחת.
CODED_SHOW_PREFIXES = {
    "A187850": "גיא פינס",   # קשת 12 - שידור חי של לילה טוב עם גיא פינס
}

# תיקרה מקסימלית לזמנים - האחרון יכול להיות 25:59 (לא 26:00)
MAX_END = timedelta(hours=25, minutes=59)

# קיצורי שמות תוכניות (השם המנוקה בצד שמאל -> השם המקוצר בצד ימין)
# מבוסס על הרפרנס של הלוח היומי
NAME_SHORTCUTS = {
    # ערוץ 11 - KAN
    "כאן חדשות +":                              "חדשות +",
    # ערוץ 12 - Keshet
    "חדשות הבוקר עם ניב רסקין":                 "חדשות הבוקר",
    "חדשות הבוקר עם נסלי ויואב":                "חדשות הבוקר",
    "מאסטר שף נבחרת החלומות":                    "מאסטר שף",
    "12 בצוהריים":                              "משדר צהריים",
    "מהדורת היום עם עמליה ועופר":                "מהדורת היום",
    "חמש עם רפי רשף":                           "חמש עם",
    "שש עם..":                                  "שש עם",
    "שבע עם קרן מרציאנו":                       "שבע עם",
    "מאחורי הקלעים - גבעה":                     "מאחורי הקלעים",
    "אולפן שישי הכתבות":                        "אולפן שישי",
    "לילה טוב עם גיא פינס":                     "גיא פינס",
    "אברי ושרקי פלוס אחד":                      "אברי ושרקי +",
    # ערוץ 11 - כותרות משנה של תוכניות
    "זהב שחור - מתוך הזוהמה":                   "זהב שחור",
    # ערוץ 14
    "7.10 גירסת יוש":                           "גירסת יוש",
    # ערוץ 11 - שינויי כתיב לפי הרפרנס
    "קלמן ליברמן":                              "קלמן-ליברמן",
    # ערוץ 14 - Arutz 14
    "מהדורת הבוקר עם דנה ורון":                  "מהדורת בוקר",
    "חדשות צהריים":                             "חדשת צהריים",
    "האולפן הפתוח עם בועז גולן":                "אולפן פתוח",
    # ("חמש" של ערוץ 14 נשמר כמו במקור - בלי "עם")
    "שבע עם שלזינגר וברדוגו":                   "ברדוגו ושלזינגר",
    "הפטריוטים בהגשת ינון מגל":                  "הפטריוטים",
    "משדר מיוחד פריימריז בליכוד":                "משדר מיוחד",
}

# מיפוי שמות תוכניות רשת 13 מתעתיק אנגלי לעברית
# (אחרי clean_program_name - בלי מספרי פרק/עונה)
RESHET_TRANSLATIONS = {
    # בוקר
    "HAOLAM HABOKER REGA LEFNEY": "רגע לפני",
    "HAOLAM HABOKER":             "העולם הבוקר",
    "POTHIM YOM":                 "פותחים יום",
    # חדשות ואקטואליה
    "HAZINOR":                    "הצינור",
    "HATZINOR":                   "הצינור",
    "HADOCH HAYOMI":              'הדו"ח היומי',
    "CHADSHOT HAYOM":             "חדשות היום",
    "SICHAT HAYOM":               "שיחת היום",
    "MORIA & BERKO":              "מוריה וברקו",
    "MORIA&BERKO":                "מוריה וברקו",
    "EZOR BECHIRA":               "אזור בחירה",
    "MAHADURA":                   "מהדורה",
    "HAYOMSHEHAYA":               "היום שהיה",
    "HAMITZAD HAMISHPATI":        "המצעד המשפטי",
    # בידור
    "POWER COUPLE":               "פאוור קאפל",
    "MAKTUB":                     "מכתוב",
    "MILON HAYOFI":               "מילון היופי",
    "THE VOICE":                  "דה וויס",
    "BLACK SPACE":                "בלאק ספייס",
    # לילה / קטעים
    "MEAHOREI HAKESEF":           "מאחורי הכסף",
    "KATAVOT SHIHSI":             "כתבות שישי",
    "NANA DISK":                  "ננה דיסק",
    "STATOSKOP":                  "סטטוסקופ",
    "CLIPTIME":                   "קליפטיים",
    "KLIP NIYAR YAPANI":          "קליפ ניר יפני",
    "HAMOHIM":                    "המוחים",
    "ANSHEI HANADLAN":            'אנשי הנדל"ן',
}

# משך מינימלי לתוכנית - כל תוכנית קצרה מזה (למשל מבזק) לא תיכלל
MIN_DURATION = timedelta(minutes=10)

# הלוח מתחיל מ-06:00 - כל תוכנית שנגמרת עד 06:00 (כולל) יוצאת
MORNING_START_CUTOFF = timedelta(hours=6, minutes=0)


def parse_time(val):
    """הופך ערך מהקובץ המקור לאובייקט timedelta."""
    if val is None or val == "":
        return None
    if isinstance(val, timedelta):
        return val
    if isinstance(val, datetime):
        return timedelta(hours=val.hour, minutes=val.minute, seconds=val.second)
    if isinstance(val, str):
        parts = val.strip().split(":")
        if len(parts) == 2:
            h, m = parts
            return timedelta(hours=int(h), minutes=int(m))
        if len(parts) == 3:
            h, m, s = parts
            return timedelta(hours=int(h), minutes=int(m), seconds=int(s))
    return None


def clean_program_name(name):
    """
    מנקה שם תוכנית - מסיר מספרי פרק/עונה/מזהי מהדורה וכו', כדי שיוצג רק השם.
    דוגמאות:
      'הבוקר הזה 25-26 514'                          -> 'הבוקר הזה'
      'ההמצאות של המחר עונה 17 - פרק 9'              -> 'ההמצאות של המחר'
      'MORIA & BERKO #233'                            -> 'MORIA & BERKO'
      'החדשות עם מגי טביבי (#155)'                    -> 'החדשות עם מגי טביבי'
      'מאסטר שף – עונה 2, פרק 1   RERUN'              -> 'מאסטר שף'
      'כאן חדשות + 11'                                -> 'כאן חדשות +'
    """
    if not name:
        return name
    s = str(name).strip()
    if s == RERUN_LABEL:
        return s

    # מפעילים את הכללים בלולאה כדי לתפוס ניקויים שרשוריים
    for _ in range(6):
        prev = s
        # (1) הסרת סוגריים בסוף: (#155), (7 ימים), (RERUN)
        s = re.sub(r'\s*\([^)]*\)\s*$', '', s)
        # (2) הסרת "#" וכל מה שאחריו
        s = re.sub(r'\s*#.*$', '', s)
        # (3) הסרת "RERUN" בסוף
        s = re.sub(r'\s+RERUN\s*$', '', s, flags=re.IGNORECASE)
        # (4) הסרת "ש.ח" בסוף
        s = re.sub(r'\s+ש\.?\s*ח\s*$', '', s)
        # (5) הסרת " - פרק X..." או " – עונה X..."
        s = re.sub(r'\s*[-–]\s*(פרק|עונה)\b.*$', '', s)
        # (6) הסרת " פרק X" או " עונה X" בלי מקף
        s = re.sub(r'\s+(פרק|עונה)\s+\S+.*$', '', s)
        # (7) הסרת תאריך פורמט "DD.M" או "DD.MM" בסוף (כמו "16.8")
        s = re.sub(r'\s+\d{1,2}\.\d{1,2}(\s+\d+)*\s*$', '', s)
        # (7ב) הסרת מספרי אירוע/עונה בסוף: " 25-26 514", " 2026", " 66"
        s = re.sub(r'\s+\d+([-/]\d+)*(\s+\d+)*\s*$', '', s)
        # (8) הסרת " - X" בסוף כשה-X קצר (סימן שהשם נחתך)
        s = re.sub(r'\s*[-–]\s*\S{1,3}\s*$', '', s)
        # (9) ניקוי מקפים/פסיקים/סוגריים בסוף
        s = re.sub(r'[\s,(\-–]+$', '', s)
        if s == prev:
            break

    return s.strip()


def shorten_name(name):
    """מקצר שם תוכנית לפי מילון NAME_SHORTCUTS או כללי הסתעפות בסוף."""
    if not name or name == RERUN_LABEL:
        return name
    s = str(name).strip()
    # מיפוי ישיר
    if s in NAME_SHORTCUTS:
        return NAME_SHORTCUTS[s]
    # שם עם ' - ' באמצע: מקצץ את החלק שאחרי המקף רק אם הוא ארוך (3+ מילים).
    # דוגמאות:
    #   'זהב שחור - מי רצח את ההיפ הופ...'  -> 'זהב שחור'  (6 מילים אחרי -> חתך)
    #   'הפשוטע - הלוקוס הלבן'               -> 'הפשוטע - הלוקוס הלבן'  (2 מילים -> נשאר)
    # דרישה: רווח לפני המקף כדי לא לפגוע ב-'קלמן-ליברמן'
    m = re.match(r'^(.+?)\s+[-–]\s+(.+)$', s)
    if m:
        head = m.group(1).strip()
        tail = m.group(2).strip()
        tail_word_count = len(tail.split())
        if tail_word_count >= 3:
            s = head
    return s


def consolidate_unrecognized_in_window(programs, window_start, window_end=None, ignore_protection=False):
    """
    מאחדת רצפים של 2+ תוכניות שאינן בPROTECTED_SHOWS בחלון זמן נתון.
    מיועד לתוכניות אמצע-יום (בין הבוקר לחדשות היום) או לילה מאוחר.
    window_end=None -> ללא סוף (מתאים לאיחוד לילה).
    ignore_protection=True -> אף תוכנית לא מוגנת (חוץ מהמהדורה המרכזית ושידורים חוזרים
                              קיימים) - מתאים ללילה שבו הכל בעצם רירן.
    """
    def in_window(td):
        if td is None: return False
        if window_end is None:
            return td >= window_start
        return window_start <= td < window_end

    def is_protected(name):
        if not name: return True
        s = str(name).strip()
        if ignore_protection:
            # בלילה - רק המהדורה המרכזית והשידורים החוזרים עצמם 'מוגנים'
            return s in (MAIN_NEWS_LABEL, RERUN_LABEL)
        if s in PROTECTED_SHOWS:
            return True
        return False

    result = []
    i = 0
    while i < len(programs):
        start, end, name = programs[i]
        # מתחילים לבדוק רק אם התוכנית בתוך החלון ולא מוגנת
        if start is not None and in_window(start) and not is_protected(name):
            group = [(start, end, name)]
            j = i + 1
            while j < len(programs):
                s2, e2, n2 = programs[j]
                if s2 is not None and in_window(s2) and not is_protected(n2):
                    group.append((s2, e2, n2))
                    j += 1
                else:
                    break
            if len(group) >= 2:
                g_start = group[0][0]
                g_end = group[-1][1] if group[-1][1] is not None else group[-1][0]
                result.append((g_start, g_end, RERUN_LABEL))
                i = j
                continue
        result.append((start, end, name))
        i += 1
    return result


def rename_main_news(programs):
    """התוכנית שנופלת בחלון המהדורה המרכזית (19:30-20:15) משנה שם ל-'מרכזית'."""
    result = []
    for start, end, name in programs:
        if is_main_news(start):
            name = MAIN_NEWS_LABEL
        result.append((start, end, name))
    return result


# תוכנית "ישראל הבוקר" בערוץ 14 שמתחילה לפני 07:00 היא בעצם "ישראל כותרות"
# (מקדים קצר של כותרות לפני משדר הבוקר המלא בשעה 07:04).
ISRAEL_HEADLINES_CUTOFF = timedelta(hours=7, minutes=0)


def rename_israel_headlines(programs):
    """
    בערוץ 14: משנה את השם של 'ישראל הבוקר' שמתחיל לפני 07:00 ל'ישראל כותרות'.
    כך הפעימה המוקדמת (06:04-07:00) תוצג בנפרד מ'ישראל הבוקר' של 07:04-09:32.
    """
    result = []
    for start, end, name in programs:
        if (name and str(name).strip() == "ישראל הבוקר"
                and start is not None and start < ISRAEL_HEADLINES_CUTOFF):
            name = "ישראל כותרות"
        result.append((start, end, name))
    return result


def translate_reshet_name(clean_name):
    """מתרגם שם תוכנית של רשת 13 מתעתיק אנגלי לעברית לפי מילון RESHET_TRANSLATIONS."""
    if not clean_name or clean_name == RERUN_LABEL:
        return clean_name
    upper = clean_name.upper().strip()
    # התאמה מדויקת
    if upper in RESHET_TRANSLATIONS:
        return RESHET_TRANSLATIONS[upper]
    # התאמה חלקית - שם המיפוי מופיע כמחרוזת משנה בשם המנוקה
    # ממויין לפי אורך יורד כדי לתפוס תחילה מפתחות ספציפיים ("HAOLAM HABOKER REGA LEFNEY" לפני "HAOLAM HABOKER")
    for key in sorted(RESHET_TRANSLATIONS.keys(), key=len, reverse=True):
        if key in upper:
            return RESHET_TRANSLATIONS[key]
    # אם לא נמצא תרגום - משאירים את השם המקורי (באנגלית)
    return clean_name


def adjust_consecutive_endings(programs):
    """
    לוגיקה: אם שתי תוכניות רצופות כמעט בלי הפסקה,
    הסיום של הקודמת יוגדר לדקה לפני ההתחלה של הבאה
    (למשל: הבאה 11:01 -> הקודמת מסתיימת 11:00).
    לא מאריך תוכנית אם יש פער אמיתי (למשל תוכנית שנחתכה) - רק מקצר סיום.
    """
    result = list(programs)
    for i in range(len(result) - 1):
        s_i, e_i, n_i = result[i]
        start_next = result[i + 1][0]
        if start_next is None or e_i is None:
            continue
        new_end = start_next - timedelta(minutes=1)
        # לא מאריכים סיום - רק מקצרים כשהתוכנית הבאה 'רודפת' לפני הסיום המקורי
        if new_end >= e_i:
            continue
        # בטיחות: אל תיצור סיום לפני התחלה
        if s_i is not None and new_end < s_i:
            continue
        result[i] = (s_i, new_end, n_i)
    return result


def round_to_minute(td):
    """מעגל timedelta לדקה הקרובה."""
    if td is None:
        return None
    total_s = td.total_seconds()
    minutes = round(total_s / 60)
    return timedelta(minutes=minutes)


def floor_end_minute(td):
    """
    מחשב שעת סיום להצגה: מפחית 30 שניות ואז מקצץ לדקה.
    הכלל: אם השניות בסיום >= 30 -> הדקה הנוכחית;
           אם < 30 -> הדקה הקודמת (הרגע הזה עדיין 'נגמר' מהדקה הקודמת).
    דוגמאות:
      07:08:44 -> 07:08 (44 >= 30)
      22:51:00 -> 22:50 (0 < 30)
      12:00:03 -> 11:59 (3 < 30)
      10:59:39 -> 10:59 (39 >= 30)
      13:03:22 -> 13:02 (22 < 30)
    """
    if td is None:
        return None
    total_s = int(td.total_seconds()) - 30
    if total_s < 0:
        return td
    return timedelta(minutes=total_s // 60)


def merge_same_named_adjacent(programs, max_gap_minutes=2):
    """
    מאחדת שורות סמוכות בעלות שם זהה + רווח קטן ביניהן (עד max_gap_minutes דקות).
    זה תופס:
      * תוכנית שנחתכה על ידי מבזק/פרסומת (ישראל הבוקר עם 4 דק' רווח - יישאר 2 שורות)
      * חדשות הבוקר עם שני מגישים ברצף (0 דק' רווח - יתאחד)
    """
    result = []
    for start, end, name in programs:
        if result and name and str(result[-1][2] or "").strip() == str(name).strip():
            prev_start, prev_end, prev_name = result[-1]
            gap_ok = True
            if prev_end is not None and start is not None:
                gap_sec = (start - prev_end).total_seconds()
                if gap_sec > max_gap_minutes * 60:
                    gap_ok = False
            if gap_ok:
                result[-1] = (prev_start, end, prev_name)
                continue
        result.append((start, end, name))
    return result


def filter_pre_morning(programs):
    """
    מסיר תוכניות טרום-בוקר:
     * אם התוכנית מסתיימת עד 06:00 - מסולקת
     * אם התוכנית מתחילה לפני 06:00 ומסתיימת לפני 07:00 (שאריות לילה
       שחוצות את 06:00 בכמה שניות) - גם מסולקת
    תוכנית שמתחילה לפני 06:00 ומתמשכת עד 07:00+ (תוכנית בוקר אמיתית) - נשארת.
    """
    morning_real_start = timedelta(hours=7, minutes=0)
    result = []
    for start, end, name in programs:
        if end is not None and end <= MORNING_START_CUTOFF:
            continue
        if (start is not None and end is not None
                and start < MORNING_START_CUTOFF
                and end < morning_real_start):
            continue
        result.append((start, end, name))
    return result


def apply_max_end_rules(programs):
    """
    חוקי הגבול העליון של הלוח (25:59):
     - תוכנית שמתחילה בשעה 26:00 או אחריה - מסולקת
     - תוכנית שמסתיימת אחרי 26:00 - מסולקת (מתמשכת מעבר לגבול)
     - תוכנית שמסתיימת בדיוק ב-26:00 - הסיום נקצץ ל-25:59
    """
    hr26 = timedelta(hours=26)
    result = []
    for start, end, name in programs:
        if start is not None and start >= hr26:
            continue
        if end is not None:
            if end > hr26:
                continue
            if end == hr26:
                end = MAX_END
        result.append((start, end, name))
    return result


def is_rerun(name, start):
    """מזהה אם תוכנית היא שידור חוזר: לפי סימון בשם או לפי שעה מוקדמת."""
    if start is not None and start < EARLY_RERUN_CUTOFF:
        return True
    if name:
        n = str(name).upper()
        for marker in RERUN_MARKERS:
            if marker.upper() in n:
                return True
    return False


def is_main_news(start):
    """האם התוכנית היא המהדורה המרכזית (מתחילה בחלון 19:30-20:15)."""
    if start is None:
        return False
    # אנחנו משווים למודולו-יום כדי לתמוך בשעות שידור אחרי-חצות
    seconds_in_day = start.total_seconds() % 86400
    start_of_day = timedelta(seconds=seconds_in_day)
    return MAIN_NEWS_START <= start_of_day <= MAIN_NEWS_END


def resolve_coded_names(programs):
    """
    ממיר שמות של קודים פנימיים (למשל 'A187850_1') לשם התוכנית האמיתית
    לפי CODED_SHOW_PREFIXES. חייב לרוץ לפני filter_short כדי שהקטעים הקצרים
    יקבלו את השם הנכון וייאחדו לפני שיסולקו.
    """
    result = []
    for start, end, name in programs:
        if name:
            s = str(name).strip()
            for prefix, real_name in CODED_SHOW_PREFIXES.items():
                if s.startswith(prefix):
                    name = real_name
                    break
        result.append((start, end, name))
    return result


def filter_short_programs(programs):
    """מסננת תוכניות קצרות מ-10 דקות (מבזקים וכו')."""
    result = []
    for start, end, name in programs:
        if start is not None and end is not None:
            duration = end - start
            # תוכניות שחוצות חצות - הסיום מוצג כ-00:35 בעוד ההתחלה 22:29,
            # אז נקבל משך שלילי; מוסיפים 24 שעות כדי לחשב את המשך הנכון.
            if duration.total_seconds() < 0:
                duration += timedelta(hours=24)
            if duration < MIN_DURATION:
                continue
        result.append((start, end, name))
    return result


def consolidate_reruns(programs):
    """
    מאחדת רצפים של 2+ שידורים חוזרים לשורה אחת בשם 'שידורים חוזרים'.
    שידור חוזר בודד (לא רצוף לעוד אחד) נשאר עם שמו המקורי.
    """
    result = []
    i = 0
    while i < len(programs):
        start, end, name = programs[i]
        if is_rerun(name, start):
            group = [(start, end, name)]
            j = i + 1
            while j < len(programs):
                s2, e2, n2 = programs[j]
                if is_rerun(n2, s2):
                    group.append((s2, e2, n2))
                    j += 1
                else:
                    break
            if len(group) >= 2:
                # רצף של 2+ - מאחדים
                g_start = group[0][0]
                g_end = group[-1][1] if group[-1][1] is not None else group[-1][0]
                result.append((g_start, g_end, RERUN_LABEL))
            else:
                # שידור חוזר בודד - שומרים כמו שהוא
                result.append(group[0])
            i = j
        else:
            result.append((start, end, name))
            i += 1
    return result


def apply_midnight_rollover(programs):
    """
    כשהתוכניות עוברות את חצות, הזמנים בקובץ המקור נראים כאילו חוזרים ל-00:00,
    אבל בפועל צריך להציג אותם כ-24:00, 25:00 וכו' (זמן שידור).
    הפונקציה מוסיפה 24 שעות (או יותר) לכל תוכנית שעברה חצות.
    """
    day_offset = timedelta(hours=0)
    prev_end = None
    result = []
    for start, end, name in programs:
        s, e = start, end
        if s is not None:
            s_adjusted = s + day_offset
            # אם זמן ההתחלה 'המוזז' עדיין לפני זמן הסיום הקודם - חצינו את חצות
            if prev_end is not None and s_adjusted + timedelta(minutes=1) < prev_end:
                day_offset += timedelta(hours=24)
                s_adjusted = s + day_offset
            s = s_adjusted
        if e is not None:
            e_adjusted = e + day_offset
            # אם הסיום קטן מההתחלה - התוכנית עצמה חוצה חצות
            if s is not None and e_adjusted < s:
                e_adjusted += timedelta(hours=24)
            e = e_adjusted
        result.append((s, e, name))
        if e is not None:
            prev_end = e
        elif s is not None:
            prev_end = s
    return result


def read_programs_from_source(source_path):
    """קורא את כל התוכניות מכל 4 גיליונות הערוצים."""
    wb = openpyxl.load_workbook(source_path, data_only=True)
    channels = {}  # channel_number -> list of (start, end, name)

    for sheet_name, channel_num in SOURCE_SHEET_TO_CHANNEL.items():
        if sheet_name not in wb.sheetnames:
            print(f"  [!] גיליון '{sheet_name}' לא נמצא בקובץ המקור")
            channels[channel_num] = []
            continue

        ws = wb[sheet_name]
        programs = []
        # מבנה עמודות: A=StartTime B=EndTime C=Duration D=ID E=Name
        for row in range(2, ws.max_row + 1):
            start = parse_time(ws.cell(row=row, column=1).value)
            end   = parse_time(ws.cell(row=row, column=2).value)
            name  = ws.cell(row=row, column=5).value

            if start is None and end is None and not name:
                continue
            programs.append((start, end, name))

        # 0. המרת קודים פנימיים לשם תוכנית אמיתי (למשל A187850_* -> גיא פינס)
        programs = resolve_coded_names(programs)

        # 0ב. איחוד ראשוני של קטעים סמוכים עם אותו שם, לפני סינון תוכניות קצרות.
        # מונע שקטעים של 'גיא פינס' (כל אחד ~7 דק') ייעלמו כי הם קצרים מ-10 דק'.
        programs = merge_same_named_adjacent(programs, max_gap_minutes=10)

        # 1. סינון תוכניות קצרות מ-10 דקות (מבזקים)
        n0 = len(programs)
        programs = filter_short_programs(programs)
        n_filtered_short = n0 - len(programs)

        # 2. שכתוב שעות אחרי חצות: 00:00->24:00, 01:00->25:00 וכו'
        programs = apply_midnight_rollover(programs)

        # 3. סינון תוכניות טרום-בוקר: הכל שנגמר עד 06:00 יוצא
        n1 = len(programs)
        programs = filter_pre_morning(programs)
        n_filtered_early = n1 - len(programs)

        # 4. חוקי גבול עליון 25:59
        programs = apply_max_end_rules(programs)

        # 5. איחוד ראשוני: שורות סמוכות עם שם מקורי זהה + רווח עד 10 דק'
        #    (מטפל ב-'POWER COUPLE 4 #14' שנחתך על ידי מבזק ונשאר עם רווח 7 דק')
        n_before_merge1 = len(programs)
        programs = merge_same_named_adjacent(programs, max_gap_minutes=10)
        n_merged1 = n_before_merge1 - len(programs)

        # 6. איחוד רצפי שידורים חוזרים (משתמש בשמות המקוריים לזיהוי ש.ח / RERUN)
        n2 = len(programs)
        programs = consolidate_reruns(programs)
        n_consolidated = n2 - len(programs)

        # 6. ניקוי שמות תוכניות (הסרת מספרי פרק/עונה/מזהי מהדורה)
        programs = [(s, e, clean_program_name(n)) for s, e, n in programs]

        # 7. תרגום שמות רשת 13 מתעתיק אנגלי לעברית
        if channel_num == 13:
            programs = [(s, e, translate_reshet_name(n)) for s, e, n in programs]

        # 8. קיצור שמות (מילון + הסרה של ' - כותרת משנה')
        programs = [(s, e, shorten_name(n)) for s, e, n in programs]

        # 9. שינוי שם המהדורה המרכזית ל-'מרכזית' (לפי חלון הזמן 19:30-20:15)
        programs = rename_main_news(programs)

        # 9ב. ערוץ 14 בלבד: 'ישראל הבוקר' לפני 07:00 -> 'ישראל כותרות'
        if channel_num == 14:
            programs = rename_israel_headlines(programs)

        # 10. איחוד שורות סמוכות עם שם זהה + רווח קטן (עד 2 דק')
        n_before_merge = len(programs)
        programs = merge_same_named_adjacent(programs)
        n_merged = n_before_merge - len(programs)

        # 11. עיגול זמנים:
        #     - התחלה: עיגול לדקה הקרובה
        #     - סיום: רצפה לדקה הגמורה (07:08:44 -> 07:08; 22:51:00 -> 22:50)
        programs = [(round_to_minute(s), floor_end_minute(e), n) for s, e, n in programs]

        # 12. איחוד תוכניות לא-מוכרות באמצע היום (11:00-15:00) לשידורים חוזרים
        #     (חייב לרוץ אחרי העיגול כדי שהזמנים יתאימו לגבולות החלון)
        programs = consolidate_unrecognized_in_window(programs, MIDDAY_START, MIDDAY_END)

        # 13. איחוד תוכניות לא-מוכרות בלילה המאוחר (רק לערוצים מסוימים).
        # ignore_protection=True: בלילה, גם תוכניות "מוכרות" (למשל הדו"ח היומי) מתאחדות
        # כי הן בעצם רירן שאיננה חלק מהחדשות של אותו יום.
        if channel_num in LATE_NIGHT_CONSOLIDATE_CHANNELS:
            programs = consolidate_unrecognized_in_window(
                programs, LATE_NIGHT_START, ignore_protection=True
            )

        # 14. הסרת שורת "זנב" קצרה מאוד בסוף היום (< 15 דק') - למשל 25:48→25:59 שהוא סתם ממלא
        if programs:
            last_s, last_e, last_n = programs[-1]
            if (last_s is not None and last_e is not None
                    and last_n not in (RERUN_LABEL, MAIN_NEWS_LABEL)
                    and (last_e - last_s) < timedelta(minutes=15)
                    and last_s >= timedelta(hours=25)):
                programs = programs[:-1]

        # 12. סיום של כל תוכנית = דקה לפני התחלה של התוכנית הבאה
        programs = adjust_consecutive_endings(programs)

        # 13. חוקי גבול עליון שוב אחרי עיגול (למקרה שרנד יצר 26:00)
        programs = apply_max_end_rules(programs)

        channels[channel_num] = programs
        print(f"  [OK] {sheet_name} -> ערוץ {channel_num}: "
              f"{len(programs)} שורות "
              f"(סוננו: {n_filtered_short} קצרות + {n_filtered_early} טרום-בוקר, "
              f"אוחדו {n_consolidated} חוזרים + {n_merged} כפילויות)")

    return channels


def clear_channel_columns(ws, cols, max_row):
    """מנקה את התוכן (בלי לגעת בפורמט/נוסחאות אחרות) בעמודות start/end/name של ערוץ."""
    start_col, end_col, name_col = cols
    for row in range(FIRST_DATA_ROW, max_row + 1):
        for col in (start_col, end_col, name_col):
            ws.cell(row=row, column=col).value = None


def write_programs(ws, cols, programs):
    """כותב תוכניות ללשונית - התחלה/סיום/שם.
    מיישרת את כל השורות באותו עיצוב (פונט + יישור) של שורה 2 - כדי שהכל יראה אחיד.
    מדגישה בבולד את המהדורה המרכזית."""
    start_col, end_col, name_col = cols

    # שולפים את העיצוב "המנחה" משורה 2 לכל אחת מ-3 העמודות (התחלה/סיום/שם)
    template_fonts = {
        col: copy(ws.cell(row=FIRST_DATA_ROW, column=col).font)
        for col in (start_col, end_col, name_col)
    }
    template_alignments = {
        col: copy(ws.cell(row=FIRST_DATA_ROW, column=col).alignment)
        for col in (start_col, end_col, name_col)
    }

    def apply_uniform_style(cell, col, bold=False):
        # פונט - משכפל מהתבנית ומוסיף בולד אם צריך
        base_font = template_fonts[col]
        cell.font = Font(
            name=base_font.name,
            size=base_font.size,
            color=base_font.color,
            bold=bold,
            italic=base_font.italic,
        )
        # יישור - כמו התבנית
        cell.alignment = copy(template_alignments[col])

    for i, (start, end, name) in enumerate(programs):
        row = FIRST_DATA_ROW + i
        if start is not None:
            c = ws.cell(row=row, column=start_col)
            c.value = start
            c.number_format = "[h]:mm"
            apply_uniform_style(c, start_col)
        if end is not None:
            c = ws.cell(row=row, column=end_col)
            c.value = end
            c.number_format = "[h]:mm"
            apply_uniform_style(c, end_col)
        if name is not None:
            c = ws.cell(row=row, column=name_col)
            c.value = name
            # בולד למהדורה המרכזית
            apply_uniform_style(c, name_col, bold=is_main_news(start))


def find_target_sheet(wb):
    """מוצא את לשונית 'לוח' (מזהה גם וריאציות)."""
    for name in wb.sheetnames:
        if name.strip() == TARGET_SHEET_NAME or "לוח" in name:
            return wb[name]
    raise ValueError(f"לא נמצאה לשונית בשם '{TARGET_SHEET_NAME}' בקובץ היעד")


def backup_file(path):
    """יוצר גיבוי של הקובץ לפני שינוי."""
    base, ext = os.path.splitext(path)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{base}__backup_{ts}{ext}"
    shutil.copy2(path, backup_path)
    return backup_path


def update_luach(source_path, target_path, skip_backup=False):
    print(f"\n[קריאה] מקור: {os.path.basename(source_path)}")
    programs_by_channel = read_programs_from_source(source_path)

    if not skip_backup:
        print(f"\n[גיבוי] יוצר גיבוי לקובץ היעד...")
        backup = backup_file(target_path)
        print(f"   גיבוי נשמר ב: {os.path.basename(backup)}")

    print(f"\n[עדכון] פותח קובץ יעד: {os.path.basename(target_path)}")
    wb = openpyxl.load_workbook(target_path)
    ws = find_target_sheet(wb)
    max_row = max(ws.max_row, 60)

    for channel, cols in TARGET_COLS_BY_CHANNEL.items():
        programs = programs_by_channel.get(channel, [])
        print(f"  ערוץ {channel}: מנקה שורות ישנות ומכניס {len(programs)} תוכניות")
        clear_channel_columns(ws, cols, max_row)
        write_programs(ws, cols, programs)

    print(f"\n[שמירה] שומר קובץ...")
    wb.save(target_path)
    print(f"\n[סיום] הקובץ עודכן: {target_path}")


def _find_newest(folders, name_filter):
    """מחפש בכל התיקיות (לפי סדר) את הקובץ העדכני שעונה לתנאי."""
    candidates = []
    for folder in folders:
        if not folder or not os.path.isdir(folder):
            continue
        for p in glob.glob(os.path.join(folder, "*.xlsx")):
            base = os.path.basename(p)
            if base.startswith("~$"):  # קובץ פתוח באקסל
                continue
            if "__backup_" in base or "_מעודכן" in base:
                continue
            if name_filter(base):
                candidates.append(p)
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


def _get_real_downloads_folder():
    """
    מחזיר את הנתיב האמיתי של תיקיית ההורדות של המשתמש.
    לפעמים תיקיית ההורדות מועברת לדיסק אחר (למשל D:\\Users\\user\\Downloads
    במקום C:\\Users\\user\\Downloads). מחלץ את הנתיב האמיתי דרך Shell API של Windows.
    """
    # נסיון ראשון: שאילתה ל-Windows Shell (הכי אמין)
    try:
        import subprocess
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command",
             "(New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path"],
            capture_output=True, text=True, timeout=5
        )
        path = result.stdout.strip()
        if path and os.path.isdir(path):
            return path
    except Exception:
        pass

    # נסיון שני: הרישום של Windows
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        path, _ = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")
        winreg.CloseKey(key)
        if path and os.path.isdir(path):
            return path
    except Exception:
        pass

    # ברירת מחדל: תיקיית ההורדות של הפרופיל
    return os.path.join(os.path.expanduser("~"), "Downloads")


def auto_find_files(base_folder):
    """
    מוצא אוטומטית את קובץ המקור (Preliminary) וקובץ היעד (רייטינג).
    מחפש קודם בתיקיית inbox, אחרי כן בתיקיית הבסיס, ולבסוף בתיקיית ההורדות
    האמיתית של המשתמש (גם אם הועברה לדיסק אחר).
    """
    inbox = os.path.join(base_folder, "inbox")
    downloads = _get_real_downloads_folder()
    search_folders = [inbox, base_folder, downloads]

    source = _find_newest(search_folders,
                          lambda name: name.lower().startswith("preliminaryprogramsreport"))
    target = _find_newest(search_folders,
                          lambda name: not name.lower().startswith("preliminaryprogramsreport"))

    if not source:
        raise FileNotFoundError(
            "לא נמצא קובץ מקור (PreliminaryProgramsReport*.xlsx). "
            f"חפשנו ב: inbox, תיקיית הפרויקט, ותיקיית ההורדות."
        )
    if not target:
        raise FileNotFoundError(
            "לא נמצא קובץ יעד (הקובץ עם לשונית לוח). "
            f"חפשנו ב: inbox, תיקיית הפרויקט, ותיקיית ההורדות."
        )
    return source, target


def open_file_windows(path):
    """פותח את הקובץ במערכת ההפעלה (Excel)."""
    try:
        os.startfile(path)
    except Exception as e:
        print(f"   [!] לא הצלחתי לפתוח אוטומטית: {e}")


def main():
    if len(sys.argv) == 3:
        source_path, target_path = sys.argv[1], sys.argv[2]
        auto_mode = False
    else:
        default_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        print(f"[חיפוש] מחפש קבצים אוטומטית...")
        source_path, target_path = auto_find_files(default_folder)
        print(f"   מקור: {source_path}")
        print(f"   יעד:  {target_path}")
        auto_mode = True

    # במצב אוטומטי: יוצרים עותק בתיקיית results ולא נוגעים במקור
    if auto_mode:
        default_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        results_dir = os.path.join(default_folder, "results")
        os.makedirs(results_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(target_path))[0]
        # מציאת מספר גרסה עוקב (V1, V2, V3...) כדי להבדיל מהקובץ שהמשתמשת העלתה
        version = 1
        while True:
            candidate = os.path.join(results_dir, f"{base}_V{version}.xlsx")
            if not os.path.exists(candidate):
                output_path = candidate
                break
            version += 1
        shutil.copy2(target_path, output_path)
        print(f"\n[העתקה] יצרתי עותק לעדכון: {output_path}")
        target_path = output_path

    # במצב אוטומטי - לא צריך גיבוי (כבר יש עותק בתיקיית results)
    update_luach(source_path, target_path, skip_backup=auto_mode)

    # פותח את הקובץ אוטומטית באקסל
    if auto_mode:
        print(f"\n[פתיחה] פותח את הקובץ באקסל...")
        open_file_windows(target_path)


if __name__ == "__main__":
    main()
