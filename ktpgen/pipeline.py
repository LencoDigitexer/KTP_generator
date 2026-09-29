"""Общий «пайплайн» генератора КТП для CLI и GUI.

Никакого ввода-вывода пользователя — только функции, которые возвращают
структурированные результаты (ошибки — через исключения). GUI (Flask) и CLI
(run.py) используют один и тот же код, поэтому поведение не расходится.
"""

from __future__ import annotations

import datetime
import os
import re
from collections import defaultdict

from .calendar_parser import parse_calendar, Calendar
from .doc_model import parse_lessons_from_text
from .doc_writer import extract_doc_text, rewrite_doc_dates
from .scheduler import build_schedule, ScheduleError

WEEKDAY_SHORT = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
WEEKDAY_FULL = [
    "понедельник", "вторник", "среда", "четверг",
    "пятница", "суббота", "воскресенье",
]

WD_ALIASES = {
    0: ("понедельник", "пн", "пон"),
    1: ("вторник", "вт", "тор"),
    2: ("среда", "ср"),
    3: ("четверг", "чт", "четв"),
    4: ("пятница", "пт", "пятн"),
    5: ("суббота", "сб", "субб"),
    6: ("воскресенье", "вс", "воскр"),
}


class InputError(ValueError):
    """Ошибка во входных данных/параметрах пользователя."""


def parse_weekdays(spec: str) -> list[int]:
    """'вт,пт' / '1,5' / 'вторник пятница' -> [1, 5] (0=пн ... 6=вс)."""
    tokens = [t.strip().lower() for t in
              spec.replace(";", ",").replace("\n", ",").split(",") if t.strip()]
    result: list[int] = []
    for t in tokens:
        if t.isdigit():
            w = int(t)
            if not 0 <= w <= 6:
                raise InputError(f"Число дня недели вне диапазона 0..6: {t}")
        else:
            match = [i for i, aliases in WD_ALIASES.items()
                     if any(a == t or a.startswith(t) for a in aliases)]
            if len(match) == 1:
                w = match[0]
            else:
                raise InputError(
                    f"Не понятен день недели: '{t}'. "
                    "Допустимо: пн вт ср чт пт сб вс (или полные названия, или числа 0..6)"
                )
        if w not in result:
            result.append(w)
    return sorted(result)


def parse_start_date(spec: str) -> datetime.date:
    try:
        return datetime.datetime.strptime(spec.strip(), "%d.%m.%Y").date()
    except ValueError:
        raise InputError(f"Дата начала должна быть в формате дд.мм.гггг, получено: '{spec}'")


def load_calendar(path: str, target_year: int | None = None) -> Calendar:
    if not os.path.isfile(path):
        raise InputError(f"Не найден файл календаря: {path}")
    return parse_calendar(path, target_year=target_year)


def load_document(path: str):
    """Читает .doc, возвращает (raw_bytes, KtpDocument)."""
    if not os.path.isfile(path):
        raise InputError(f"Не найден doc-файл: {path}")
    with open(path, "rb") as f:
        data = f.read()
    text = extract_doc_text(data)
    doc = parse_lessons_from_text(text)
    return data, doc


# ---- Календарь часов в формате Word (.doc/.docx): сетка месяцев ----------

WD_ROWS = ["понедельник", "вторник", "среда", "четверг", "пятница",
           "суббота", "воскресенье"]
WD_ABBR = {"пн": 0, "пон": 0, "вт": 1, "втр": 1, "ср": 2, "сред": 2,
           "чт": 3, "чет": 3, "пт": 4, "пят": 4, "сб": 5, "суб": 5,
           "вс": 6, "вос": 6}
MONTHS_RU = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
             "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
DOC_DATE_RE = re.compile(r"\b(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\b")


def _month_order(name: str) -> int | None:
    n = name.lower().strip().rstrip(".")
    for i, m in enumerate(MONTHS_RU):
        if n.startswith(m[:4]):
            return i
    return None


def _weekday_of_line(s: str) -> int | None:
    t = s.lower().strip().rstrip(".:")
    if t in WD_ROWS:
        return WD_ROWS.index(t)
    if t in WD_ABBR:
        return WD_ABBR[t]
    return None


def calendar_from_doc_text(text: str) -> Calendar | None:
    """Пробует собрать календарь из текста Word-документа.

    Поддерживаются два типовых формата «календаря часов»:
      * сетка месяцев (названия месяцев + дни недели + блоки дат по 7 колонок);
      * помесячные списки рабочих дней («Сентябрь 2026», ниже строки вида
        «02.09 — среда, 6 ч.»).
    Возвращает None, если текст на календарь не похож.
    """
    low = text.lower()
    months_found = sum(1 for m in MONTHS_RU if m[:4] in low)
    wd_found = sum(1 for w in WD_ROWS if w in low)
    has_dates = DOC_DATE_RE.search(text) is not None
    if not ((months_found >= 2 and wd_found >= 1) or (wd_found >= 5 and has_dates)):
        return None

    # Год учебного года: явные годы в датах, иначе текущий сентябрьский цикл
    years = sorted({int(y) for _, _, y in DOC_DATE_RE.findall(text) if y})
    if years:
        ty = years[0]
        ty_next = years[-1] if years[-1] > years[0] else years[0] + 1
    else:
        today = datetime.date.today()
        ty = today.year if today.month >= 9 else today.year + 1
        ty_next = ty + 1

    def year_for(month: int) -> int:
        return ty if month >= 9 else ty_next

    working: set[datetime.date] = set()
    holidays: set[datetime.date] = set()
    cur_month: int | None = None
    col = 0
    row_wd: list[int] = []
    row_vals: list[datetime.date] = []

    def add(d: datetime.date, wd: int):
        if d in holidays:
            return
        if wd >= 5 and not (d.month == 12 and d.day == 31):
            return
        working.add(d)

    def flush_row():
        nonlocal col, row_wd, row_vals
        for wd, d in zip(row_wd, row_vals):
            add(d, wd)
        col, row_wd, row_vals = 0, [], []

    for line in text.replace("\x0b", "\r").split("\r"):
        s = line.strip(" \t\x07")
        if not s:
            continue
        wd_i = _weekday_of_line(s)
        if wd_i is not None:
            flush_row()
            col = wd_i
            continue
        mi = _month_order(s) if len(s) <= 14 else None
        if mi is not None:
            flush_row()
            cur_month = mi + 1
            continue
        # даты в строке: дд.мм или дд.мм.гггг
        tokens = [m for m in DOC_DATE_RE.finditer(s)]
        if not tokens or cur_month is None:
            continue
        if len(tokens) >= 7:                 # строка недельной сетки
            flush_row()
            start_col = col
            for k, m in enumerate(tokens):
                dd_, mm_ = int(m.group(1)), int(m.group(2))
                yy = int(m.group(3)) if m.group(3) else year_for(mm_)
                if yy < 1990:
                    yy += 2000
                try:
                    d = datetime.date(yy, mm_, dd_)
                except ValueError:
                    continue
                row_vals.append(d)
                row_wd.append((start_col + len(row_vals) - 1) % 7)
            col = 0
            if len(row_vals) >= 7:
                flush_row()
        else:                                # список: одна дата = один день
            for m in tokens:
                dd_, mm_ = int(m.group(1)), int(m.group(2))
                yy = int(m.group(3)) if m.group(3) else year_for(mm_)
                if yy < 1990:
                    yy += 2000
                try:
                    d = datetime.date(yy, mm_, dd_)
                except ValueError:
                    continue
                add(d, d.weekday())
    flush_row()

    if len(working) < 30:
        return None
    # Красный цвет в Word без рендера недоступен: учитываем типовые
    # нерабочие праздничные дни РФ.
    extra_hol = {(1, 1), (1, 2), (1, 3), (1, 4), (1, 5), (1, 6), (1, 7), (1, 8),
                 (2, 23), (3, 8), (5, 1), (5, 9), (6, 12), (11, 4)}
    for d in list(working):
        if (d.month, d.day) in extra_hol:
            working.discard(d)
            holidays.add(d)
    by_weekday: dict[int, list[datetime.date]] = defaultdict(list)
    for d in sorted(working):
        by_weekday[d.weekday()].append(d)
    return Calendar(working_dates=sorted(working), holidays=holidays,
                    by_weekday=dict(by_weekday), sheet_name="Word-календарь")


def load_calendar_any(path: str) -> tuple[Calendar, str]:
    """Календарь из .xlsx или из Word (.doc/.docx). Возвращает (календарь, тип)."""
    if not os.path.isfile(path):
        raise InputError(f"Не найден файл календаря: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return load_calendar(path), "xlsx"
    if ext == ".doc":
        _, doc = load_document(path)
        with open(path, "rb") as f:
            cal = calendar_from_doc_text(extract_doc_text(f.read()))
        if cal is None:
            raise InputError("В этом Word-файле не найдена сетка календаря "
                             "(месяцы/дни недели/даты). Загрузите .xlsx.")
        return cal, "doc"
    if ext == ".docx":
        cal = calendar_from_doc_text(_docx_plain_text(path))
        if cal is None:
            raise InputError("В этом .docx не найдена сетка календаря. "
                             "Загрузите .xlsx.")
        return cal, "docx"
    raise InputError("Нужен файл календаря .xlsx (или Word .doc/.docx с сеткой)")


def _docx_plain_text(path: str) -> str:
    import zipfile
    from xml.etree import ElementTree as ET
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    parts: list[str] = []
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        for el in root.iter():
            if el.tag == ns + "t":
                parts.append(el.text or "")
            elif el.tag in (ns + "tc", ns + "tr", ns + "p"):
                parts.append("\r")
    return "".join(parts)


def calendar_info(cal: Calendar) -> dict:
    """Сводка по календарю для интерфейса/отчётов (плюс подсказка даты начала)."""
    first, last = cal.working_dates[0], cal.working_dates[-1]
    hol = sorted(d.strftime("%d.%m.%Y") for d in getattr(cal, "holidays", set()))
    return {
        "working": len(cal.working_dates),
        "holidays": hol,
        "range": f"{first:%d.%m.%Y} — {last:%d.%m.%Y}",
        "cal_min": first.isoformat(),
        "cal_max": last.isoformat(),
        "start_suggested": first.isoformat(),
    }


def plan_lessons(doc, calendar: Calendar, weekdays: list[int],
                 start_date: datetime.date):
    """Возвращает (items, rows, warnings). rows — строки занятий из документа."""
    rows = doc.session_rows
    titles = [row.title for row in rows]
    warnings: list[str] = []
    if not titles:
        raise InputError("В документе не распознаны строки занятий (темы).")
    items = build_schedule(titles, calendar, weekdays, start_date)
    missing = [it for it in items if it.date is None]
    if missing:
        warnings.append(
            f"Не хватило рабочих дат для {len(missing)} занятий "
            f"(номера {missing[0].lesson_index}..{missing[-1].lesson_index}). "
            "Добавьте дни недели или начните раньше."
        )
    no_pos = sum(1 for r in rows if not r.date_pos)
    if no_pos:
        warnings.append(f"{no_pos} строк без позиции даты — они не будут изменены.")
    return items, rows, warnings


def apply_to_document(data: bytes, items, rows, out_path: str) -> dict:
    """Записывает даты в .doc. Возвращает статистику + список замен."""
    replacements = []
    changes = []
    for row, it in zip(rows, items):
        if not it.text or not row.date_pos:
            continue
        start, end = row.date_pos
        new_text = it.text
        if len(new_text) != end - start:      # выравниваем длину ("4.09" -> "04.09")
            new_text = new_text.rjust(end - start, "0")
        if row.date and row.date == it.text:
            pass  # дата уже верная — всё равно перезапишем для надёжности
        replacements.append((start, end, new_text))
        changes.append({
            "index": it.lesson_index,
            "title": it.title,
            "old": row.date or "",
            "new": it.text,
        })
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    rewrite_doc_dates(data, replacements, out_path=out_path)
    return {"count": len(replacements), "changes": changes, "out_path": out_path}


def verify_output(out_path: str, items) -> dict:
    """Перечитывает результат и сверяет даты."""
    with open(out_path, "rb") as f:
        check = parse_lessons_from_text(extract_doc_text(f.read()))
    got_rows = check.session_rows
    ok = sum(1 for src, got in zip(items, got_rows) if (got.date or "") == src.text)
    return {"ok": ok, "total": len(items), "rows_match": len(got_rows) == len(items)}


def run_pipeline(doc_path: str, calendar_path: str, weekdays_spec: str,
                 start_spec: str, out_path: str, target_year: int | None = None) -> dict:
    """Полный цикл: файлы -> расписание -> запись -> самопроверка."""
    weekdays = parse_weekdays(weekdays_spec)
    start_date = parse_start_date(start_spec)
    cal, _cal_type = load_calendar_any(calendar_path)
    data, doc = load_document(doc_path)
    items, rows, warnings = plan_lessons(doc, cal, weekdays, start_date)
    res = apply_to_document(data, items, rows, out_path)
    ver = verify_output(out_path, items)
    return {
        "calendar": {
            "working": len(cal.working_dates),
            "holidays": sorted(d.strftime("%d.%m.%Y") for d in getattr(cal, "holidays", set())),
        },
        "lessons": [
            {
                "index": it.lesson_index,
                "title": it.title,
                "date": it.text,
                "weekday": WEEKDAY_SHORT[it.date.weekday()] if it.date else "",
            }
            for it in items
        ],
        "changes": res["changes"],
        "warnings": warnings,
        "verification": ver,
        "out_path": out_path,
    }
