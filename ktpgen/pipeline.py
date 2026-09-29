"""Общий «пайплайн» генератора КТП для CLI и GUI.

Никакого ввода-вывода пользователя — только функции, которые возвращают
структурированные результаты (ошибки — через исключения). GUI (Flask) и CLI
(run.py) используют один и тот же код, поэтому поведение не расходится.
"""

from __future__ import annotations

import datetime
import os

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
    cal = load_calendar(calendar_path, target_year=target_year)
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
