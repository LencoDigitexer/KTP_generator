#!/usr/bin/env python3
"""CLI-генератор КТП: читает .doc и календарь .xlsx, проставляет даты занятий.

Запуск из корня проекта (там, где лежат ktpgen/ и файлы примеров):

    python run.py --doc "КТП ПР1 Робототехника.doc" \
                  --calendar "календарь часов 2026-2027.xlsx" \
                  --weekdays вт,пт --start 01.09.2026 \
                  --out output/КТП_готовый.doc

Без аргументов запускается интерактивный режим с вопросами.
Модули пакета импортируются корректно даже при запуске по пути
(python ktpgen/... / python ./run.py), потому что это обычный скрипт,
а не модуль с относительными импортами.
"""

from __future__ import annotations

import argparse
import datetime
import os
import sys

# Позволяем запускать файл из любой директории: добавляем его папку в sys.path.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ktpgen.calendar_parser import parse_calendar, WEEKDAY_NAMES
from ktpgen.doc_model import parse_lessons_from_text
from ktpgen.doc_writer import extract_doc_text, rewrite_doc_dates
from ktpgen.scheduler import build_schedule, ScheduleError


# Стандартные и краткие названия дней недели (0=пн ... 6=вс)
WD_ALIASES = {
    0: ("понедельник", "пн", "пон"),
    1: ("вторник", "вт", "тор"),
    2: ("среда", "ср"),
    3: ("четверг", "чт", "четв"),
    4: ("пятница", "пт", "пятн"),
    5: ("суббота", "сб", "субб"),
    6: ("воскресенье", "вс", "воскр"),
}


def parse_weekdays(spec: str) -> list[int]:
    """'вт,пт' / '1,5' / 'вторник пятница' -> [1, 5] (0=пн ... 6=вс)."""
    tokens = [t.strip().lower() for t in spec.replace(";", ",").replace(" ", ",").split(",") if t.strip()]
    result: list[int] = []
    for t in tokens:
        if t.isdigit():
            w = int(t)
            if not 0 <= w <= 6:
                raise ValueError(f"Число дня недели вне диапазона 0..6: {t}")
        else:
            match = [i for i, aliases in WD_ALIASES.items()
                     if any(a == t or a.startswith(t) for a in aliases)]
            if len(match) == 1:
                w = match[0]
            else:
                raise ValueError(
                    f"Не понятен день недели: '{t}'. "
                    f"Допустимо: пн вт ср чт пт сб вс (или полные названия, или числа 0..6)"
                )
        if w not in result:
            result.append(w)
    return sorted(result)


def ask(prompt: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    ans = input(f"{prompt}{suffix}: ").strip()
    return ans or (default or "")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Генератор дат в плане КТП")
    p.add_argument("--doc", help="исходный .doc файл плана")
    p.add_argument("--calendar", help="файл календаря часов .xlsx")
    p.add_argument("--weekdays", help="дни недели через запятую, напр. вт,пт или 1,4")
    p.add_argument("--start", help="дата начала занятий дд.мм.гггг, напр. 01.09.2026")
    p.add_argument("--out", default="output/КТП_готовый.doc", help="куда сохранить результат")
    args = p.parse_args(argv)

    doc_path = args.doc or ask("Путь к .doc плану")
    cal_path = args.calendar or ask("Путь к календарю .xlsx")
    wd_spec = args.weekdays or ask("Дни недели (напр. вт,пт)")
    start_spec = args.start or ask("Дата начала занятий дд.мм.гггг")

    for path, label in ((doc_path, "doc-файл"), (cal_path, "календарь")):
        if not os.path.isfile(path):
            print(f"ОШИБКА: не найден {label}: {path}", file=sys.stderr)
            return 2

    try:
        weekdays = parse_weekdays(wd_spec)
        start_date = datetime.datetime.strptime(start_spec, "%d.%m.%Y").date()
    except ValueError as e:
        print(f"ОШИБКА во вводе: {e}", file=sys.stderr)
        return 2

    # 1. Календарь
    cal = parse_calendar(cal_path)
    print(f"Календарь: рабочих дат {len(cal.working_dates)}, "
          f"праздников {len(getattr(cal, 'holidays', []) or [])}")

    # 2. План: извлекаем темы занятий
    with open(doc_path, "rb") as f:
        data = f.read()
    text = extract_doc_text(data)
    doc = parse_lessons_from_text(text)
    rows = doc.session_rows  # только занятия, без заголовков кейсов
    titles = [row.title for row in rows]
    print(f"Найдено занятий: {len(titles)}")
    if not titles:
        print("ОШИБКА: в документе не распознаны строки занятий.", file=sys.stderr)
        return 3

    # 3. Расписание
    try:
        items = build_schedule(titles, cal, weekdays, start_date)
    except ScheduleError as e:
        print(f"ОШИБКА расписания: {e}", file=sys.stderr)
        return 3

    missing = [it for it in items if it.date is None]
    if missing:
        print(f"ВНИМАНИЕ: не хватило дат для {len(missing)} занятий "
              f"(последний номер {missing[0].lesson_index}..{missing[-1].lesson_index}). "
              "Добавьте дни недели или начните раньше.", file=sys.stderr)

    print("Первые 5 занятий:")
    for it in items[:5]:
        print(f"  {it.lesson_index:>3}. {it.text}  {it.title[:60]}")

    # 4. Замена дат в .doc: rewrite_doc_dates ждёт (cp_start, cp_end, new_text)
    replacements = []
    for row, it in zip(rows, items):
        if not it.text or not row.date_pos:
            continue
        start, end = row.date_pos
        new_text = it.text
        if len(new_text) != end - start:  # выравниваем длину ("4.09" -> "04.09")
            new_text = new_text.rjust(end - start, "0")
        replacements.append((start, end, new_text))
    out_dir = os.path.dirname(args.out)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    rewrite_doc_dates(data, replacements, out_path=args.out)
    print(f"Сохранено: {args.out}")

    # 5. Самопроверка: перечитаем результат
    with open(args.out, "rb") as f:
        check = parse_lessons_from_text(extract_doc_text(f.read()))
    check_rows = check.session_rows
    ok = sum(1 for src, got in zip(items, check_rows) if (got.date or "") == src.text)
    print(f"Самопроверка: {ok}/{len(items)} дат совпадают в записанном файле")
    return 0 if ok == len(items) and len(check_rows) == len(items) else 4


if __name__ == "__main__":
    raise SystemExit(main())
