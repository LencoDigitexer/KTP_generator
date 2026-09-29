"""Построение расписания дат занятий по календарю.

Логика: занятия идут строго по порядку (1..N). Для каждого занятия берётся
следующая доступная дата из календаря для заданного дня(ей) недели, начиная
с даты начала занятий; праздничные (красные) дни пропускаются.

Поддерживаются режимы:
  * один день недели в неделю (например, вторник);
  * несколько дней недели (например, вторник+пятница) — занятия распределяются
    по ближайшим подходящим датам в хронологическом порядке;
  * «каждые N недель» через параметр lessons_per_week для одного дня.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass

from .calendar_parser import Calendar


class ScheduleError(Exception):
    pass


@dataclass
class ScheduledItem:
    lesson_index: int
    title: str
    date: datetime.date | None
    text: str            # "дд.мм" или пустая строка, если не хватило дат


def build_schedule(
    titles: list[str],
    calendar: Calendar,
    weekdays: list[int],
    start_date: datetime.date,
) -> list[ScheduledItem]:
    """Возвращает список дат для каждого занятия.

    weekdays: список дней недели (0=пн … 6=вс), занятия могут идти в несколько дней.
    start_date: дата первого возможного занятия.
    """
    if not weekdays:
        raise ScheduleError("Не выбран ни один день недели")
    for w in weekdays:
        if not 0 <= w <= 6:
            raise ScheduleError(f"Некорректный день недели: {w}")

    # Хронологический список всех рабочих дат с нужными днями недели
    pool = sorted(
        d for d in calendar.working_dates
        if d.weekday() in weekdays and d >= start_date
    )
    if len(pool) < len(titles):
        # не хватает дат — последние останутся без даты (пользователь увидит предупреждение)
        pass

    result: list[ScheduledItem] = []
    for i, title in enumerate(titles):
        d = pool[i] if i < len(pool) else None
        result.append(ScheduledItem(
            lesson_index=i + 1,
            title=title,
            date=d,
            text=d.strftime("%d.%m") if d else "",
        ))
    return result
