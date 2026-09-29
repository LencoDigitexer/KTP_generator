#!/usr/bin/env python3
"""Веб-GUI генератора КТП.

Запуск из корня проекта:

    python gui.py                  # http://127.0.0.1:5000
    python gui.py --port 8080
    python gui.py --host 0.0.0.0   # доступ из сети (например, с рабочего места)

Работает без интернета: все стили/скрипты встроены в шаблон.
Логика совпадает с CLI (общий модуль ktpgen/pipeline.py).
"""

from __future__ import annotations

import argparse
import datetime
import os
import re
import uuid

from flask import (Flask, jsonify, render_template, request,
                   send_from_directory)

from ktpgen.pipeline import (InputError, apply_to_document, load_calendar,
                             load_document, parse_start_date, parse_weekdays,
                             plan_lessons, run_pipeline, verify_output,
                             WEEKDAY_SHORT)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

SAFE_NAME_RE = re.compile(r"^[\w.\- ()\[\]\u0400-\u04FF]+$")

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "app", "templates"))
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 МБ


def _safe_name(filename: str) -> str:
    """Имя файла без путей и «страшных» символов, кириллица сохраняется."""
    name = os.path.basename(filename or "")
    name = name.replace("&", "and").strip()
    if not name or not SAFE_NAME_RE.match(name):
        name = "file"
    return name[:120]


@app.get("/")
def index():
    today = datetime.date.today()
    start_year = today.year if today.month >= 9 else today.year + 1
    defaults = {
        "start": f"01.09.{start_year}",
        "weekdays": ["вт", "пт"],
        "out_name": "КТП_готовый.doc",
    }
    return render_template("index.html", weekdays=list(WEEKDAY_SHORT),
                           defaults=defaults)


@app.post("/api/upload")
def upload():
    f = request.files.get("file")
    kind = request.form.get("kind", "")
    if not f or not f.filename:
        return jsonify(error="Файл не выбран"), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if kind == "doc" and ext not in (".doc",):
        return jsonify(error="Для плана нужен файл .doc (Word 97-2003)"), 400
    if kind == "calendar" and ext not in (".xlsx", ".xlsm"):
        return jsonify(error="Для календаря нужен файл .xlsx"), 400
    token = uuid.uuid4().hex[:8]
    name = f"{token}_{_safe_name(f.filename)}"
    path = os.path.join(UPLOAD_DIR, name)
    f.save(path)

    info = {"path": path, "name": f.filename, "kind": kind}
    try:
        if kind == "doc":
            _, doc = load_document(path)
            rows = doc.session_rows
            info["lessons"] = len(rows)
            info["preview"] = [r.title[:80] for r in rows[:5]]
            if not rows:
                os.remove(path)
                return jsonify(error="В документе не распознаны строки занятий "
                                     "(темы вида «Тема 1.1 …»)."), 400
        elif kind == "calendar":
            cal = load_calendar(path)
            info["working"] = len(cal.working_dates)
            info["holidays"] = sorted(d.strftime("%d.%m.%Y")
                                      for d in getattr(cal, "holidays", set()))
            first, last = cal.working_dates[0], cal.working_dates[-1]
            info["range"] = f"{first:%d.%m.%Y} — {last:%d.%m.%Y}"
    except InputError as e:
        os.remove(path)
        return jsonify(error=str(e)), 400
    except Exception as e:  # noqa: BLE001
        os.remove(path)
        return jsonify(error=f"Не удалось прочитать файл: {e}"), 400
    return jsonify(info)


@app.post("/api/preview")
def preview():
    """Без записи файла: показывает, какие даты получит каждое занятие."""
    data = request.get_json(silent=True) or {}
    try:
        weekdays = parse_weekdays(data.get("weekdays", ""))
        start = parse_start_date(data.get("start", ""))
        cal = load_calendar(data["calendar_path"])
        _, doc = load_document(data["doc_path"])
        items, rows, warnings = plan_lessons(doc, cal, weekdays, start)
    except (InputError, KeyError) as e:
        return jsonify(error=str(e) or "Загрузите оба файла"), 400
    except Exception as e:  # noqa: BLE001
        return jsonify(error=str(e)), 500
    lessons = [{
        "index": it.lesson_index,
        "title": it.title,
        "date": it.text or "—",
        "weekday": WEEKDAY_SHORT[it.date.weekday()] if it.date else "",
        "old": rows[i - 1].date or "" if 0 <= it.lesson_index - 1 < len(rows) else "",
    } for i, it in enumerate(items, 1)]
    return jsonify(lessons=lessons, warnings=warnings,
                   holidays=sorted(d.strftime("%d.%m.%Y")
                                   for d in getattr(cal, "holidays", set())),
                   working=len(cal.working_dates))


@app.post("/api/generate")
def generate():
    data = request.get_json(silent=True) or {}
    out_name = _safe_name(data.get("out_name") or "КТП_готовый.doc")
    if not out_name.lower().endswith(".doc"):
        out_name += ".doc"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    try:
        result = run_pipeline(
            doc_path=data["doc_path"],
            calendar_path=data["calendar_path"],
            weekdays_spec=data.get("weekdays", ""),
            start_spec=data.get("start", ""),
            out_path=out_path,
        )
    except (InputError, KeyError) as e:
        return jsonify(error=str(e) or "Загрузите оба файла"), 400
    except Exception as e:  # noqa: BLE001
        return jsonify(error=f"Ошибка при генерации: {e}"), 500
    ver = result["verification"]
    return jsonify(
        ok=ver["ok"] == ver["total"] and ver["rows_match"],
        saved=result["out_path"],
        download="/download/" + os.path.basename(result["out_path"]),
        replaced=len(result["changes"]),
        verification=ver,
        warnings=result["warnings"],
        lessons=result["lessons"],
    )


@app.get("/download/<path:name>")
def download(name):
    if not SAFE_NAME_RE.match(os.path.basename(name)):
        return "Недопустимое имя файла", 400
    return send_from_directory(OUTPUT_DIR, name, as_attachment=True)


def main():
    p = argparse.ArgumentParser(description="GUI генератора КТП (Flask)")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=5000)
    p.add_argument("--debug", action="store_true")
    args = p.parse_args()
    print(f"Откройте в браузере: http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
