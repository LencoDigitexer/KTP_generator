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
import threading
import uuid
import webbrowser

from flask import (Flask, jsonify, render_template, request,
                   send_from_directory)

from ktpgen.pipeline import (InputError, apply_to_document, calendar_info,
                             load_calendar_any, load_document,
                             parse_start_date, parse_weekdays,
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
        "start_iso": f"{start_year}-09-01",
        "cal_min": f"{start_year}-08-01",
        "cal_max": f"{start_year + 1}-09-30",
        "cal_note": "После загрузки календаря границы и дата подстроятся автоматически.",
        "weekdays": ["вт", "пт"],
        "out_name": "КТП.doc",
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
    if kind == "doc" and ext not in (".doc", ".docx"):
        return jsonify(error="Для плана нужен файл .doc (Word 97-2003)"), 400
    if kind == "calendar" and ext not in (".xlsx", ".xlsm", ".doc", ".docx"):
        return jsonify(error="Для календаря нужен файл .xlsx "
                             "(или Word .doc/.docx с сеткой месяцев)"), 400
    token = uuid.uuid4().hex[:8]
    name = f"{token}_{_safe_name(f.filename)}"
    path = os.path.join(UPLOAD_DIR, name)
    f.save(path)

    info = {"path": path, "name": f.filename, "kind": kind}
    try:
        if kind == "doc":
            rows = []
            if ext == ".docx":      # план пишется в .doc; .docx принимаем как календарь
                cal, _ = load_calendar_any(path)
                info.update(calendar_info(cal))
                info["is_calendar"] = True
            else:
                _, doc = load_document(path)
                rows = doc.session_rows
            if rows:                                   # это план КТП
                info["lessons"] = len(rows)
                info["preview"] = [r.title[:80] for r in rows[:5]]
            elif not info.get("is_calendar"):          # может быть, это календарь в Word?
                cal, _ = load_calendar_any(path)
                info.update(calendar_info(cal))
                info["is_calendar"] = True
        elif kind == "calendar":
            cal, ctype = load_calendar_any(path)
            info["calendar_type"] = ctype
            info.update(calendar_info(cal))
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
        cal, _ = load_calendar_any(data["calendar_path"])
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
    out_name = _safe_name(data.get("out_name") or "КТП.doc")
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


def _open_browser_later(url: str, delay: float = 1.2):
    """Открывает браузер после старта сервера; если не удалось — печатает ссылку."""
    def _try():
        import time
        time.sleep(delay)
        opened = False
        try:
            opened = webbrowser.open(url, new=2)   # новая вкладка
        except Exception:
            opened = False
        if not opened:
            for cmd in (("xdg-open", url), ("open", url)):
                try:
                    import subprocess
                    subprocess.Popen(cmd, stdout=subprocess.DEVNULL,
                                     stderr=subprocess.DEVNULL)
                    opened = True
                    break
                except OSError:
                    continue
        if not opened:
            print(f"Откройте ссылку вручную: {url}")
    threading.Thread(target=_try, daemon=True).start()


def main():
    p = argparse.ArgumentParser(description="GUI генератора КТП (Flask)")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=5000)
    p.add_argument("--debug", action="store_true")
    p.add_argument("--no-browser", action="store_true",
                   help="не открывать браузер автоматически")
    args = p.parse_args()
    show_host = "localhost" if args.host in ("127.0.0.1", "0.0.0.0") else args.host
    url = f"http://{show_host}:{args.port}"
    print(f"Генератор КТП запущен: {url}")
    if not args.no_browser:
        _open_browser_later(url)
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
