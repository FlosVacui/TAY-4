#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Учёт проекта «ТАУ-3 NG»: tracker.yaml -> Excel-график (и обратно: заметки пользователя).

Источник правды - docs/tracker/tracker.yaml. Excel генерируется командой build.
Команды: build, sync, add, mark, log, review, review-status, decision.
Подробности: python tools/tracker.py --help, docs/tracker/README.md.
"""
import argparse
import datetime as dt
import json
import math
import os
import re
import sys
import tempfile
from pathlib import Path

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DEF_YAML = ROOT / "docs" / "tracker" / "tracker.yaml"
DEF_XLSX = ROOT / "docs" / "tracker" / "ТАУ-3_NG_график.xlsx"

# ----------------------------------------------------------------- константы
# статус -> (цвет заливки, цвет текста, пояснение)
STATUSES = {
    "План": ("7B5CB8", "FFFFFF", "Запланировано на эту неделю."),
    "В работе": ("2F6FB5", "FFFFFF", "Исполнитель работает над пунктом."),
    "Выполнено": ("2F6B2F", "FFFFFF", "Принято оркестратором (после ревью)."),
    "Доработка": ("F4C542", "333333", "После ревью есть блокирующие или существенные замечания."),
    "Ждёт решения": ("C0392B", "FFFFFF", "Нужен ответ пользователя (развилка), срок наступил."),
}
SEVERITIES = ["блокирующее", "существенное", "незначительное", "предложение"]
REVIEW_STATUSES = ["открыто", "исправлено", "отложено", "отклонено"]
PARITY_STATES = ["не начато", "сделано", "проверено"]

HEAD_FILL = "2E5E1E"
STAGE_FILL = "D9D9D9"
HOLIDAY_FILL = "FCE4D6"
SESSION_FILL = "FFF2CC"
BORDER_COLOR = "BFBFBF"
MONTHS = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Нояб", "Дек"]
DEFAULT_GRID_END = dt.date(2027, 4, 30)
SHEET_MAIN = "График"
SNAP_SHEET = "_snapshot"

DEFAULT_HEADER = """\
# ============================================================================
# ТАУ-3 NG - учёт проекта. ИСТОЧНИК ПРАВДЫ для Excel-графика.
# Excel (docs/tracker/ТАУ-3_NG_график.xlsx) собирается командой:
#     python tools/tracker.py build
# Правьте этот файл (руками или командами tracker.py add/mark/log/...), а не Excel.
# Исключение: колонка «Комментарий» листа «График» и колонки «Решение пользователя» /
# «Дата решения» листа «Решения» - их можно писать в Excel; `tracker.py sync` (и любая сборка) вернёт их сюда.
#
# Структура файла (ключи верхнего уровня):
#   project        название проекта (строка)
#   start_monday   понедельник 1-й недели сетки (дата; 2026-11-02)
#   grid_end       (необязательно) последняя дата сетки; по умолчанию 2027-04-30.
#                  Сетка сама расширяется, если план или факт выходят за неё.
#   calendar       календарь: holidays - праздники, session - сессия;
#                  списки пар [начало, конец] (включительно), недели закрашиваются в «Графике»
#   items          пункты графика (см. ниже)
#   history        записи листа «История»
#   reviews        замечания ревью (лист «Ревью»)
#   decisions      развилки Р-x (лист «Решения»)
#   parity         строки «Паритета» (заполняются на подэтапе 1.1)
#   roles          таблица ролей (лист «Легенда»)
#   external_requests  запросы во внешние нейросети ВН-xx (лист «Легенда»)
#
# items - список пунктов в порядке показа:
#   id         номер: "0" (этап), "1.2" (подэтап), "9.2.1" (глубже); всегда в кавычках
#   title      короткое название (до ~90 символов)
#   stage      true для строки этапа (жирная, серая; план и сроки берутся из подэтапов)
#   owner      исполнитель: роль (модель / effort)
#   review     ревью: кто проверяет
#   plan       {start: дата, end: дата} - плановые даты; недели выводятся из них
#   fact       {номер недели: статус} - факт по неделям. Статусы: План, В работе,
#              Выполнено, Доработка, Ждёт решения
#   fact_dates (необязательно) {start, end} - точные даты факта вместо границ недель
#   comment    заметка (синхронизируется с колонкой «Комментарий» Excel)
#   Правило клетки недели: есть факт - показываем его; иначе неделя внутри плана - «План»;
#   иначе пусто (с календарной подсветкой).
# Пустые поля в файле не пишутся. Комментарии внутри данных при перезаписи командами
# не сохраняются (этот заголовок сохраняется) - пишите заметки в поле comment.
#
# history:   {date, item, what, who, commit, rounds, tokens, link}
# reviews:   {id, item, round, severity, text, repro, by, status, comment}
#            severity: блокирующее | существенное | незначительное | предложение
#            status:   открыто | исправлено | отложено | отклонено
# decisions: {id, question, options, recommendation, deadline, default, answer, answer_date}
# parity:    {id, area, behavior, source, delphi, fr, item, state, test}
#            state: не начато | сделано | проверено
# ============================================================================
"""

KNOWN_KEYS = {"project", "start_monday", "grid_end", "calendar", "items", "history", "reviews",
              "decisions", "parity", "roles", "external_requests"}
ITEM_ORDER = ["id", "title", "stage", "owner", "review", "plan", "fact", "fact_dates", "comment"]
HISTORY_ORDER = ["date", "item", "what", "who", "commit", "rounds", "tokens", "link"]
REVIEW_ORDER = ["id", "item", "round", "severity", "text", "repro", "by", "status", "comment"]
DECISION_ORDER = ["id", "question", "options", "recommendation", "deadline", "default", "answer", "answer_date"]
PARITY_ORDER = ["id", "area", "behavior", "source", "delphi", "fr", "item", "state", "test"]
ROLE_ORDER = ["role", "model", "task"]
EXT_ORDER = ["id", "item", "what"]


class TrackerError(Exception):
    """Ошибка данных или аргументов (код выхода 1)."""


class LockedError(Exception):
    """Excel-файл занят (код выхода 2)."""


# ----------------------------------------------------------------- YAML
def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise TrackerError(f"Не удалось разобрать дату: {v!r} (нужно ГГГГ-ММ-ДД или ДД.ММ.ГГГГ)")


def extract_header(text):
    lines = []
    for ln in text.splitlines():
        if ln.startswith("#") or not ln.strip():
            lines.append(ln)
        else:
            break
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines) + "\n" if lines else ""


def normalize(data):
    if not isinstance(data, dict):
        raise TrackerError("tracker.yaml: корень должен быть словарём")
    data.setdefault("project", "ТАУ-3 NG")
    data["start_monday"] = to_date(data.get("start_monday"))
    if data["start_monday"] is None:
        raise TrackerError("tracker.yaml: не задан start_monday")
    if data["start_monday"].weekday() != 0:
        raise TrackerError("tracker.yaml: start_monday должен быть понедельником")
    if data.get("grid_end") is not None:
        data["grid_end"] = to_date(data["grid_end"])
    cal = data.setdefault("calendar", {}) or {}
    data["calendar"] = cal
    for key in ("holidays", "session"):
        cal[key] = [[to_date(a), to_date(b)] for a, b in (cal.get(key) or [])]
    for key in ("items", "history", "reviews", "decisions", "parity", "roles", "external_requests"):
        data[key] = data.get(key) or []
    seen = set()
    for it in data["items"]:
        if "id" not in it or not it.get("title"):
            raise TrackerError(f"У пункта {it.get('id', it)} нет обязательных полей id/title")
        it["id"] = str(it["id"])
        if it["id"] in seen:
            raise TrackerError(f"Повтор номера пункта: {it['id']}")
        seen.add(it["id"])
        if it.get("plan"):
            p = it["plan"]
            p["start"], p["end"] = to_date(p.get("start")), to_date(p.get("end"))
            if p["start"] is None or p["end"] is None or p["start"] > p["end"]:
                raise TrackerError(f"Пункт {it['id']}: некорректный plan {p}")
        fact = {}
        for k, v in (it.get("fact") or {}).items():
            if v not in STATUSES:
                raise TrackerError(f"Пункт {it['id']}: неизвестный статус {v!r} (допустимо: {', '.join(STATUSES)})")
            fact[int(k)] = v
        it["fact"] = fact
        if it.get("fact_dates"):
            fd = it["fact_dates"]
            fd["start"], fd["end"] = to_date(fd.get("start")), to_date(fd.get("end"))
    for h in data["history"]:
        h["date"] = to_date(h.get("date"))
        if "item" in h:
            h["item"] = str(h["item"])
    for r in data["reviews"]:
        r["item"] = str(r.get("item", ""))
    for d in data["decisions"]:
        d["id"] = str(d["id"])
        d["deadline"] = to_date(d.get("deadline"))
        d["answer_date"] = to_date(d.get("answer_date"))
    return data


def load_data(path):
    text = Path(path).read_text(encoding="utf-8")
    data = yaml.safe_load(text)
    return normalize(data), extract_header(text)


def _qkey(k):
    if isinstance(k, int):
        return str(k)
    return k if re.fullmatch(r"[\w.\-]+", k) else json.dumps(k, ensure_ascii=False)


def flow(v):
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (dt.datetime, dt.date)):
        return (v.date() if isinstance(v, dt.datetime) else v).isoformat()
    if isinstance(v, (int, float)):
        return repr(v)
    if isinstance(v, dict):
        return "{" + ", ".join(f"{_qkey(k)}: {flow(x)}" for k, x in v.items()) + "}"
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(flow(x) for x in v) + "]"
    return json.dumps(str(v), ensure_ascii=False)


def _empty(v):
    return v is None or v == "" or v == {} or v == []


def emit_list(name, rows, order, always=("id",)):
    if not rows:
        return [f"{name}: []"]
    out = [f"{name}:"]
    for row in rows:
        keys = [k for k in order if k in row] + [k for k in row if k not in order]
        keys = [k for k in keys if k in always or not _empty(row[k])]
        for i, k in enumerate(keys):
            out.append(f"  {'- ' if i == 0 else '  '}{k}: {flow(row[k])}")
    return out


def dump_yaml(data, header):
    out = []
    out.append((header or DEFAULT_HEADER).rstrip("\n"))
    out.append("")
    out.append(f"project: {flow(data['project'])}")
    out.append(f"start_monday: {flow(data['start_monday'])}")
    if data.get("grid_end"):
        out.append(f"grid_end: {flow(data['grid_end'])}")
    out.append("")
    out.append("calendar:")
    for key in ("holidays", "session"):
        out.append(f"  {key}:")
        for a, b in data["calendar"].get(key, []):
            out.append(f"    - [{flow(a)}, {flow(b)}]")
        if not data["calendar"].get(key):
            out[-1] = f"  {key}: []"
    out.append("")
    out.append("items:")
    for it in data["items"]:
        if it.get("stage"):
            out.append("")
        row = dict(it)
        keys = [k for k in ITEM_ORDER if k in row] + [k for k in row if k not in ITEM_ORDER]
        keys = [k for k in keys if k in ("id", "title") or not _empty(row[k])]
        for i, k in enumerate(keys):
            out.append(f"  {'- ' if i == 0 else '  '}{k}: {flow(row[k])}")
    out.append("")
    out += emit_list("history", data["history"], HISTORY_ORDER, always=("date",))
    out.append("")
    out += emit_list("reviews", data["reviews"], REVIEW_ORDER)
    out.append("")
    out += emit_list("decisions", data["decisions"], DECISION_ORDER)
    out.append("")
    out += emit_list("parity", data["parity"], PARITY_ORDER)
    out.append("")
    out += emit_list("roles", data["roles"], ROLE_ORDER, always=("role",))
    out.append("")
    out += emit_list("external_requests", data["external_requests"], EXT_ORDER)
    for k in data:
        if k not in KNOWN_KEYS:
            out.append("")
            out.append(f"{_qkey(k)}: {flow(data[k])}")
    return "\n".join(out) + "\n"


def atomic_write_text(path, text):
    path = Path(path)
    fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=".tmp_", suffix=".yaml")
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise


def save_data(path, data, header):
    atomic_write_text(path, dump_yaml(data, header))


# ----------------------------------------------------------------- модель: недели и сроки
def week_of(d, start):
    return (d - start).days // 7 + 1


def week_monday(n, start):
    return start + dt.timedelta(days=7 * (n - 1))


def week_sunday(n, start):
    return week_monday(n, start) + dt.timedelta(days=6)


def fmt_date(d):
    return d.strftime("%d.%m.%Y") if d else ""


def descendants(items, it):
    pref = it["id"] + "."
    return [x for x in items if x["id"].startswith(pref)]


def eff_plan(items, it):
    """(start, end) или None; для этапа - объединение планов подэтапов."""
    if it.get("stage"):
        ds = [x["plan"] for x in descendants(items, it) if x.get("plan")]
        if ds:
            return min(p["start"] for p in ds), max(p["end"] for p in ds)
    p = it.get("plan")
    return (p["start"], p["end"]) if p else None


def plan_weeks(items, it, start):
    p = eff_plan(items, it)
    if not p:
        return set()
    return set(range(week_of(p[0], start), week_of(p[1], start) + 1))


def own_fact_range(it, start):
    """(fact_start, fact_end) по собственному факту пункта."""
    fd = it.get("fact_dates") or {}
    fs, fe = fd.get("start"), fd.get("end")
    fact = it.get("fact") or {}
    if fact:
        last = max(fact)
        if fs is None:
            fs = week_monday(min(fact), start)
        if fe is None and fact[last] == "Выполнено":
            fe = week_sunday(last, start)
    return fs, fe


def eff_fact(items, it, start):
    if it.get("stage"):
        ds = [x for x in descendants(items, it) if not x.get("stage")]
        ranges = [own_fact_range(x, start) for x in ds]
        starts = [r[0] for r in ranges if r[0]]
        fs = min(starts) if starts else None
        fe = None
        if ds and all(r[1] for r in ranges):
            fe = max(r[1] for r in ranges)
        o = own_fact_range(it, start)
        return o[0] or fs, o[1] or fe
    return own_fact_range(it, start)


def eff_status(items, it):
    fact = it.get("fact") or {}
    if fact:
        return fact[max(fact)]
    if it.get("stage"):
        ds = [x for x in descendants(items, it) if not x.get("stage")]
        if ds and all(eff_status(items, x) == "Выполнено" for x in ds):
            return "Выполнено"
        if any(x.get("fact") for x in ds):
            return "В работе"
    return "План"


PRIORITY = ["Ждёт решения", "Доработка", "В работе"]


def week_status(items, it, w, start):
    """Статус клетки недели w у пункта (None - пусто)."""
    fact = it.get("fact") or {}
    if w in fact:
        return fact[w]
    if not it.get("stage"):
        return "План" if w in plan_weeks(items, it, start) else None
    sts = []
    for x in descendants(items, it):
        if x.get("stage"):
            continue
        st = week_status(items, x, w, start)
        if st:
            sts.append(st)
    for name in PRIORITY:
        if name in sts:
            return name
    if sts and all(st == "Выполнено" for st in sts):
        return "Выполнено"
    if sts:
        return "План"
    return "План" if w in plan_weeks(items, it, start) else None


def grid_weeks(data):
    start = data["start_monday"]
    end = data.get("grid_end") or DEFAULT_GRID_END
    n = week_of(end, start)
    for it in data["items"]:
        p = eff_plan(data["items"], it)
        if p:
            n = max(n, week_of(p[1], start))
        if it.get("fact"):
            n = max(n, max(it["fact"]))
        fd = it.get("fact_dates") or {}
        if fd.get("end"):
            n = max(n, week_of(fd["end"], start))
    return max(n, 1)


def week_label(n, start):
    a, b = week_monday(n, start), week_sunday(n, start)
    if a.month == b.month:
        return MONTHS[a.month - 1], f"{a.day}-{b.day}"
    return f"{MONTHS[a.month - 1]}-{MONTHS[b.month - 1]}", f"{a.day}-{b.day}"


def week_tint(n, data):
    a, b = week_monday(n, data["start_monday"]), week_sunday(n, data["start_monday"])
    cal = data["calendar"]
    for s, e in cal.get("holidays", []):
        if s <= b and e >= a:
            return HOLIDAY_FILL
    for s, e in cal.get("session", []):
        if s <= b and e >= a:
            return SESSION_FILL
    return None


# ----------------------------------------------------------------- Excel: стили
thin = Side(style="thin", color=BORDER_COLOR)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_BORDER = Border(*(Side(style="thin", color="FFFFFF"),) * 4)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def style_head(c, size=10):
    c.fill = fill(HEAD_FILL)
    c.font = Font(name="Calibri", size=size, bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = HEAD_BORDER


def style_range(ws, rng, fn):
    for row in ws[rng]:
        for c in row:
            fn(c)


def setup_print(ws, title_rows, a3=False):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if a3 else ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = title_rows
    ws.page_margins.left = ws.page_margins.right = 0.4


def build_main(wb, data):
    ws = wb.active
    ws.title = SHEET_MAIN
    items, start = data["items"], data["start_monday"]
    n = grid_weeks(data)
    c0 = 3                      # первая колонка недель
    cl = c0 + n - 1             # последняя колонка недель
    cc = cl + 1                 # «Комментарий»

    ws.merge_cells(start_row=1, start_column=1, end_row=5, end_column=1)
    ws.merge_cells(start_row=1, start_column=2, end_row=5, end_column=2)
    ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=cl)
    ws.merge_cells(start_row=1, start_column=cc, end_row=5, end_column=cc)
    ws.cell(1, 1, "№ п/п")
    ws.cell(1, 2, "Этап")
    ws.cell(1, c0, "Сроки выполнения/месяцы")
    ws.cell(1, cc, "Комментарий")
    # годы
    for w in range(1, n + 1):
        col = c0 + w - 1
        mon, rng = week_label(w, start)
        ws.cell(3, col, mon)
        ws.cell(4, col, rng)
        ws.cell(5, col, f"{w} нед.")
    col = c0
    while col <= cl:
        y = week_monday(col - c0 + 1, start).year
        e = col
        while e + 1 <= cl and week_monday(e + 1 - c0 + 1, start).year == y:
            e += 1
        ws.cell(2, col, y)
        if e > col:
            ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=e)
        col = e + 1
    for r in range(1, 6):
        for c in range(1, cc + 1):
            style_head(ws.cell(r, c), size=10 if r < 5 else 9)
    ws.cell(1, 2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 18
    for r in (2, 3, 4, 5):
        ws.row_dimensions[r].height = 15

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 60
    for c in range(c0, cl + 1):
        ws.column_dimensions[get_column_letter(c)].width = 9.5
    ws.column_dimensions[get_column_letter(cc)].width = 45

    r = 6
    for it in items:
        stage = bool(it.get("stage"))
        a = ws.cell(r, 1, it["id"])
        a.number_format = "@"
        b = ws.cell(r, 2, it["title"])
        cm = ws.cell(r, cc, it.get("comment") or None)
        cm.number_format = "@"
        base_font = Font(name="Calibri", size=10, bold=stage)
        row_fill = fill(STAGE_FILL) if stage else None
        for c in range(1, cc + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            cell.font = base_font
            if row_fill:
                cell.fill = row_fill
        a.alignment = Alignment(horizontal="center", vertical="center")
        b.alignment = Alignment(vertical="center", wrap_text=True)
        cm.alignment = Alignment(vertical="center", wrap_text=True)
        cm.font = Font(name="Calibri", size=9, bold=False)
        tall = False
        for w in range(1, n + 1):
            cell = ws.cell(r, c0 + w - 1)
            st = week_status(items, it, w, start)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if st:
                f_, fc, _ = STATUSES[st]
                cell.value = st
                cell.fill = fill(f_)
                cell.font = Font(name="Calibri", size=8, color=fc)
                if st == "Ждёт решения":
                    tall = True
            elif not stage:
                tint = week_tint(w, data)
                if tint:
                    cell.fill = fill(tint)
        lines = max(1, math.ceil(len(it["title"]) / (52 if stage else 58)))
        h = max(15.0, 13.5 * lines + 2)
        if tall:
            h = max(h, 24.0)
        ws.row_dimensions[r].height = h
        r += 1
    ws.freeze_panes = "C6"
    ws.sheet_view.zoomScale = 90
    setup_print(ws, "1:5", a3=True)


def simple_sheet(wb, title, headers, widths, rows, date_cols=(), bold_rows=(), note=None):
    ws = wb.create_sheet(title)
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, i, h)
        style_head(c)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 32
    r = 2
    for idx, row in enumerate(rows):
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i, v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.font = Font(name="Calibri", size=10, bold=idx in bold_rows)
            if i in date_cols and v is not None and not isinstance(v, str):
                c.number_format = "DD.MM.YYYY"
                c.alignment = Alignment(horizontal="center", vertical="top")
            if idx in bold_rows:
                c.fill = fill(STAGE_FILL)
        r += 1
    if note:
        ws.cell(r, 1, note).font = Font(name="Calibri", size=10, italic=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers))
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
    ws.freeze_panes = "A2"
    setup_print(ws, "1:1")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(r - (1 if note else 0), 2)}"
    return ws


def owner_text(it):
    return it.get("owner") or ""


def build_schedule(wb, data):
    items, start = data["items"], data["start_monday"]
    rows, bold = [], set()
    for i, it in enumerate(items):
        p = eff_plan(items, it)
        fs, fe = eff_fact(items, it, start)
        dev = (fe - p[1]).days if (fe and p) else None
        if it.get("stage"):
            bold.add(i)
        rows.append([it["id"], it["title"], owner_text(it), it.get("review") or "",
                     p[0] if p else None, p[1] if p else None, fs, fe, dev, eff_status(items, it)])
    ws = simple_sheet(
        wb, "Сроки",
        ["№", "Этап", "Исполнитель (роль, модель/effort)", "Ревью", "План начало", "План конец",
         "Факт начало", "Факт конец", "Отклонение, дн.", "Статус"],
        [8, 60, 28, 34, 13, 13, 13, 13, 13, 15], rows, date_cols=(5, 6, 7, 8), bold_rows=bold)
    for r in range(2, len(rows) + 2):
        ws.cell(r, 1).number_format = "@"
        ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(r, 9).alignment = Alignment(horizontal="center", vertical="top")
        c = ws.cell(r, 10)
        st = c.value
        if st in STATUSES:
            c.fill = fill(STATUSES[st][0])
            c.font = Font(name="Calibri", size=10, color=STATUSES[st][1])
        c.alignment = Alignment(horizontal="center", vertical="top")


def build_history(wb, data):
    rows = [[h.get("date"), h.get("item", ""), h.get("what", ""), h.get("who", ""), h.get("commit", ""),
             h.get("rounds"), h.get("tokens"), h.get("link", "")] for h in data["history"]]
    simple_sheet(wb, "История",
                 ["Дата", "Пункт", "Что сделано", "Кто (роль, модель, effort или «пользователь»)",
                  "Коммит", "Раунды ревью", "Токены", "Ссылка"],
                 [12, 9, 60, 32, 24, 10, 12, 30], rows, date_cols=(1,))


def build_reviews(wb, data):
    rows = [[r.get("id", ""), r.get("item", ""), r.get("round"), r.get("severity", ""), r.get("text", ""),
             r.get("repro", ""), r.get("by", ""), r.get("status", ""), r.get("comment", "")] for r in data["reviews"]]
    simple_sheet(wb, "Ревью",
                 ["ID", "Пункт", "Раунд", "Серьёзность", "Суть", "Воспроизведение", "Кто нашёл", "Статус", "Комментарий"],
                 [14, 9, 8, 16, 55, 45, 24, 13, 35], rows)


def build_decisions(wb, data):
    rows = [[d["id"], d.get("question", ""), d.get("options", ""), d.get("recommendation", ""), d.get("deadline"),
             d.get("default", ""), d.get("answer", ""), d.get("answer_date")] for d in data["decisions"]]
    ws = simple_sheet(wb, "Решения",
                      ["ID", "Вопрос", "Варианты", "Рекомендация", "Срок", "Если ответа нет",
                       "Решение пользователя", "Дата решения"],
                      [10, 40, 45, 40, 12, 35, 40, 13], rows, date_cols=(5, 8))
    for r in range(2, len(rows) + 2):
        for c in (7, 8):
            ws.cell(r, c).fill = fill("FFFBE6")
            ws.cell(r, c).number_format = "@" if c == 7 else "DD.MM.YYYY"


def build_parity(wb, data):
    rows = [[p.get("id", ""), p.get("area", ""), p.get("behavior", ""), p.get("source", ""), p.get("delphi", ""),
             p.get("fr", ""), p.get("item", ""), p.get("state", ""), p.get("test", "")] for p in data["parity"]]
    note = None
    if not rows:
        note = "Таблица паритета заполняется на подэтапе 1.1 (требования и паритет по поведению)."
    simple_sheet(wb, "Паритет",
                 ["ID", "Область", "Функция / поведение", "Источник (1.1 / К2)", "Где в Delphi",
                  "Требование FR", "Подэтап", "Состояние (не начато / сделано / проверено)", "Тест"],
                 [10, 18, 50, 14, 28, 14, 10, 18, 30], rows, note=note)


def build_legend(wb, data):
    ws = wb.create_sheet("Легенда")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 70
    r = 1

    def title(text):
        nonlocal r
        c = ws.cell(r, 1, text)
        c.font = Font(name="Calibri", size=12, bold=True, color=HEAD_FILL)
        r += 1

    def head(cols):
        nonlocal r
        for i, h in enumerate(cols, 1):
            style_head(ws.cell(r, i, h))
        r += 1

    def row(vals, bold=False):
        nonlocal r
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.font = Font(name="Calibri", size=10, bold=bold)
        r += 1

    title("Статусы в клетках недель")
    head(["Статус", "Образец", "Значение"])
    for name, (f_, fc, desc) in STATUSES.items():
        row([name, name, desc])
        s = ws.cell(r - 1, 2)
        s.fill = fill(f_)
        s.font = Font(name="Calibri", size=8, color=fc)
        s.alignment = Alignment(horizontal="center", vertical="center")
    row(["Недели праздников", "", "Светло-оранжевая подсветка пустых клеток."])
    ws.cell(r - 1, 2).fill = fill(HOLIDAY_FILL)
    row(["Недели сессии", "", "Светло-жёлтая подсветка пустых клеток."])
    ws.cell(r - 1, 2).fill = fill(SESSION_FILL)
    row(["Правило клетки", "", "Есть факт - показан факт; иначе неделя в плане - «План»; иначе пусто."])
    r += 1
    title("Правила учёта")
    rules = [
        "Источник правды - docs/tracker/tracker.yaml. Excel собирается командой «python tools/tracker.py build»; правки в Excel, кроме указанных ниже, при пересборке теряются.",
        "Excel и tracker.yaml пишет только оркестратор. Субагенты их не правят и сборку не запускают.",
        "Заметки пользователя - в колонку «Комментарий» листа «График» и в колонки «Решение пользователя» / «Дата решения» листа «Решения». Команда «tracker.py sync» (и любая сборка) вернёт их в tracker.yaml.",
        "Чтобы стереть заметку, очистите ячейку. Строки, добавленные в Excel вручную, при пересборке исчезают: новые пункты добавляйте командой add.",
        "Если файл открыт в Excel (есть «~$…») или занят OneDrive, Excel не перезаписывается, команда сообщает об этом (код выхода 2); изменения команд add/mark/log/… при этом уже сохранены в tracker.yaml. Закройте файл и запустите build.",
        "Перед каждой записью Excel заметки пользователя автоматически считываются в YAML (sync). Если правка есть и в YAML, и в Excel, сохраняются обе версии («… | Excel: …»).",
        "Даты ориентировочные и пересчитываются на воротах каждого этапа.",
    ]
    for t in rules:
        c = ws.cell(r, 1, t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    title("Роли субагентов (модель / effort)")
    head(["Роль", "Модель / effort", "Что делает"])
    for ro in data["roles"]:
        row([ro.get("role", ""), ro.get("model", ""), ro.get("task", "")])
    r += 1
    title("Запросы во внешние нейросети (ВН-xx)")
    head(["Запрос", "Подэтап", "Что"])
    for e in data["external_requests"]:
        row([e.get("id", ""), e.get("item", ""), e.get("what", "")])
    ws.freeze_panes = "A2"
    setup_print(ws, "1:1")


def build_snapshot(wb, data):
    """Скрытый лист: что записано в редактируемые пользователем ячейки (для трёхсторонней склейки в sync)."""
    ws = wb.create_sheet(SNAP_SHEET)
    for it in data["items"]:
        ws.append(["comment", it["id"], it.get("comment") or ""])
    for d in data["decisions"]:
        ws.append(["answer", d["id"], d.get("answer") or ""])
        ws.append(["answer_date", d["id"], d["answer_date"].isoformat() if d.get("answer_date") else ""])
    ws.sheet_state = "veryHidden"


def build_workbook(data):
    wb = Workbook()
    wb.properties.creator = "tools/tracker.py"
    wb.properties.title = data.get("project", "")
    build_main(wb, data)
    build_schedule(wb, data)
    build_history(wb, data)
    build_reviews(wb, data)
    build_decisions(wb, data)
    build_parity(wb, data)
    build_legend(wb, data)
    build_snapshot(wb, data)
    wb.active = 0
    return wb


# ----------------------------------------------------------------- сборка файла
def check_lock(target):
    target = Path(target)
    locks = sorted(target.parent.glob("~$*")) if target.parent.exists() else []
    if locks:
        raise LockedError(
            f"Файл {target.name} открыт в Excel (найден {locks[0].name}). "
            f"Закройте его и повторите команду. Ничего не записано.")
    if target.exists():
        try:
            with open(target, "ab"):
                pass
        except PermissionError:
            raise LockedError(
                f"Файл {target.name} занят (Excel или синхронизация OneDrive). "
                f"Закройте его/дождитесь синхронизации и повторите. Ничего не записано.")


def build_xlsx(data, target):
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    check_lock(target)
    wb = build_workbook(data)
    fd, tmp = tempfile.mkstemp(dir=target.parent, prefix=".tmp_", suffix=".xlsx")
    os.close(fd)
    try:
        wb.save(tmp)
        try:
            os.replace(tmp, target)
        except PermissionError:
            raise LockedError(f"Не удалось заменить {target.name}: файл занят. Ничего не записано.")
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return target


# ----------------------------------------------------------------- sync
def _cell_text(v):
    """Текст ячейки (числа и даты -> строка); пусто -> ''."""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return to_date(v).isoformat()
    return str(v).strip()


def _find_col(ws, names, rows=(1, 2, 3, 4, 5)):
    for r in rows:
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value in names:
                return c
    return None


def merge3(label, y, e, s, warnings):
    """Трёхсторонняя склейка: y - YAML, e - Excel, s - снимок при сборке (None - снимка нет)."""
    y, e = (y or "").strip(), (e or "").strip()
    if e == "-":
        e = ""
    if s is None:                       # старый файл без снимка: пустое = без изменений
        return y if e == "" else e
    s = s.strip()
    if e == s:                          # пользователь не трогал - верен YAML
        return y
    if y == s or y == e:                # YAML не менялся (или совпал) - берём Excel
        return e
    if not y:
        return e
    if not e:
        return y
    warnings.append(f"конфликт {label}: правка и в YAML, и в Excel; сохранены обе версии")
    return f"{y} | Excel: {e}"


def read_snapshot(wb):
    if SNAP_SHEET not in wb.sheetnames:
        return None
    snap = {}
    for row in wb[SNAP_SHEET].iter_rows(values_only=True):
        row = list(row) + [None] * 3
        if row[0]:
            snap[(row[0], str(row[1]))] = "" if row[2] is None else str(row[2])
    return snap


def sync_from_xlsx(data, xlsx):
    """Читает заметки пользователя из Excel в data (трёхсторонняя склейка со снимком)."""
    if not Path(xlsx).exists():
        raise TrackerError(f"Нет файла Excel для sync: {xlsx}. Запустите build, чтобы создать его.")
    locks = list(Path(xlsx).parent.glob("~$*"))
    if locks:
        print(f"Предупреждение: файл открыт в Excel ({locks[0].name}); будут прочитаны только сохранённые данные.")
    try:
        wb = load_workbook(xlsx, data_only=True)
    except Exception as ex:
        raise TrackerError(f"Не удалось прочитать Excel {Path(xlsx).name} ({type(ex).__name__}: {ex}). "
                           f"Ничего не записано; исправьте или удалите файл и запустите build.")
    snap = read_snapshot(wb)
    changes, warnings = [], []

    def S(kind, key):
        return None if snap is None else snap.get((kind, key), "")

    by_id = {it["id"]: it for it in data["items"]}
    if SHEET_MAIN in wb.sheetnames:
        ws = wb[SHEET_MAIN]
        cc = _find_col(ws, ("Комментарий",))
        extra = []
        if cc:
            for r in range(6, ws.max_row + 1):
                iid = _cell_text(ws.cell(r, 1).value)
                if not iid:
                    continue
                if iid not in by_id:
                    extra.append(iid)
                    continue
                it = by_id[iid]
                old = it.get("comment") or ""
                new = merge3(f"комментария {iid}", old, _cell_text(ws.cell(r, cc).value), S("comment", iid), warnings)
                if new != old.strip():
                    it["comment"] = new
                    changes.append(f"комментарий к {iid}")
        if extra:
            warnings.append("строки " + ", ".join(extra) + " есть в Excel, но не в YAML: пропущены "
                            "(при пересборке исчезнут; добавляйте пункты командой add)")
    if "Решения" in wb.sheetnames:
        ws = wb["Решения"]
        cols = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
        ca, cd = cols.get("Решение пользователя"), cols.get("Дата решения")
        dec = {d["id"]: d for d in data["decisions"]}
        if ca:
            for r in range(2, ws.max_row + 1):
                did = _cell_text(ws.cell(r, 1).value)
                if not did or did not in dec:
                    continue
                d = dec[did]
                y_ans = d.get("answer") or ""
                e_ans = _cell_text(ws.cell(r, ca).value)
                new_ans = merge3(f"решения {did}", y_ans, e_ans, S("answer", did), warnings)
                ans_changed = new_ans != y_ans.strip()
                if ans_changed:
                    d["answer"] = new_ans
                    changes.append(f"решение {did}")
                # дата
                y_date = d["answer_date"].isoformat() if d.get("answer_date") else ""
                s_date = S("answer_date", did)
                e_date = s_date or ""
                if cd:
                    raw = ws.cell(r, cd).value
                    try:
                        e_date = _cell_text(to_date(raw)) if raw not in (None, "", "-") else ""
                    except TrackerError:
                        warnings.append(f"дата решения {did} не распознана ({raw!r}): пропущена")
                new_date = merge3(f"даты решения {did}", y_date, e_date, s_date, warnings)
                if "| Excel:" in new_date:
                    new_date = y_date
                date_touched = (e_date != "") if s_date is None else (e_date != s_date)
                if ans_changed and not date_touched:
                    new_date = dt.date.today().isoformat() if new_ans else ""
                if new_ans and not new_date:
                    new_date = dt.date.today().isoformat()
                if new_date != y_date:
                    d["answer_date"] = to_date(new_date) if new_date else None
                    changes.append(f"дата решения {did}")
    for w in warnings:
        print(f"Предупреждение: {w}")
    return changes


def autosync(data, xpath):
    """Перед записью Excel подтягивает заметки пользователя (sync). Возвращает список изменений."""
    xpath = Path(xpath)
    if not xpath.exists():
        return []
    if list(xpath.parent.glob("~$*")):
        return []                       # файл занят: сборка всё равно откажет с кодом 2
    changes = sync_from_xlsx(data, xpath)
    if changes:
        print("Из Excel подтянуто: " + "; ".join(changes))
    return changes


# ----------------------------------------------------------------- разбор аргументов
def norm_status(s):
    key = s.strip().lower().replace("ё", "е")
    for name in STATUSES:
        if name.lower().replace("ё", "е") == key:
            return name
    if key in ("-", "нет", "очистить", "clear"):
        return None
    raise TrackerError(f"Неизвестный статус {s!r}. Допустимо: {', '.join(STATUSES)} (или «-» - убрать факт).")


MAX_WEEKS = 520


def parse_weeks(arg, start):
    ws = _parse_weeks(arg, start)
    if max(ws) > MAX_WEEKS:
        raise TrackerError(f"Номер недели больше {MAX_WEEKS} (предел разумной сетки; у Excel не более 16384 колонок).")
    return ws


def _parse_weeks(arg, start):
    s = arg.strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}|\d{1,2}\.\d{1,2}\.\d{2,4}", s):
        n = week_of(to_date(s), start)
        if n < 1:
            raise TrackerError(f"Дата {s} раньше первой недели ({fmt_date(start)}).")
        return [n]
    m = re.fullmatch(r"(\d+)-(\d+)", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if a < 1 or b < a:
            raise TrackerError(f"Некорректный диапазон недель {s}")
        return list(range(a, b + 1))
    if s.isdigit() and int(s) >= 1:
        return [int(s)]
    raise TrackerError(f"Не понял неделю/дату {arg!r}: нужен номер недели (3), диапазон (3-5) или дата.")


def get_item(data, iid):
    for it in data["items"]:
        if it["id"] == iid:
            return it
    raise TrackerError(f"Нет пункта с номером {iid}")


def paths(args):
    y = Path(args.yaml or os.environ.get("TRACKER_YAML") or DEF_YAML)
    x = Path(args.out or os.environ.get("TRACKER_XLSX") or DEF_XLSX)
    return y, x


def finish(args, data, header, ypath, xpath, msg):
    if not getattr(args, "no_build", False):
        autosync(data, xpath)           # заметки пользователя из Excel не теряются; при ошибке ничего не пишется
    save_data(ypath, data, header)
    print(msg)
    if getattr(args, "no_build", False):
        return 0
    return do_build(data, xpath, after_save=True)


def do_build(data, xpath, after_save=False):
    try:
        out = build_xlsx(data, xpath)
    except LockedError as e:
        print(f"ОШИБКА: {e}")
        if after_save:
            print("tracker.yaml сохранён; Excel не обновлён - запустите «build» позже.")
        return 2
    print(f"Excel собран: {out}")
    return 0


# ----------------------------------------------------------------- команды
def cmd_build(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    if autosync(data, xpath):
        save_data(ypath, data, header)
    return do_build(data, xpath)


def cmd_sync(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    changes = sync_from_xlsx(data, xpath)
    if changes:
        save_data(ypath, data, header)
        print("Прочитано из Excel: " + "; ".join(changes))
    else:
        print("В Excel нет новых заметок - tracker.yaml без изменений.")
    if args.no_build:
        return 0
    return do_build(data, xpath, after_save=True)


def cmd_add(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    items = data["items"]
    if any(it["id"] == args.id for it in items):
        raise TrackerError(f"Пункт {args.id} уже есть")
    new = {"id": args.id, "title": args.title}
    if args.owner:
        new["owner"] = args.owner
    if args.review:
        new["review"] = args.review
    if bool(args.start) != bool(args.end):
        raise TrackerError("Укажите вместе --start и --end")
    if args.start:
        s, e = to_date(args.start), to_date(args.end)
        if s > e:
            raise TrackerError("--start позже --end")
        new["plan"] = {"start": s, "end": e}
    new["fact"] = {}
    parent_id = args.parent
    if not parent_id and "." in args.id:
        parent_id = args.id.rsplit(".", 1)[0]
        print(f"Родитель определён по номеру: {parent_id}")
    if parent_id:
        parent = get_item(data, parent_id)
        if not args.id.startswith(parent_id + "."):
            raise TrackerError(f"Номер {args.id} должен начинаться с «{parent_id}.»")
        last = max(i for i, it in enumerate(items) if it["id"] == parent["id"] or it["id"].startswith(parent["id"] + "."))
        items.insert(last + 1, new)
    else:
        new["stage"] = True
        items.append(new)
    return finish(args, data, header, ypath, xpath, f"Добавлен пункт {args.id}: {args.title}")


def cmd_mark(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    it = get_item(data, args.id)
    weeks = parse_weeks(args.week, data["start_monday"])
    status = norm_status(args.status)
    for w in weeks:
        if status is None:
            it["fact"].pop(w, None)
        else:
            it["fact"][w] = status
    it["fact"] = dict(sorted(it["fact"].items()))
    what = status or "факт убран"
    return finish(args, data, header, ypath, xpath,
                  f"Пункт {args.id}: недели {','.join(map(str, weeks))} -> {what}")


def cmd_log(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    get_item(data, args.item)
    rec = {"date": to_date(args.date) if args.date else dt.date.today(), "item": args.item, "what": args.what,
           "who": args.who}
    if args.commit:
        rec["commit"] = args.commit
    if args.rounds is not None:
        rec["rounds"] = args.rounds
    if args.tokens is not None:
        rec["tokens"] = args.tokens
    if args.link:
        rec["link"] = args.link
    data["history"].append(rec)
    return finish(args, data, header, ypath, xpath, f"В историю добавлено ({args.item}): {args.what}")


def cmd_review(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    get_item(data, args.item)
    pref = f"R-{args.item}-"
    nums = [int(r["id"][len(pref):]) for r in data["reviews"]
            if str(r.get("id", "")).startswith(pref) and r["id"][len(pref):].isdigit()]
    n = max(nums, default=0) + 1
    rid = f"R-{args.item}-{n}"
    rec = {"id": rid, "item": args.item, "round": args.round, "severity": args.severity, "text": args.text,
           "status": "открыто"}
    if args.repro:
        rec["repro"] = args.repro
    if args.by:
        rec["by"] = args.by
    data["reviews"].append(rec)
    return finish(args, data, header, ypath, xpath, f"Замечание {rid} ({args.severity}) добавлено")


def cmd_review_status(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    st = args.status.strip().lower()
    if st not in REVIEW_STATUSES:
        raise TrackerError(f"Статус замечания: {', '.join(REVIEW_STATUSES)}")
    for r in data["reviews"]:
        if r.get("id") == args.id:
            r["status"] = st
            if args.comment:
                r["comment"] = args.comment
            return finish(args, data, header, ypath, xpath, f"Замечание {args.id}: {st}")
    raise TrackerError(f"Нет замечания {args.id}")


def cmd_decision(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    for d in data["decisions"]:
        if d["id"].lower() == args.id.lower():
            d["answer"] = args.answer
            d["answer_date"] = to_date(args.date) if args.date else dt.date.today()
            return finish(args, data, header, ypath, xpath, f"Решение {d['id']} записано: {args.answer}")
    raise TrackerError(f"Нет развилки {args.id}. Есть: {', '.join(d['id'] for d in data['decisions'])}")


# ----------------------------------------------------------------- CLI
def cmd_parity_add(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    if any(str(r.get("id")) == args.id for r in data["parity"]):
        raise TrackerError(f"Строка паритета {args.id} уже есть")
    rec = {"id": args.id, "area": args.area, "behavior": args.what, "source": args.source}
    if args.delphi:
        rec["delphi"] = args.delphi
    if args.fr:
        rec["fr"] = args.fr
    if args.substage:
        get_item(data, args.substage)
        rec["item"] = args.substage
    rec["state"] = "не начато"
    if args.test:
        rec["test"] = args.test
    data["parity"].append(rec)
    return finish(args, data, header, ypath, xpath, f"Паритет: добавлена строка {args.id}")


def cmd_parity_set(args):
    ypath, xpath = paths(args)
    data, header = load_data(ypath)
    st = args.state.strip().lower()
    if st not in PARITY_STATES:
        raise TrackerError(f"Состояние паритета: {', '.join(PARITY_STATES)}")
    for r in data["parity"]:
        if str(r.get("id")) == args.id:
            r["state"] = st
            if args.test:
                r["test"] = args.test
            return finish(args, data, header, ypath, xpath, f"Паритет {args.id}: {st}")
    raise TrackerError(f"Нет строки паритета {args.id}")


def make_parser():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--yaml", help="путь к tracker.yaml (по умолчанию docs/tracker/tracker.yaml; env TRACKER_YAML)")
    common.add_argument("--out", help="путь к xlsx (по умолчанию docs/tracker/ТАУ-3_NG_график.xlsx; env TRACKER_XLSX)")
    chg = argparse.ArgumentParser(add_help=False, parents=[common])
    chg.add_argument("--no-build", action="store_true", help="только записать YAML, Excel не пересобирать")

    p = argparse.ArgumentParser(
        prog="tracker.py",
        description="Учёт проекта «ТАУ-3 NG»: tracker.yaml -> Excel-график. Источник правды - YAML.")
    sub = p.add_subparsers(dest="cmd", required=True, metavar="команда")

    s = sub.add_parser("build", parents=[common], help="собрать Excel из tracker.yaml (атомарно)")
    s.set_defaults(fn=cmd_build)

    s = sub.add_parser("sync", parents=[chg], help="считать из Excel «Комментарий» и «Решение пользователя» в YAML, затем собрать")
    s.set_defaults(fn=cmd_sync)

    s = sub.add_parser("add", parents=[chg], help="добавить пункт после последнего потомка родителя")
    s.add_argument("--id", required=True, help="номер, напр. 1.8")
    s.add_argument("--title", required=True, help="название")
    s.add_argument("--parent", help="номер родителя (по умолчанию - из номера: 3.7 -> 3; номер без точки создаёт новый этап)")
    s.add_argument("--owner", help="исполнитель")
    s.add_argument("--review", help="ревьюер")
    s.add_argument("--start", help="план: начало (ГГГГ-ММ-ДД или ДД.ММ.ГГГГ)")
    s.add_argument("--end", help="план: конец")
    s.set_defaults(fn=cmd_add)

    s = sub.add_parser("mark", parents=[chg], help="поставить факт-статус пункту на неделю")
    s.add_argument("id", help="номер пункта")
    s.add_argument("week", help="номер недели (3), диапазон (3-5) или любая дата внутри недели")
    s.add_argument("status", help="; ".join(STATUSES) + " или «-» (убрать)")
    s.set_defaults(fn=cmd_mark)

    s = sub.add_parser("log", parents=[chg], help="добавить запись в «Историю» (с сегодняшней датой)")
    s.add_argument("--item", required=True)
    s.add_argument("--what", required=True)
    s.add_argument("--who", default="Оркестратор (Opus 5.5)", help="роль, модель, effort или «пользователь»")
    s.add_argument("--commit")
    s.add_argument("--rounds", type=int, help="раунды ревью")
    s.add_argument("--tokens", type=int)
    s.add_argument("--link")
    s.add_argument("--date", help="по умолчанию сегодня")
    s.set_defaults(fn=cmd_log)

    s = sub.add_parser("review", parents=[chg], help="добавить замечание ревью (ID R-<пункт>-<n>)")
    s.add_argument("--item", required=True)
    s.add_argument("--round", type=int, required=True)
    s.add_argument("--severity", required=True, choices=SEVERITIES)
    s.add_argument("--text", required=True)
    s.add_argument("--repro", help="воспроизведение")
    s.add_argument("--by", help="кто нашёл")
    s.set_defaults(fn=cmd_review)

    s = sub.add_parser("review-status", parents=[chg], help="сменить статус замечания")
    s.add_argument("id")
    s.add_argument("status", help="/".join(REVIEW_STATUSES))
    s.add_argument("--comment")
    s.set_defaults(fn=cmd_review_status)

    s = sub.add_parser("decision", parents=[chg], help="записать решение пользователя по развилке")
    s.add_argument("id", help="напр. Р-0")
    s.add_argument("--answer", required=True)
    s.add_argument("--date", help="по умолчанию сегодня")
    s.set_defaults(fn=cmd_decision)

    s = sub.add_parser("parity", help="строки листа «Паритет»: add, set")
    psub = s.add_subparsers(dest="pcmd", required=True, metavar="действие")
    q = psub.add_parser("add", parents=[chg], help="добавить строку паритета (состояние - «не начато»)")
    q.add_argument("--id", required=True, help="напр. P-SCH-001")
    q.add_argument("--area", required=True, help="область")
    q.add_argument("--what", required=True, help="функция / поведение")
    q.add_argument("--source", required=True, choices=["1.1", "К2"], help="источник: 1.1 или К2")
    q.add_argument("--delphi", help="где в Delphi")
    q.add_argument("--fr", help="требование FR")
    q.add_argument("--substage", help="подэтап, закрывающий строку")
    q.add_argument("--test", help="тест")
    q.set_defaults(fn=cmd_parity_add)
    q = psub.add_parser("set", parents=[chg], help="сменить состояние строки паритета")
    q.add_argument("id")
    q.add_argument("state", help="/".join(PARITY_STATES))
    q.add_argument("--test", help="тест")
    q.set_defaults(fn=cmd_parity_set)
    return p


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    args = make_parser().parse_args(argv)
    try:
        return args.fn(args) or 0
    except TrackerError as e:
        print(f"ОШИБКА: {e}")
        return 1
    except LockedError as e:
        print(f"ОШИБКА: {e}")
        return 2
    except FileNotFoundError as e:
        print(f"ОШИБКА: файл не найден: {e.filename or e}")
        return 1
    except yaml.YAMLError as e:
        print(f"ОШИБКА: tracker.yaml не разобран: {e}")
        return 1
    except OSError as e:
        print(f"ОШИБКА: {e}")
        return 1
    except (KeyError, ValueError, TypeError, AttributeError, IndexError) as e:
        print(f"ОШИБКА: некорректные данные ({type(e).__name__}: {e}). Проверьте tracker.yaml.")
        return 1
    except Exception as e:                      # непредвиденное: без трейсбека, но с типом ошибки
        print(f"ОШИБКА: непредвиденный сбой ({type(e).__name__}: {e}). Ничего не записано.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
