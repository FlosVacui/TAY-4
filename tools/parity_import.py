"""Перенос строк «Паритета» из docs/requirements/fr/*.md в tracker.yaml (подэтап 1.1).

Запуск из папки TAY-4: python tools/parity_import.py [--dry] docs/requirements/fr/a.md ...
Формат строки: | ID | Область | Поведение | Источник | Delphi | FR | Подэтап |
Подэтап для графика работ: первый 6.x из списка (этап 6 закрывает строки «Паритета»), иначе первый.
Скрипт только добавляет строки; изменённые строки правятся командой tracker.py parity set или в yaml.
"""
import argparse, re, sys
from pathlib import Path

sys.path.insert(0, str(Path("tools").resolve()))
import tracker  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("--dry", action="store_true")
ap.add_argument("files", nargs="+")
a = ap.parse_args()

rows, errs = [], []
for f in a.files:
    for n, line in enumerate(Path(f).read_text(encoding="utf-8").splitlines(), 1):
        if not line.startswith("| P-"):
            continue
        c = [x.strip() for x in line.strip().strip("|").split("|")]
        if len(c) != 7:
            errs.append(f"{f}:{n}: {len(c)} столбцов")
            continue
        pid, area, what, src, delphi, fr, sub = c
        src = src.replace("K", "К")
        if src not in ("1.1", "К2"):
            errs.append(f"{f}:{n}: источник {src!r}")
        subs = re.findall(r"\b\d+\.\d+\b", sub)
        pick = next((s for s in subs if s.startswith("6.")), subs[0] if subs else "")
        rows.append(dict(id=pid, area=area, what=what.replace("`", ""), source=src,
                         delphi=delphi.replace("`", ""), fr=fr.replace("`", ""), substage=pick, where=f"{f}:{n}"))

ids = [r["id"] for r in rows]
dups = {i for i in ids if ids.count(i) > 1}
if dups:
    errs.append(f"повторы ID: {sorted(dups)}")
data, _ = tracker.load_data(Path("docs/tracker/tracker.yaml"))
items = {str(i["id"]) for i in data["items"]}
for r in rows:
    if r["substage"] and r["substage"] not in items:
        errs.append(f"{r['where']}: подэтапа {r['substage']} нет в tracker.yaml")
print(f"строк: {len(rows)}, ошибок: {len(errs)}")
for e in errs:
    print("  ", e)
if a.dry or errs:
    sys.exit(1 if errs else 0)

added = 0
for r in rows:
    ns = argparse.Namespace(yaml=None, out=None, no_build=True, id=r["id"], area=r["area"], what=r["what"],
                            source=r["source"], delphi=r["delphi"] or None, fr=r["fr"] or None,
                            substage=r["substage"] or None, test=None)
    tracker.cmd_parity_add(ns)
    added += 1
print(f"добавлено: {added}")
