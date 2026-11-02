"""Проверка относительных ссылок в Markdown-документах репозитория.

Ищет ссылки вида [текст](путь) и [текст](путь#якорь) во всех *.md, кроме
служебных папок, и проверяет, что файл существует, а якорь есть среди
заголовков целевого файла. Внешние ссылки (http, https, mailto) не проверяет.

Запуск: python tools/check_links.py [корень]   (по умолчанию — корень репозитория)
Код возврата: 0 — всё в порядке, 1 — есть битые ссылки.
"""
import re
import sys
from pathlib import Path

SKIP_DIRS = {".git", "node_modules", "target", "graphify-out", "local", ".backup"}
LINK_RE = re.compile(r"(?<!!)\[[^\]]*\]\(([^)\s]+)\)")
CODE_FENCE_RE = re.compile(r"^```.*?^```", re.S | re.M)
INLINE_CODE_RE = re.compile(r"`[^`\n]*`")


def slug(heading: str) -> str:
    """Якорь заголовка в стиле GitHub: нижний регистр, пробелы в дефисы, без знаков."""
    s = heading.strip().lower()
    s = re.sub(r"[^\w\- ]", "", s)
    return s.replace(" ", "-")


def anchors(path: Path) -> set[str]:
    text = CODE_FENCE_RE.sub("", path.read_text(encoding="utf-8-sig"))
    return {slug(m.group(1)) for m in re.finditer(r"^#{1,6}\s+(.+?)\s*#*\s*$", text, re.M)}


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    root = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent)
    bad = 0
    files = [p for p in root.rglob("*.md") if not SKIP_DIRS.intersection(p.relative_to(root).parts)]
    for md in sorted(files):
        text = CODE_FENCE_RE.sub("", md.read_text(encoding="utf-8-sig"))
        text = INLINE_CODE_RE.sub("", text)
        for m in LINK_RE.finditer(text):
            target = m.group(1)
            if re.match(r"^[a-z]+:", target) or target.startswith("#") and not target[1:]:
                continue
            file_part, _, anchor = target.partition("#")
            dest = (md.parent / file_part).resolve() if file_part else md
            line = text.count("\n", 0, m.start()) + 1
            if not dest.exists():
                print(f"{md.relative_to(root)}:{line}: нет файла {target}")
                bad += 1
            elif anchor and dest.suffix == ".md" and anchor.lower() not in anchors(dest):
                print(f"{md.relative_to(root)}:{line}: нет якоря #{anchor} в {file_part or md.name}")
                bad += 1
    print(f"Проверено файлов: {len(files)}, битых ссылок: {bad}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
