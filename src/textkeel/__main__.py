"""Command line: ``textkeel <file or folder> [--out folder] [--engine ...]``."""

from __future__ import annotations

import argparse
import csv
import logging
import sys
from pathlib import Path

from pdf_ar.extract import MissingFontsError
from pdf_ar.required_fonts import table

from textkeel.extract import SUPPORTED, check_pdf_fonts, extract, write

EXIT_MISSING_FONTS = 2


def main(argv: list[str] | None = None) -> int:
    logging.getLogger("fontTools").setLevel(logging.ERROR)
    parser = argparse.ArgumentParser(prog="textkeel")
    parser.add_argument("path", type=Path)
    parser.add_argument("--out", type=Path, default=Path("artifacts/textkeel"))
    parser.add_argument("--engine", choices=("auto", "native", "fonts", "ocr"), default="fonts",
                        help="pdf-ar engine for PDFs")
    parser.add_argument("--fonts-dir", type=Path, default=Path("fonts"))
    parser.add_argument("--font-check", action="store_true")
    parser.add_argument("--no-word", action="store_true",
                        help="never start Word for .doc files; they go to review "
                             "(needs_conversion)")
    parser.add_argument("--doc-cache", type=Path,
                        help="folder for .docx converted from .doc, reused on reruns "
                             "(default: <out>/.doc-cache)")
    args = parser.parse_args(argv)
    doc_cache = args.doc_cache or args.out / ".doc-cache"
    out = args.out.resolve()
    files = [args.path] if args.path.is_file() else sorted(
        p for p in args.path.rglob("*") if p.suffix.lower() in SUPPORTED
        and not p.resolve().is_relative_to(out))
    root = args.path if args.path.is_dir() else args.path.parent
    try:
        fonts = check_pdf_fonts(files, args.engine, args.fonts_dir, args.font_check, root,
                                args.out)
    except MissingFontsError as error:
        print(table(error.all_rows), flush=True)
        print(f"\n{len(error.rows)} needed font(s) are missing or a different version. Add "
              f"them to {error.fonts_dir} and run again; nothing was extracted. Full table: "
              f"{args.out / 'fonts_report.csv'}", file=sys.stderr)
        return EXIT_MISSING_FONTS
    rows = []
    for index, path in enumerate(files, 1):
        print(f"[{index}/{len(files)}] {path.name}", file=sys.stderr)
        document = extract(path, args.engine, args.fonts_dir, args.font_check,
                           doc_cache=doc_cache, use_word=not args.no_word, fonts=fonts)
        write(document, args.out / path.parent.relative_to(args.path if args.path.is_dir()
                                                            else path.parent))
        rows.append({"source": str(path), "format": document.format, "reader": document.reader,
                     "status": document.status, "reasons": ";".join(document.reasons),
                     "items": len(document.items), "seconds": document.meta.get("seconds")})
    args.out.mkdir(parents=True, exist_ok=True)
    with (args.out / "summary.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, list(rows[0]) if rows else ["source"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"{len(rows)} files, {sum(r['status'] != 'ok' for r in rows)} for review")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
