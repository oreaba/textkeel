"""Legacy Word (.doc): converted to .docx by a local Microsoft Word, then read as .docx.

Word is driven through ``word_com.convert_with_word``: a hidden instance,
macros disabled, no dialogs, opening a private copy read-only. The original file is
never opened by Word. Conversions are cached by the source's SHA-256 when a cache folder
is given, so reruns do not start Word again. The Office check runs against the
converted file's XML.

Where Word cannot run (not Windows, Word not installed, or ``use_word=False``), the file
goes to review with ``needs_conversion``; any other conversion failure gives
``conversion_failed``. LibreOffice is never used.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
from pathlib import Path

from textkeel import check
from textkeel.model import Document
from textkeel.readers import docx_native
from textkeel.readers.word_com import ConversionError, convert_with_word

TIMEOUT_SECONDS = 180
NO_WORD = ("word_unavailable", "powershell_unavailable")
HINT = "convert to .docx first, e.g. with Word"


def word_possible() -> bool:
    return os.name == "nt"


def convert(source: Path, target: Path, work_dir: Path) -> str | None:
    """Convert a private copy to .docx at ``target``; return Word's version."""
    return convert_with_word(source, target, work_dir, TIMEOUT_SECONDS)


def _review(path: Path, reason: str, **meta) -> Document:
    return Document(str(path), "doc", "none", status="review", reasons=[reason],
                    meta={"hint": HINT, **meta})


def _derive(path: Path, digest: str, cache: Path) -> tuple[Path, str | None]:
    derived, version_file = cache / f"{digest}.docx", cache / f"{digest}.word_version"
    if derived.is_file():
        return derived, (version_file.read_text("ascii").strip() if version_file.is_file()
                         else None) or None
    private, partial = cache / f"{digest}.doc", cache / f"{digest}.partial.docx"
    shutil.copyfile(path, private)
    try:
        version = convert(private, partial, cache)
        os.replace(partial, derived)
    finally:
        partial.unlink(missing_ok=True)
        private.unlink(missing_ok=True)
    if version:
        version_file.write_text(version, encoding="ascii")
    return derived, version


def read(path: Path, cache: Path | None = None, use_word: bool = True) -> Document:
    if not (use_word and word_possible()):
        return _review(path, "needs_conversion")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory() as tmp:
        folder = cache if cache is not None else Path(tmp)
        folder.mkdir(parents=True, exist_ok=True)
        try:
            derived, version = _derive(path, digest, folder)
        except ConversionError as error:
            reason = "needs_conversion" if error.reason in NO_WORD else "conversion_failed"
            return _review(path, reason, conversion_error=error.reason)
        document = docx_native.read(derived)
        check.apply(document, derived)
        docx_sha256 = hashlib.sha256(derived.read_bytes()).hexdigest()
    document.source, document.format, document.reader = str(path), "doc", "word+native"
    document.meta["converted"] = {"via": "word", "word_version": version,
                                  "docx_sha256": docx_sha256}
    return document
