"""Prototype multi-format extractor: native readers, the shared model and the Office gate."""

from __future__ import annotations

from pathlib import Path

import pytest

docx = pytest.importorskip("docx")
openpyxl = pytest.importorskip("openpyxl")
pptx = pytest.importorskip("pptx")

from textkeel import check  # noqa: E402
from textkeel.extract import extract  # noqa: E402
from textkeel.model import Cell, Document, Item, markdown_table  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _kinds(document, layer="body"):
    return [(i.kind, i.level, i.text) for i in document.items if i.layer == layer]


def test_docx_structure_merged_cells_and_furniture(tmp_path):
    d = docx.Document()
    d.add_heading("سياسة البيانات", level=1)
    d.add_paragraph("Body text.")
    d.add_paragraph("First", style="List Bullet")
    d.add_paragraph("Second", style="List Number")
    table = d.add_table(rows=2, cols=3)
    table.cell(0, 0).merge(table.cell(0, 1)).text = "Merged"
    table.cell(0, 2).text = "C"
    table.cell(1, 0).text = "a"
    d.sections[0].header.paragraphs[0].text = "Document Code: X-01"
    d.sections[0].footer.paragraphs[0].text = "Page 1"
    path = tmp_path / "sample.docx"
    d.save(path)
    document = extract(path)
    kinds = _kinds(document)
    assert ("heading", 1, "سياسة البيانات") in kinds
    assert [k for k in kinds if k[0] == "list_item"] == [("list_item", 0, "First"),
                                                          ("list_item", 0, "Second")]
    table_item = next(i for i in document.items if i.kind == "table")
    merged = next(c for c in table_item.cells if c.text == "Merged")
    assert (merged.row, merged.col, merged.col_span) == (0, 0, 2)
    assert [i.text for i in document.items if i.layer == "furniture"] == [
        "Document Code: X-01", "Page 1"]
    assert document.status == "ok" and document.meta["office_check"]["f1"] == 1.0


def test_xlsx_sheet_heading_and_merged_range(tmp_path):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Register"
    sheet["A1"], sheet["C1"] = "Title spanning", "x"
    sheet.merge_cells("A1:B1")
    sheet["A2"], sheet["B2"] = 5.0, "نص"
    path = tmp_path / "sample.xlsx"
    book.save(path)
    document = extract(path)
    assert ("heading", 1, "Register") in _kinds(document)
    table = next(i for i in document.items if i.kind == "table")
    assert (table.rows, table.cols) == (2, 3)
    title = next(c for c in table.cells if c.text == "Title spanning")
    assert title.col_span == 2
    assert any(c.text == "5" for c in table.cells)  # an integer-valued float, as Excel shows it
    assert document.meta["office_check"]["f1"] == 1.0


def test_pptx_title_bullets_and_notes(tmp_path):
    deck = pptx.Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[1])
    slide.shapes.title.text = "Quarterly Review"
    body = slide.placeholders[1].text_frame
    body.text = "Results"
    para = body.add_paragraph()
    para.text, para.level = "Detail", 1
    slide.notes_slide.notes_text_frame.text = "Speaker note"
    path = tmp_path / "sample.pptx"
    deck.save(path)
    document = extract(path)
    kinds = _kinds(document)
    assert kinds[0] == ("heading", 1, "Quarterly Review")
    assert ("list_item", 0, "Detail") in kinds
    assert "Speaker note" in [i.text for i in document.items if i.layer == "notes"]


def test_html_structure_furniture_and_spans():
    document = extract(FIXTURES / "sample.html")
    assert document.items[0].kind == "title"
    assert "must not appear" not in document.text()
    assert {i.text for i in document.items if i.layer == "furniture"} == {
        "الرئيسية | Home | Contact", "This is classified as Confidential"}
    table = next(i for i in document.items if i.kind == "table")
    spans = {c.text: (c.row_span, c.col_span) for c in table.cells}
    assert spans["Definition"] == (1, 2) and spans["ISO"] == (2, 1)
    assert [(i.level, i.text) for i in document.items if i.kind == "list_item"][:3] == [
        (0, "أمن المعلومات"), (1, "Access control"), (1, "Encryption")]


def test_markdown_and_csv():
    md = extract(FIXTURES / "sample.md")
    assert [i.kind for i in md.items] == ["heading", "paragraph", "heading", "list_item",
                                          "list_item", "list_item", "list_item", "list_item",
                                          "table", "code"]
    csv_doc = extract(FIXTURES / "sample.csv")
    table = csv_doc.items[0]
    assert (table.rows, table.cols) == (3, 4) and table.cells[2].text == "الاسم"


def test_doc_and_unknown_formats_go_to_review(tmp_path, monkeypatch):
    from textkeel.readers import doc_word

    (tmp_path / "old.doc").write_bytes(b"\xd0\xcf\x11\xe0")
    (tmp_path / "data.xyz").write_text("x")
    assert extract(tmp_path / "old.doc", use_word=False).reasons == ["needs_conversion"]
    monkeypatch.setattr(doc_word, "word_possible", lambda: False)
    assert extract(tmp_path / "old.doc").reasons == ["needs_conversion"]
    assert extract(tmp_path / "data.xyz").reasons == ["unsupported_format"]


def test_missing_pdf_fonts_stop_the_run_before_any_file_is_read(tmp_path):
    from textkeel.__main__ import EXIT_MISSING_FONTS, main

    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "a.md").write_text("# Title\n", encoding="utf-8")
    (folder / "b.pdf").write_bytes((FIXTURES / "arabic_ligatures.pdf").read_bytes())
    out = tmp_path / "out"
    assert main([str(folder), "--out", str(out), "--fonts-dir", str(tmp_path / "fonts")]
                ) == EXIT_MISSING_FONTS
    assert (out / "fonts_report.csv").is_file() and not (out / "a.md").exists()
    assert main([str(folder), "--out", str(out), "--engine", "native"]) == 0
    assert (out / "a.md" / "a.md").is_file() and (out / "b.pdf" / "b.md").is_file()


def test_each_file_gets_its_own_output_folder(tmp_path):
    from textkeel.__main__ import main

    folder = tmp_path / "in"
    (folder / "sub").mkdir(parents=True)
    (folder / "report.md").write_text("# From Markdown\n", encoding="utf-8")
    (folder / "report.txt").write_text("From text\n", encoding="utf-8")
    (folder / "sub" / "notes.txt").write_text("Nested\n", encoding="utf-8")
    out = tmp_path / "out"
    assert main([str(folder), "--out", str(out)]) == 0
    assert "From Markdown" in (out / "report.md" / "report.txt").read_text(encoding="utf-8")
    assert "From text" in (out / "report.txt" / "report.txt").read_text(encoding="utf-8")
    assert {p.name for p in (out / "report.md").iterdir()} == {
        "report.json", "report.md", "report.txt", "report.chunks.jsonl"}
    assert (out / "sub" / "notes.txt" / "notes.json").is_file()
    single = tmp_path / "single"
    assert main([str(folder / "report.md"), "--out", str(single)]) == 0
    assert (single / "report.md" / "report.json").is_file()


def test_webp_images_are_wrapped_for_ocr(tmp_path, monkeypatch):
    image_module = pytest.importorskip("PIL.Image")
    import pymupdf

    from textkeel.readers import pdf_bridge

    path = tmp_path / "scan.webp"
    image_module.new("RGB", (120, 60), "white").save(path)
    pages = []

    def fake_read_pdf(wrapped, **_kwargs):
        with pymupdf.open(wrapped) as pdf:
            pages.append(pdf.page_count)
        return Document(str(wrapped), "pdf", "pdf-ar:ocr")

    monkeypatch.setattr(pdf_bridge, "read_pdf", fake_read_pdf)
    document = extract(path)
    assert pages == [1] and (document.source, document.format) == (str(path), "webp")


def test_one_failing_file_does_not_stop_the_run(tmp_path, monkeypatch):
    import csv

    from textkeel import __main__ as cli

    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "a.md").write_text("# Title\n", encoding="utf-8")
    (folder / "b.txt").write_text("text\n", encoding="utf-8")
    real = cli.extract

    def flaky(path, *args, **kwargs):
        if path.suffix == ".md":
            raise ValueError("broken")
        return real(path, *args, **kwargs)

    monkeypatch.setattr(cli, "extract", flaky)
    assert cli.main([str(folder), "--out", str(tmp_path / "out"), "--engine", "native"]) == 0
    with (tmp_path / "out" / "summary.csv").open(encoding="utf-8") as handle:
        rows = {Path(r["source"]).name: r for r in csv.DictReader(handle)}
    assert (rows["a.md"]["status"], rows["a.md"]["reasons"]) == ("review", "error:ValueError")
    assert rows["b.txt"]["status"] == "ok"


def test_word_gets_absolute_paths(tmp_path, monkeypatch):
    import subprocess

    from textkeel.readers import word_com

    seen = []

    def fake_powershell(_script, *args, timeout):
        seen.extend(args)
        Path(args[3]).write_bytes(b"docx")
        return subprocess.CompletedProcess(args, 0, stdout="word_version=16.0\n")

    monkeypatch.setattr(word_com, "_powershell", fake_powershell)
    monkeypatch.chdir(tmp_path)
    Path("cache").mkdir()
    assert word_com.convert_with_word(Path("cache/a.doc"), Path("cache/a.docx"),
                                      Path("cache"), timeout=5) == "16.0"
    assert all(Path(p).is_absolute() for p in (seen[1], seen[3]))


def test_doc_converted_by_word_is_read_checked_and_cached(tmp_path, monkeypatch):
    from textkeel.readers import doc_word
    from textkeel.readers.word_com import ConversionError

    calls = []

    def fake_word(source, target, work_dir):
        calls.append(source)
        assert source != original and source.read_bytes() == original.read_bytes()
        d = docx.Document()
        d.add_heading("Legacy procedure", level=1)
        d.add_paragraph("Converted body text.")
        d.save(target)
        return "16.0"

    original = tmp_path / "legacy.doc"
    original.write_bytes(b"\xd0\xcf\x11\xe0 legacy")
    monkeypatch.setattr(doc_word, "word_possible", lambda: True)
    monkeypatch.setattr(doc_word, "convert", fake_word)
    cache = tmp_path / "cache"
    first = extract(original, doc_cache=cache)
    assert (first.format, first.reader, first.status) == ("doc", "word+native", "ok")
    assert first.source == str(original)
    assert ("heading", 1, "Legacy procedure") in _kinds(first)
    assert first.meta["office_check"]["f1"] == 1.0
    assert first.meta["converted"]["word_version"] == "16.0"
    second = extract(original, doc_cache=cache)
    assert len(calls) == 1 and second.text() == first.text()
    assert [p.suffix for p in cache.iterdir() if p.suffix != ".word_version"] == [".docx"]

    def no_word(source, target, work_dir):
        raise ConversionError("word_unavailable")

    def broken(source, target, work_dir):
        raise ConversionError("word_open_failed")

    other = tmp_path / "other.doc"
    other.write_bytes(b"\xd0\xcf\x11\xe0 other")
    monkeypatch.setattr(doc_word, "convert", no_word)
    assert extract(other).reasons == ["needs_conversion"]
    monkeypatch.setattr(doc_word, "convert", broken)
    failed = extract(other)
    assert failed.reasons == ["conversion_failed"]
    assert failed.meta["conversion_error"] == "word_open_failed"


def test_markdown_table_writes_a_merged_cell_once():
    item = Item("table", cells=[Cell(0, 0, "A", 1, 2), Cell(1, 0, "b"), Cell(1, 1, "c")],
                rows=2, cols=2)
    assert markdown_table(item) == "| A |  |\n| --- | --- |\n| b | c |"


def test_chunks_follow_headings():
    document = Document("x", "md", "native")
    document.add("heading", "Intro", level=1)
    document.add("paragraph", "one")
    document.add("heading", "Details", level=2)
    document.add("paragraph", "two")
    document.add("paragraph", "footer", layer="furniture")
    assert document.chunks() == [{"headings": ["Intro"], "text": "one", "pages": []},
                                 {"headings": ["Intro", "Details"], "text": "two", "pages": []}]


def test_pdf_headings_and_lists_from_size_weight_and_markers():
    from textkeel.readers.pdf_bridge import _classify

    def entry(text, size=10.0, bold=0.0, part="pdf", x0=72.0, width=400.0):
        return {"item": Item("paragraph", text, prov=__import__("textkeel.model").model
                             .Provenance(page=1)), "size": size, "bold": bold, "part": part,
                "chars": len(text), "x0": x0, "width": width}

    entries = [entry("Body text that runs long enough to set the size of the page. " * 4),
               entry("Spatial Data Policy", size=16), entry("2.1 Scope", bold=1.0),
               entry("Purpose", bold=1.0, part="shape", width=450),
               entry("Review the request", bold=1.0, part="shape", width=120),
               entry("• First point", x0=90), entry("o Sub point", x0=110),
               entry("1. INTRODUCTION ........ 3", bold=1.0, x0=90)]
    _classify(entries)
    got = [(e["item"].kind, e["item"].level, e["item"].text) for e in entries]
    assert got[1] == ("heading", 1, "Spatial Data Policy")
    assert got[2] == ("heading", 2, "2.1 Scope")  # depth from its numbering
    assert got[3][0] == "heading"  # a wide shaded section bar
    assert got[4][0] == "paragraph"  # a narrow flowchart box
    assert got[5] == ("list_item", 0, "First point") and got[6] == ("list_item", 1, "Sub point")
    assert got[7][0] == "list_item"  # a contents line is never a heading


def test_gate_flags_lost_and_repeated_words(tmp_path):
    d = docx.Document()
    d.add_paragraph("alpha beta gamma delta")
    path = tmp_path / "gate.docx"
    d.save(path)
    assert check.score("alpha beta gamma delta", path)["reasons"] == []
    assert check.score("alpha beta", path)["reasons"] == ["lost_words"]
    assert check.score("alpha beta gamma delta alpha beta", path)["reasons"] == [
        "repeated_words"]
