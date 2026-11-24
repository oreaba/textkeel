# textkeel

One extractor for every common document format, with one structured output close to Docling's. The name pictures a keel: one text backbone under every format.

- **PDFs** go through [`pdf-ar`](https://github.com/oreaba/pdf-ar), unchanged: the same engines, page checks and text. Arabic is rebuilt from the fonts, so Word-export PDFs with broken text maps read correctly.
- **Images** go through `pdf-ar`'s OCR engine.
- **Word, Excel, PowerPoint, HTML, Markdown, CSV and text** are read by textkeel's own readers. Docling is not a dependency.
- **`.doc`** is converted to `.docx` by a local Microsoft Word on Windows, then read as `.docx`. Where Word is not available, as in Docker, it goes to review.

No language model is used, and everything works offline.

## Formats

| Input | Reader |
|---|---|
| `.pdf` | `pdf_bridge`: `pdf-ar` (`fonts`, `native`, `ocr` or `auto`), tables with spans, headings and lists, per-page status |
| `.docx` | `docx_native`: the file's XML, standard library only |
| `.xlsx` | `xlsx_native` (openpyxl): each sheet a heading and a table |
| `.pptx` | `pptx_native` (python-pptx): slides as pages, speaker notes in the `notes` layer |
| `.html`, `.htm` | `html_native` (standard library `html.parser`) |
| `.md`, `.markdown`, `.csv`, `.txt` | `text_native` (standard library) |
| images (`.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`, `.bmp`, `.gif`, `.webp`) | wrapped as a one-page PDF, then `pdf-ar`'s OCR |
| `.doc` | `doc_word`: converted by Word (Windows only), then `docx_native`; otherwise `review` |

Any other suffix goes to `review` with reason `unsupported_format`.

## Legacy Word (`.doc`)

`.doc` is Word's old binary format. Only Word or LibreOffice converts it faithfully, and LibreOffice would add about 400 MB to the image, so textkeel uses a local Word when there is one ([`doc_word.py`](https://github.com/oreaba/textkeel/blob/main/src/textkeel/readers/doc_word.py), [`word_com.py`](https://github.com/oreaba/textkeel/blob/main/src/textkeel/readers/word_com.py)):

- Word runs as a hidden instance, with macros disabled and no dialogs, and opens a private copy read-only. Word never opens the original file.
- The converted `.docx` is read by `docx_native` and checked against its own XML, like any `.docx`. The output keeps the `.doc` as its source, with `format: doc`, `reader: word+native` and `meta.converted` (Word's version and the converted file's SHA-256).
- Conversions are cached by the `.doc`'s SHA-256 in `<out>/.doc-cache` (or `--doc-cache`), so reruns do not start Word again. The cache holds document text: keep it where the outputs are kept.

The file goes to review with:

- `needs_conversion` where Word cannot run: not Windows (Docker included), Word not installed, or `--no-word`. Convert those files with Word elsewhere and extract the `.docx`.
- `conversion_failed` when Word starts but cannot open or save the file; `meta.conversion_error` gives the reason.

In our tests, a conversion took about 6 seconds and a cached rerun under a second.

## The structured output ([`model.py`](https://github.com/oreaba/textkeel/blob/main/src/textkeel/model.py))

A `Document` is an ordered list of `Item`s, in reading order:

- **Kinds:** `title`, `heading` (with level), `paragraph`, `list_item` (level, ordered), `table`, `picture` (alt text or caption), `caption`, `footnote`, `code`.
- **Layers:** `body`, `furniture` (page headers and footers) and `notes` (speaker notes).
- **Provenance:** page, slide or sheet; position when known; the source part, for example `header`, `textbox`, `sheet:Register` or `shape`.
- **Tables:** `rows`, `cols` and `cells`, each cell with row, column, spans and a header flag. A merged cell is stored once.

It exports to:

- JSON (`to_dict`);
- Markdown, with a merged cell written once and the rest of its positions left empty;
- plain text (`text`), which leaves out picture alt text;
- heading-aware chunks for retrieval (`chunks`), each with its heading path and pages.

The document also carries `status` (`ok` or `review`), `reasons` and `meta`: per-page results for PDFs, and the Office check.

## The Office check ([`check.py`](https://github.com/oreaba/textkeel/blob/main/src/textkeel/check.py))

For `.docx`, `.xlsx` and `.pptx`, the extracted words are compared with the text the file itself stores. That is every paragraph's runs, or every cell value and sheet name, skipping `mc:Fallback` duplicates and identical header or footer variants. The check measures:

- recall per part (body, table, text box, header, footer, cell, sheet, slide, notes);
- overall precision.

A document goes to `review` with the reason `lost_words` (recall below 0.98) or `repeated_words` (precision below 0.98).

The readers read the same XML the check compares against, so the check catches reader bugs such as lost text boxes or repeated fallback copies. It cannot catch text the file stores in a place neither side reads.

## PDF headings and lists

`pdf-ar` gives blocks of text. textkeel adds headings and list items from text size, weight and markers ([`pdf_bridge.py`](https://github.com/oreaba/textkeel/blob/main/src/textkeel/readers/pdf_bridge.py) `_classify`):

- a short line at least 15% larger than the body text, or all bold, is a heading;
- a numbered section ("2.1") takes its depth from its number;
- a wide shaded bar counts as a section label;
- bullets and table-of-contents lines are never headings;
- a line that starts with a bullet or a list number is a list item, nested by its indent.

The words stay exactly `pdf-ar`'s. The heading rules are heuristic: check them on a labelled sample of your own documents before relying on them.

## Install

Python 3.11 or later:

```
pip install textkeel            # with pdf-ar, openpyxl, python-pptx
pip install "textkeel[ocr]"     # adds pytesseract
```

Or from a clone of the [repository](https://github.com/oreaba/textkeel): `pip install -e ".[ocr,dev]"`.

OCR also needs Tesseract 5 with the `ara` and `eng` models; see [pdf-ar's install notes](https://github.com/oreaba/pdf-ar#install).

PDFs read with the `fonts` or `auto` engine (the default is `fonts`) need their fonts in the fonts folder, which holds only the files you put there. Before reading any file, textkeel checks every PDF's fonts. If a needed font is missing or a different version, it prints the table, writes `fonts_report.csv` to the output folder, and stops with exit code 2 without extracting anything. `pdf-ar fonts <folder>` prints the same table on its own; see [Fonts](https://github.com/oreaba/pdf-ar#fonts).

## Use

```
textkeel <file or folder> [--out artifacts/textkeel] [--engine auto|native|fonts|ocr]
                          [--font-check] [--fonts-dir fonts]
                          [--no-word] [--doc-cache folder]
```

`python -m textkeel` works the same way. A folder is searched recursively for every supported suffix, skipping the output folder. `--engine` and `--font-check` apply to PDFs and mean the same as in `pdf-ar`. `--no-word` and `--doc-cache` apply to `.doc`.

Per file, it writes:

- `<stem>.json`: the whole document;
- `<stem>.md`: Markdown;
- `<stem>.txt`: body text;
- `<stem>.chunks.jsonl`: retrieval chunks.

For the run, it writes `summary.csv`, with each file's format, status, reasons, item count and seconds.

From Python:

```python
from pathlib import Path
from textkeel.extract import extract

document = extract(Path("report.docx"))
print(document.status, document.reasons)
print(document.markdown())
```

## Docker

The [`Dockerfile`](https://github.com/oreaba/textkeel/blob/main/Dockerfile) builds a CPU-only image with the same Python, PyMuPDF, fontTools and Tesseract models as pdf-ar's image, plus openpyxl and python-pptx. It installs no PyTorch and no models. The image is about 730 MB in `docker images`.

Build it from the repository root:

```
docker build -t textkeel:0.1 .
```

[`.dockerignore`](https://github.com/oreaba/textkeel/blob/main/.dockerignore) keeps fonts, outputs and `.venv` out of the build. The finished image needs no network.

The container works in `/data`. Mount a folder that holds the input and a filled `fonts/` folder:

```
# PowerShell
docker run --rm -v "${PWD}:/data" textkeel:0.1 input --out output
# Linux or macOS
docker run --rm --user "$(id -u):$(id -g)" -v "$PWD:/data" textkeel:0.1 input --out output
```

## Tests

```
pip install -e ".[ocr,dev]"
pytest
```

## Limits

- **Word:** charts, SmartArt, embedded objects, comments and tracked changes are not read.
- **Excel:** formatted numbers (dates, percentages) are written as stored, not as displayed.
- **PDF headings and lists:** heuristic, as above. On 24 English pages they agreed with Docling's layout model on 89% of textkeel's headings and recovered 60% of Docling's; neither side is ground truth.
- **`.doc`:** read only where Word is installed. In Docker or on Linux it needs a manual conversion step.

## License

[AGPL-3.0](https://github.com/oreaba/textkeel/blob/main/LICENSE). textkeel requires pdf-ar and PyMuPDF, which are AGPL-3.0 (PyMuPDF also sells a commercial license).
