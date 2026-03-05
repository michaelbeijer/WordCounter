# WordCounter

Free and open-source word counting software for translators.

**WordCounter** is built to rival paid counting tools while staying simple, transparent, and free forever.  
Created by **Michael Beijer** for real translation workflows.

## Why this project exists

Translators often need reliable counts across multiple file types, but many tools are locked behind subscriptions or expensive licenses. WordCounter is an alternative:

- Free to use
- Open source
- Focused on practical translator needs
- Easy to inspect, adapt, and improve

## Current version

`0.3.0`

## What WordCounter does

- Batch counts supported files:
  - `.docx`, `.pptx`, `.xlsx` (core)
  - `.pdf` (optional — requires `pdfminer.six`)
  - With optional [Apache Tika](https://tika.apache.org/): **50+ additional formats** including `.doc`, `.xls`, `.ppt`, `.rtf`, `.odt`, `.odp`, `.ods`, `.html`, `.xml`, `.txt`, `.epub`, `.srt`, `.xliff`, `.tmx`, `.po`, images (OCR), and many more
- Calculates per-file metrics:
  - Words
  - Characters
  - Characters (no spaces)
  - Numbers
  - Number percentage
  - Sentences
  - Paragraphs
  - Estimated pages
- Includes billing panel:
  - Bill by words, characters, or estimated pages
  - Rate, currency, discount, tax
  - Running total amount
- Exports results:
  - CSV export
  - Markdown export (with full document text included)
  - Fixed-width clipboard report (great for Gmail with a monospace font)
- Reports include extracted document text below the count data

## UX highlights

- `Browse…` lets you choose either individual files or a folder.
- `Count` runs counts directly from selected files or the selected folder.
- `Add files…`, `Remove selected`, and `Remove all` support quick list refinement.

## Install

Requires Python 3.10+ (3.12 tested).

Install dependencies:

```bash
pip install python-docx python-pptx openpyxl pdfminer.six
```

`pdfminer.six` is optional if you do not need PDF support.

### Optional: Apache Tika (50+ extra formats)

Tika unlocks support for legacy Office (.doc, .xls, .ppt), OpenDocument, RTF, HTML, EPUB, subtitles, translation formats (XLIFF, TMX, PO), and more — including OCR for images if [Tesseract](https://github.com/tesseract-ocr/tesseract) is installed.

**Requires Java (JRE 8+).** On first run, `tika-python` downloads the Tika server JAR (~70 MB).

```bash
pip install tika
```

Without Tika, WordCounter still works for .docx, .pptx, .xlsx, and .pdf.

## Run

From the project root:

```bash
python WordCounter.py
```

## Test assets

A starter test set is included in:

- `test_assets/`

It contains sample Word documents, a PowerPoint file, and a PDF for quick verification.

## Comparable tools

WordCounter aims to be a free, open-source replacement for commercial word counting tools used by translators. Here is how it compares:

| Tool | Price | Platform | Formats | Invoicing | Status | Notes |
|------|-------|----------|---------|-----------|--------|-------|
| **[WordCounter](https://github.com/michaelbeijer/WordCounter)** | Free (MIT) | Windows, macOS, Linux | 4 core + 50 via Tika | Billing panel (rate/tax/discount) | Active | Open source, lightweight, cross-platform via Python |
| **[AnyCount](https://www.anycount.com/)** | EUR 89-399/yr or EUR 199-399 perpetual | Windows only | 70+ formats incl. OCR, CAT files, URLs | No (separate via TO3000) | Active | Most feature-rich; expensive; heavy (6 GB RAM) |
| **[PractiCount](https://practiline.com/)** | ~USD 60 one-time | Windows only | 20+ formats | Yes, built-in with client DB | Low activity | Good value; dated UI; last major update references Office 2016 |
| **[FineCount](https://www.finecount.eu/)** | EUR 39/yr subscription | Windows only | ~15 formats | Basic quoting/invoicing | Maintenance mode | No major version update since 2018; subscription-only |
| **[CountAnything](https://ginstrom.com/CountAnything/)** | Free | Windows only | ~12 formats | No | Dormant | Freeware but not open source; bare-bones; tiny user base |

**Key differences:**

- **All four commercial/freeware alternatives are Windows-only.** WordCounter runs anywhere Python does.
- **None of the alternatives are open source.** WordCounter can be inspected, modified, and extended by anyone.
- **AnyCount** is the most powerful but also the most expensive, targeting agencies and high-volume translators who need 70+ format support, OCR, and CAT file counting.
- **PractiCount** offers the best value for translators who want integrated invoicing with a one-time purchase, but development has slowed and the UI feels dated.
- **FineCount** was a popular mid-range option, but development appears stalled since 2018 and it requires an ongoing subscription with no perpetual license option.
- **CountAnything** is the only other free option, but it is closed-source, minimally maintained, and lacks billing features.

## Roadmap ideas

- Match-count categories (repetitions / fuzzy bands)
- Better PDF structure extraction and cleanup
- Persist user profiles and presets
- Cross-platform packaged binaries
- Plugin architecture for custom counting rules

## Contributing

Issues, suggestions, and pull requests are welcome.

If you are a translator, your real-world feedback is especially valuable.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
