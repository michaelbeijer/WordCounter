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

`0.1.0`

## What WordCounter does

- Batch counts supported files:
  - `.docx`
  - `.pptx`
  - `.xlsx`
  - `.pdf` (optional dependency)
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
  - Fixed-width clipboard report (great for Gmail with a monospace font)

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

## Run

From the project root:

```bash
python WordCounter.py
```

## Test assets

A starter test set is included in:

- `test_assets/`

It contains sample Word documents, a PowerPoint file, and a PDF for quick verification.

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

Open source. Add your preferred license file (for example MIT) to finalize licensing.
