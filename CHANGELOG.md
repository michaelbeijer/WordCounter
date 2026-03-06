# Changelog

All notable changes to this project are documented in this file.

## [0.5.0] - 2026-03-06

### Added
- About dialog with version info, author name, clickable website link, and GitHub repository link.
- Version number now shown in the window title bar.

### Changed
- Version bumped to 0.5.0.

## [0.4.0] - 2026-03-05

### Added
- Dedicated parsers for translation file formats: SDLXLIFF, XLIFF, memoQ XLIFF, TMX, PO/POT.
- Source/target toggle for XLIFF/SDLXLIFF counting (default: source segments).
- Auto-detects source and target languages from file metadata; shown in Note column.
- Inline XLIFF tags (bpt, ept, ph, it, x) are stripped; text inside formatting tags (g) is preserved.
- Translation formats now work without Tika — uses built-in XML parser.

### Fixed
- SDLXLIFF files no longer count metadata, hashes, timestamps, and duplicate segments — only actual translatable text is counted.

### Changed
- Version bumped to 0.4.0.

## [0.3.0] - 2026-03-05

### Added
- Bundled Java Runtime (JRE 17, stripped via jlink) and Apache Tika server JAR in the Windows EXE build.
- Users get 50+ file format support out of the box with zero external dependencies.
- `.gitignore` for clean repository.

### Changed
- EXE build switched from single-file to directory (onedir) mode for faster startup.
- PyInstaller spec updated with Tika/JRE data files and hidden imports.
- Version bumped to 0.3.0.

## [0.2.0] - 2026-03-05

### Added
- Apache Tika integration for 50+ additional file formats (.doc, .xls, .ppt, .rtf, .odt, .html, .epub, .xliff, .tmx, .po, and many more).
- Markdown export button alongside CSV export.
- Extracted document text included in all reports (clipboard, CSV, Markdown).
- Improved Browse dialog with clear "Select Files..." and "Select Folder..." buttons.
- Settings persistence across sessions (window size/position, last folder, billing settings).
- Keyboard shortcuts: Ctrl+O (browse), Ctrl+Enter / F5 (count).
- Comparable tools section in README with comparison table.

### Changed
- File format support is now dynamic — automatically expands when Tika is available.
- File dialog shows format categories when Tika is installed.
- Dependency status bar shows Tika availability.
- Version bumped to 0.2.0.

## [0.1.1] - 2026-03-05

### Fixed
- Resolved Windows EXE startup crash (`name 'APP_NAME' is not defined`) by defining app metadata constants at module scope.

### Changed
- Added MIT license to the repository.
- Added release-oriented project metadata updates (`README`, `VERSION`).

## [0.1.0] - 2026-03-05

### Added
- Tkinter-based translator-focused batch word counting app.
- Support for `.docx`, `.pptx`, `.xlsx`, and optional `.pdf` extraction.
- Detailed per-file metrics: words, chars, chars-no-spaces, numbers, %numbers, sentences, paragraphs, estimated pages.
- Billing panel with bill-by mode, rate, tax, discount, currency, and total amount calculation.
- CSV export of detailed results plus totals.
- Clipboard export of fixed-width formatted reports for easy email pasting.
- File/folder selection flow via `Browse…` and simplified `Count` action.
- Test asset folder with sample documents.

### Changed
- App title updated to: `WordCounter, by Michael Beijer`.
- Toolbar and control labels streamlined for clearer workflow.
