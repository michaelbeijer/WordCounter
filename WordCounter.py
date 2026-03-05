"""
Batch word counter with GUI for translators (FineCount-inspired).

Counts: .docx, .pptx, .xlsx, optional .pdf
Adds: richer stats columns + billing panel (rate/tax/discount) + add/remove/recount controls

Install:
  pip install python-docx python-pptx openpyxl
Optional PDF:
  pip install pdfminer.six
"""

import os
import re
import threading
import queue
import csv
from datetime import datetime
from dataclasses import dataclass
from typing import List, Optional, Tuple, Dict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Optional imports ---
DOCX_OK = PPTX_OK = XLSX_OK = PDF_OK = False

try:
    from docx import Document
    DOCX_OK = True
except Exception:
    DOCX_OK = False

try:
    from pptx import Presentation
    try:
        from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
        PPTX_PLACEHOLDER_TYPES = PP_PLACEHOLDER_TYPE
    except Exception:
        PPTX_PLACEHOLDER_TYPES = None
    PPTX_OK = True
except Exception:
    PPTX_OK = False

try:
    import openpyxl
    XLSX_OK = True
except Exception:
    XLSX_OK = False

try:
    from pdfminer.high_level import extract_text as pdf_extract_text
    PDF_OK = True
except Exception:
    PDF_OK = False

    APP_NAME = "WordCounter"
    APP_AUTHOR = "Michael Beijer"
    APP_VERSION = "0.1.0"


# ---------------- Tokenisation/stat helpers ----------------
WORD_RE = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9]+(?:['’][A-Za-zÀ-ÖØ-öø-ÿ0-9]+)?")
NUM_RE = re.compile(r"\b\d+(?:[.,]\d+)?\b")

def safe_join_text(parts: List[str]) -> str:
    return "\n".join(p for p in parts if p and p.strip())

def count_words(text: str) -> int:
    return len(WORD_RE.findall(text or ""))

def count_chars_with_spaces(text: str) -> int:
    return len(text or "")

def count_chars_no_spaces(text: str) -> int:
    return len(re.sub(r"\s+", "", text or ""))

def count_numbers(text: str) -> int:
    return len(NUM_RE.findall(text or ""))

def count_sentences(text: str) -> int:
    # Simple heuristic: count sentence end punctuation.
    t = (text or "").strip()
    if not t:
        return 0
    # Split on . ! ? followed by whitespace/end; avoid counting ellipses heavily.
    parts = re.split(r"(?<=[.!?])\s+", t)
    return sum(1 for p in parts if p.strip())

def count_paragraphs(text: str) -> int:
    # paragraphs separated by blank lines
    t = (text or "").strip()
    if not t:
        return 0
    blocks = re.split(r"\n\s*\n+", t)
    return sum(1 for b in blocks if b.strip())


# ---------------- Settings ----------------
@dataclass
class Settings:
    include_subfolders: bool = True

    # DOCX
    docx_include_body: bool = True
    docx_include_tables: bool = True
    docx_include_headers: bool = False
    docx_include_footers: bool = False

    # PPTX
    pptx_include_slide_text: bool = True
    pptx_include_footer_placeholders: bool = False
    pptx_include_speaker_notes: bool = False

    # XLSX
    xlsx_include_text: bool = True
    xlsx_include_numbers: bool = False
    xlsx_include_comments: bool = False
    xlsx_include_hidden_sheets: bool = False

    # PDF
    pdf_include: bool = True
    pdf_remove_repeating_headers_footers: bool = True

    # Pages estimate
    words_per_page: int = 330  # common translation estimate; adjustable


# ---------------- DOCX ----------------
def _docx_collect(container, include_tables: bool) -> List[str]:
    parts: List[str] = []
    for p in getattr(container, "paragraphs", []):
        t = getattr(p, "text", "")
        if t and t.strip():
            parts.append(t)
    if include_tables:
        for table in getattr(container, "tables", []):
            for row in table.rows:
                for cell in row.cells:
                    ct = getattr(cell, "text", "")
                    if ct and ct.strip():
                        parts.append(ct)
    return parts

def docx_text(path: str, s: Settings) -> str:
    doc = Document(path)
    parts: List[str] = []

    if s.docx_include_body:
        parts.extend(_docx_collect(doc, include_tables=s.docx_include_tables))

    if s.docx_include_headers or s.docx_include_footers:
        for sec in doc.sections:
            if s.docx_include_headers:
                parts.extend(_docx_collect(sec.header, include_tables=s.docx_include_tables))
            if s.docx_include_footers:
                parts.extend(_docx_collect(sec.footer, include_tables=s.docx_include_tables))

    return safe_join_text(parts)

def extract_docx(path: str, s: Settings) -> Tuple[str, Optional[str]]:
    if not DOCX_OK:
        return "", "python-docx not installed"
    try:
        return docx_text(path, s), None
    except Exception as e:
        return "", f"DOCX error: {e}"


# ---------------- PPTX ----------------
def is_footer_placeholder(shape) -> bool:
    try:
        if not shape.is_placeholder:
            return False
    except Exception:
        return False
    try:
        pht = shape.placeholder_format.type
        if PPTX_PLACEHOLDER_TYPES is not None:
            footer_types = {
                PPTX_PLACEHOLDER_TYPES.DATE_AND_TIME,
                PPTX_PLACEHOLDER_TYPES.FOOTER,
                PPTX_PLACEHOLDER_TYPES.SLIDE_NUMBER,
            }
            return pht in footer_types
        name = str(pht).upper()
        return any(k in name for k in ["DATE", "FOOTER", "SLIDE_NUMBER"])
    except Exception:
        return False

def pptx_text(path: str, s: Settings) -> str:
    prs = Presentation(path)
    parts: List[str] = []

    for slide in prs.slides:
        if s.pptx_include_slide_text:
            for shape in slide.shapes:
                if (not s.pptx_include_footer_placeholders) and is_footer_placeholder(shape):
                    continue
                try:
                    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                        t = shape.text
                        if t and t.strip():
                            parts.append(t)
                except Exception:
                    continue

        if s.pptx_include_speaker_notes:
            try:
                ns = slide.notes_slide
                if ns and ns.notes_text_frame:
                    nt = ns.notes_text_frame.text
                    if nt and nt.strip():
                        parts.append(nt)
            except Exception:
                pass

    return safe_join_text(parts)

def extract_pptx(path: str, s: Settings) -> Tuple[str, Optional[str]]:
    if not PPTX_OK:
        return "", "python-pptx not installed"
    try:
        return pptx_text(path, s), None
    except Exception as e:
        return "", f"PPTX error: {e}"


# ---------------- XLSX ----------------
def xlsx_text(path: str, s: Settings) -> str:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    parts: List[str] = []

    for ws in wb.worksheets:
        if (not s.xlsx_include_hidden_sheets) and getattr(ws, "sheet_state", "visible") != "visible":
            continue

        for row in ws.iter_rows(values_only=False):
            for cell in row:
                v = cell.value

                if s.xlsx_include_text and isinstance(v, str) and v.strip():
                    parts.append(v)

                if s.xlsx_include_numbers and isinstance(v, (int, float)):
                    parts.append(str(v))

                if s.xlsx_include_comments:
                    try:
                        cmt = cell.comment
                        if cmt and cmt.text and cmt.text.strip():
                            parts.append(cmt.text)
                    except Exception:
                        pass

    return safe_join_text(parts)

def extract_xlsx(path: str, s: Settings) -> Tuple[str, Optional[str]]:
    if not XLSX_OK:
        return "", "openpyxl not installed"
    try:
        return xlsx_text(path, s), None
    except Exception as e:
        return "", f"XLSX error: {e}"


# ---------------- PDF ----------------
def _remove_repeating_lines(text: str) -> str:
    pages = text.split("\f")
    if len(pages) <= 1:
        return text

    line_freq: Dict[str, int] = {}
    page_norm_lines: List[List[str]] = []

    for p in pages:
        lines = [ln.strip() for ln in p.splitlines() if ln.strip()]
        norm = [re.sub(r"\d+", "#", ln) for ln in lines]
        page_norm_lines.append(norm)
        for ln in set(norm):
            line_freq[ln] = line_freq.get(ln, 0) + 1

    threshold = max(2, int(0.6 * len(pages)))
    repeating = {ln for ln, n in line_freq.items() if n >= threshold}

    cleaned_pages: List[str] = []
    for original_page in pages:
        cleaned_lines: List[str] = []
        for ln in original_page.splitlines():
            stripped = ln.strip()
            if not stripped:
                continue
            norm = re.sub(r"\d+", "#", stripped)
            if norm in repeating:
                continue
            cleaned_lines.append(stripped)
        cleaned_pages.append("\n".join(cleaned_lines))

    return "\n".join(cleaned_pages)

def extract_pdf(path: str, s: Settings) -> Tuple[str, Optional[str]]:
    if not s.pdf_include:
        return "", "PDF disabled"
    if not PDF_OK:
        return "", "pdfminer.six not installed (PDF skipped)"
    try:
        text = pdf_extract_text(path) or ""
        if s.pdf_remove_repeating_headers_footers:
            text = _remove_repeating_lines(text)
        return text, None
    except Exception as e:
        return "", f"PDF error: {e}"


# ---------------- Batch + metrics ----------------
SUPPORTED_EXTS = {".docx", ".pptx", ".xlsx", ".pdf"}

@dataclass
class FileMetrics:
    filepath: str
    words: int
    chars: int
    chars_nospace: int
    numbers: int
    sentences: int
    paragraphs: int
    pages_est: float
    note: str = ""

def compute_metrics(text: str, s: Settings) -> Tuple[int, int, int, int, int, int, float]:
    w = count_words(text)
    c = count_chars_with_spaces(text)
    cns = count_chars_no_spaces(text)
    n = count_numbers(text)
    sent = count_sentences(text)
    para = count_paragraphs(text)
    pages = (w / s.words_per_page) if s.words_per_page > 0 else 0.0
    return w, c, cns, n, sent, para, pages

def extract_text_by_type(path: str, s: Settings) -> Tuple[str, Optional[str]]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return extract_docx(path, s)
    if ext == ".pptx":
        return extract_pptx(path, s)
    if ext == ".xlsx":
        return extract_xlsx(path, s)
    if ext == ".pdf":
        return extract_pdf(path, s)
    return "", "Unsupported file type"

def iter_files(folder: str, include_subfolders: bool, include_pdfs: bool) -> List[str]:
    def ok_ext(fn: str) -> bool:
        ext = os.path.splitext(fn)[1].lower()
        if ext == ".pdf" and not include_pdfs:
            return False
        return ext in SUPPORTED_EXTS

    paths: List[str] = []
    if include_subfolders:
        for root, _, files in os.walk(folder):
            for fn in files:
                if ok_ext(fn):
                    paths.append(os.path.join(root, fn))
    else:
        for fn in os.listdir(folder):
            p = os.path.join(folder, fn)
            if os.path.isfile(p) and ok_ext(fn):
                paths.append(p)
    return sorted(paths)

def filter_supported(files: List[str], include_pdfs: bool) -> List[str]:
    out = []
    for p in files:
        ext = os.path.splitext(p)[1].lower()
        if ext not in SUPPORTED_EXTS:
            continue
        if ext == ".pdf" and not include_pdfs:
            continue
        out.append(p)
    return sorted(list(dict.fromkeys(out)))  # de-dupe preserve order


# ---------------- GUI ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}, by {APP_AUTHOR}")
        self.geometry("1200x760")

        self.folder_var = tk.StringVar(value="")
        self._queue = queue.Queue()
        self._results: List[FileMetrics] = []
        self._file_list: List[str] = []

        # --- Settings vars (counting) ---
        self.include_subfolders_var = tk.BooleanVar(value=True)

        self.docx_body_var = tk.BooleanVar(value=True)
        self.docx_tables_var = tk.BooleanVar(value=True)
        self.docx_headers_var = tk.BooleanVar(value=False)
        self.docx_footers_var = tk.BooleanVar(value=False)

        self.pptx_slide_text_var = tk.BooleanVar(value=True)
        self.pptx_footer_ph_var = tk.BooleanVar(value=False)
        self.pptx_notes_var = tk.BooleanVar(value=False)

        self.xlsx_text_var = tk.BooleanVar(value=True)
        self.xlsx_numbers_var = tk.BooleanVar(value=False)
        self.xlsx_comments_var = tk.BooleanVar(value=False)
        self.xlsx_hidden_sheets_var = tk.BooleanVar(value=False)

        self.pdf_include_var = tk.BooleanVar(value=True)
        self.pdf_strip_repeat_var = tk.BooleanVar(value=True)

        self.words_per_page_var = tk.IntVar(value=330)

        # --- Billing vars ---
        self.bill_by_var = tk.StringVar(value="Words")  # Words / Characters / Pages (est.)
        self.rate_var = tk.DoubleVar(value=0.0)
        self.currency_var = tk.StringVar(value="GBP")
        self.tax_var = tk.DoubleVar(value=0.0)       # percent
        self.discount_var = tk.DoubleVar(value=0.0)  # percent

        self._build_ui()
        self._poll_queue()

    def _dependency_status(self) -> str:
        bits = [
            f"DOCX: {'OK' if DOCX_OK else 'missing'}",
            f"PPTX: {'OK' if PPTX_OK else 'missing'}",
            f"XLSX: {'OK' if XLSX_OK else 'missing'}",
            f"PDF: {'OK' if PDF_OK else 'missing (optional)'}",
        ]
        return f"v{APP_VERSION} • " + " • ".join(bits)

    def _get_settings(self) -> Settings:
        return Settings(
            include_subfolders=self.include_subfolders_var.get(),
            docx_include_body=self.docx_body_var.get(),
            docx_include_tables=self.docx_tables_var.get(),
            docx_include_headers=self.docx_headers_var.get(),
            docx_include_footers=self.docx_footers_var.get(),
            pptx_include_slide_text=self.pptx_slide_text_var.get(),
            pptx_include_footer_placeholders=self.pptx_footer_ph_var.get(),
            pptx_include_speaker_notes=self.pptx_notes_var.get(),
            xlsx_include_text=self.xlsx_text_var.get(),
            xlsx_include_numbers=self.xlsx_numbers_var.get(),
            xlsx_include_comments=self.xlsx_comments_var.get(),
            xlsx_include_hidden_sheets=self.xlsx_hidden_sheets_var.get(),
            pdf_include=self.pdf_include_var.get(),
            pdf_remove_repeating_headers_footers=self.pdf_strip_repeat_var.get(),
            words_per_page=int(self.words_per_page_var.get() or 330),
        )

    def _build_ui(self):
        # Top bar
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.folder_var).grid(row=0, column=1, sticky="we", padx=(8, 8))
        ttk.Button(top, text="Browse…", command=self.browse).grid(row=0, column=2)

        ttk.Button(top, text="Add files…", command=self.add_files).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Button(top, text="Remove selected", command=self.remove_selected).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(6, 0))
        ttk.Button(top, text="Remove all", command=self.remove_all).grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(6, 0))

        self.run_btn = ttk.Button(top, text="Count", command=self.run_count)
        self.run_btn.grid(row=0, column=3, padx=(10, 0))

        self.export_btn = ttk.Button(top, text="Export CSV…", command=self.export_csv, state="disabled")
        self.export_btn.grid(row=0, column=4, padx=(8, 0))

        self.copy_btn = ttk.Button(top, text="Copy report", command=self.copy_report, state="disabled")
        self.copy_btn.grid(row=0, column=5, padx=(8, 0))

        top.columnconfigure(1, weight=1)

        ttk.Checkbutton(top, text="Include subfolders (for folder counts)", variable=self.include_subfolders_var)\
            .grid(row=1, column=4, columnspan=2, sticky="w", padx=(12, 0), pady=(6, 0))

        # Settings + Billing panel
        mid = ttk.Frame(self, padding=(10, 0, 10, 6))
        mid.pack(fill="x")

        settings = ttk.LabelFrame(mid, text="Counting settings", padding=10)
        settings.pack(side="left", fill="x", expand=True)

        billing = ttk.LabelFrame(mid, text="Billing", padding=10)
        billing.pack(side="right", fill="y", padx=(10, 0))

        # Settings grid
        # DOCX
        docx = ttk.LabelFrame(settings, text="Word (.docx)", padding=8)
        docx.grid(row=0, column=0, sticky="nwe", padx=(0, 8))
        ttk.Checkbutton(docx, text="Body", variable=self.docx_body_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(docx, text="Tables", variable=self.docx_tables_var).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(docx, text="Headers", variable=self.docx_headers_var).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(docx, text="Footers", variable=self.docx_footers_var).grid(row=3, column=0, sticky="w")

        # PPTX
        pptx = ttk.LabelFrame(settings, text="PowerPoint (.pptx)", padding=8)
        pptx.grid(row=0, column=1, sticky="nwe", padx=(0, 8))
        ttk.Checkbutton(pptx, text="Slide text", variable=self.pptx_slide_text_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(pptx, text="Footer/date/slide # placeholders", variable=self.pptx_footer_ph_var)\
            .grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(pptx, text="Speaker notes", variable=self.pptx_notes_var).grid(row=2, column=0, sticky="w")

        # XLSX
        xlsx = ttk.LabelFrame(settings, text="Excel (.xlsx)", padding=8)
        xlsx.grid(row=0, column=2, sticky="nwe")
        ttk.Checkbutton(xlsx, text="Cell text", variable=self.xlsx_text_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(xlsx, text="Numbers", variable=self.xlsx_numbers_var).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(xlsx, text="Comments/notes", variable=self.xlsx_comments_var).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(xlsx, text="Hidden sheets", variable=self.xlsx_hidden_sheets_var).grid(row=3, column=0, sticky="w")

        # PDF + pages estimate
        bottom_settings = ttk.Frame(settings)
        bottom_settings.grid(row=1, column=0, columnspan=3, sticky="we", pady=(8, 0))
        ttk.Checkbutton(bottom_settings, text="Include PDFs", variable=self.pdf_include_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(bottom_settings, text="Remove repeating header/footer lines (PDF heuristic)", variable=self.pdf_strip_repeat_var)\
            .grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Label(bottom_settings, text="Words per page (estimate):").grid(row=0, column=2, sticky="e", padx=(18, 6))
        ttk.Spinbox(bottom_settings, from_=100, to=1000, textvariable=self.words_per_page_var, width=6)\
            .grid(row=0, column=3, sticky="w")

        settings.columnconfigure(0, weight=1)
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(2, weight=1)

        # Billing controls
        ttk.Label(billing, text="Bill by:").grid(row=0, column=0, sticky="w")
        bill_by = ttk.Combobox(billing, textvariable=self.bill_by_var, values=["Words", "Characters", "Pages (est.)"], state="readonly", width=14)
        bill_by.grid(row=0, column=1, sticky="w")

        ttk.Label(billing, text="Rate:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(billing, textvariable=self.rate_var, width=10).grid(row=1, column=1, sticky="w", pady=(6, 0))

        ttk.Label(billing, text="Currency:").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(billing, textvariable=self.currency_var, width=10).grid(row=2, column=1, sticky="w", pady=(6, 0))

        ttk.Label(billing, text="Tax %:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(billing, textvariable=self.tax_var, width=10).grid(row=3, column=1, sticky="w", pady=(6, 0))

        ttk.Label(billing, text="Discount %:").grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(billing, textvariable=self.discount_var, width=10).grid(row=4, column=1, sticky="w", pady=(6, 0))

        ttk.Separator(billing).grid(row=5, column=0, columnspan=2, sticky="we", pady=10)

        self.total_billable_var = tk.StringVar(value="Billable units: 0")
        self.total_amount_var = tk.StringVar(value="Total: 0.00")
        ttk.Label(billing, textvariable=self.total_billable_var).grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Label(billing, textvariable=self.total_amount_var, font=("Segoe UI", 11, "bold")).grid(row=7, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # Dependency/status line
        self.status_var = tk.StringVar(value=self._dependency_status())
        ttk.Label(self, textvariable=self.status_var, padding=(10, 0)).pack(fill="x")

        # Results table
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)

        cols = ("words", "chars", "chars_ns", "nums", "nums_pct", "sent", "para", "pages", "note", "path")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings")

        headings = {
            "words": "Words",
            "chars": "Chars",
            "chars_ns": "Chars (no sp)",
            "nums": "Numbers",
            "nums_pct": "% nums",
            "sent": "Sent.",
            "para": "Para.",
            "pages": "Pages (est.)",
            "note": "Note",
            "path": "File",
        }
        for k, h in headings.items():
            self.tree.heading(k, text=h)

        self.tree.column("words", width=80, anchor="e")
        self.tree.column("chars", width=90, anchor="e")
        self.tree.column("chars_ns", width=110, anchor="e")
        self.tree.column("nums", width=80, anchor="e")
        self.tree.column("nums_pct", width=70, anchor="e")
        self.tree.column("sent", width=60, anchor="e")
        self.tree.column("para", width=60, anchor="e")
        self.tree.column("pages", width=90, anchor="e")
        self.tree.column("note", width=220, anchor="w")
        self.tree.column("path", width=520, anchor="w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Totals bar
        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        self.total_files_var = tk.StringVar(value="Files: 0")
        self.total_words_var = tk.StringVar(value="Total words: 0")
        ttk.Label(bottom, textvariable=self.total_files_var).pack(side="left")
        ttk.Label(bottom, textvariable=self.total_words_var).pack(side="left", padx=(16, 0))

        # Recompute billing when these change
        for v in (self.bill_by_var, self.rate_var, self.currency_var, self.tax_var, self.discount_var, self.words_per_page_var):
            v.trace_add("write", lambda *_: self.update_billing())

    # -------- File list management --------
    def _pick_files_dialog(self):
        return filedialog.askopenfilenames(
            title="Select files",
            filetypes=[
                ("Supported", "*.docx *.pptx *.xlsx *.pdf"),
                ("Word", "*.docx"),
                ("PowerPoint", "*.pptx"),
                ("Excel", "*.xlsx"),
                ("PDF", "*.pdf"),
                ("All files", "*.*"),
            ],
        )

    def browse(self):
        choice = messagebox.askyesnocancel(
            "Browse",
            "Select individual files?\n\n"
            "Yes = files\n"
            "No = folder\n"
            "Cancel = do nothing"
        )
        if choice is None:
            return

        s = self._get_settings()
        if choice:
            files = self._pick_files_dialog()
            if files:
                self._file_list = filter_supported(list(files), include_pdfs=s.pdf_include)
                if self._file_list:
                    first_dir = os.path.dirname(self._file_list[0])
                    self.folder_var.set(first_dir)
                self.status_var.set(f"Selected {len(self._file_list)} file(s). Click Count.")
            return

        folder = filedialog.askdirectory(title="Choose a folder")
        if folder:
            self.folder_var.set(folder)
            self._file_list = []
            self.status_var.set("Folder selected. Click Count.")

    def add_folder(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Folder required", "Please select a valid folder first (Browse…).")
            return
        s = self._get_settings()
        paths = iter_files(folder, s.include_subfolders, include_pdfs=s.pdf_include)
        self._file_list = filter_supported(self._file_list + paths, include_pdfs=s.pdf_include)
        self.run_count()

    def add_files(self):
        s = self._get_settings()
        files = self._pick_files_dialog()
        if files:
            self._file_list = filter_supported(self._file_list + list(files), include_pdfs=s.pdf_include)
            self.run_count()

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        paths_to_remove = set()
        for item in sel:
            vals = self.tree.item(item, "values")
            if vals and len(vals) >= 10:
                paths_to_remove.add(vals[9])
        self._file_list = [p for p in self._file_list if p not in paths_to_remove]
        self.run_count()

    def remove_all(self):
        self._file_list = []
        self.clear_results()

    # -------- Counting --------
    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._results = []
        self.export_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")
        self.total_files_var.set("Files: 0")
        self.total_words_var.set("Total words: 0")
        self.total_billable_var.set("Billable units: 0")
        self.total_amount_var.set("Total: 0.00")
        self.status_var.set(self._dependency_status())

    def run_count(self):
        s = self._get_settings()
        # if PDF got disabled, filter list now
        self._file_list = filter_supported(self._file_list, include_pdfs=s.pdf_include)

        if not self._file_list:
            folder = self.folder_var.get().strip()
            if folder and os.path.isdir(folder):
                self._file_list = iter_files(folder, s.include_subfolders, include_pdfs=s.pdf_include)

        if not self._file_list:
            self.clear_results()
            self.status_var.set("No supported files selected/found.")
            return

        # Clear UI for fresh results
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._results = []
        self.export_btn.configure(state="disabled")
        self.copy_btn.configure(state="disabled")

        self.run_btn.configure(state="disabled")
        self.status_var.set("Counting…")

        worker = threading.Thread(target=self._worker, args=(list(self._file_list), s), daemon=True)
        worker.start()

    def _worker(self, paths: List[str], s: Settings):
        self._queue.put(("meta", len(paths)))

        for idx, p in enumerate(paths, start=1):
            text, err = extract_text_by_type(p, s)
            if err:
                metrics = FileMetrics(p, 0, 0, 0, 0, 0, 0, 0.0, err)
            else:
                w, c, cns, n, sent, para, pages = compute_metrics(text, s)
                metrics = FileMetrics(p, w, c, cns, n, sent, para, pages, "")
            self._queue.put(("result", metrics, idx, len(paths)))

        self._queue.put(("done",))

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                kind = msg[0]

                if kind == "meta":
                    n = msg[1]
                    self.total_files_var.set(f"Files: {n}")
                    self.total_words_var.set("Total words: 0")

                elif kind == "result":
                    m: FileMetrics = msg[1]
                    idx, n = msg[2], msg[3]
                    self._results.append(m)

                    # % numbers relative to words (FineCount-ish)
                    nums_pct = (m.numbers / m.words * 100.0) if m.words else 0.0

                    self.tree.insert(
                        "", "end",
                        values=(
                            m.words,
                            m.chars,
                            m.chars_nospace,
                            m.numbers,
                            f"{nums_pct:.2f}%",
                            m.sentences,
                            m.paragraphs,
                            f"{m.pages_est:.2f}",
                            m.note,
                            m.filepath
                        )
                    )

                    total_words = sum(r.words for r in self._results)
                    self.total_words_var.set(f"Total words: {total_words}")
                    self.status_var.set(f"Processed {idx}/{n}")

                    self.update_billing()

                elif kind == "done":
                    self.status_var.set(f"Done. {len(self._results)} file(s).")
                    self.run_btn.configure(state="normal")
                    self.export_btn.configure(state=("normal" if self._results else "disabled"))
                    self.copy_btn.configure(state=("normal" if self._results else "disabled"))
                    self.update_billing()

        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    # -------- Billing --------
    def update_billing(self):
        if not self._results:
            self.total_billable_var.set("Billable units: 0")
            self.total_amount_var.set("Total: 0.00")
            return

        bill_by = self.bill_by_var.get()
        rate = float(self.rate_var.get() or 0.0)
        tax = float(self.tax_var.get() or 0.0)
        discount = float(self.discount_var.get() or 0.0)
        currency = (self.currency_var.get() or "").strip() or "GBP"

        if bill_by == "Words":
            units = sum(r.words for r in self._results)
        elif bill_by == "Characters":
            units = sum(r.chars for r in self._results)
        else:  # Pages (est.)
            units = sum(r.pages_est for r in self._results)

        subtotal = units * rate
        subtotal_after_discount = subtotal * (1.0 - (discount / 100.0))
        total = subtotal_after_discount * (1.0 + (tax / 100.0))

        self.total_billable_var.set(f"Billable units: {units:.2f}" if isinstance(units, float) else f"Billable units: {units}")
        self.total_amount_var.set(f"Total: {currency} {total:.2f}")

    # -------- Export --------
    def _format_clipboard_report(self) -> str:
        if not self._results:
            return ""

        def fit(text: str, width: int) -> str:
            t = str(text or "")
            if len(t) <= width:
                return t
            if width <= 1:
                return t[:width]
            return t[:width - 1] + "…"

        rows = []
        for r in self._results:
            nums_pct = (r.numbers / r.words * 100.0) if r.words else 0.0
            rows.append({
                "file": os.path.basename(r.filepath),
                "words": str(r.words),
                "chars": str(r.chars),
                "chars_ns": str(r.chars_nospace),
                "nums": str(r.numbers),
                "nums_pct": f"{nums_pct:.2f}%",
                "sent": str(r.sentences),
                "para": str(r.paragraphs),
                "pages": f"{r.pages_est:.2f}",
                "note": r.note,
            })

        file_width = max(len("File"), *(len(row["file"]) for row in rows))
        note_width = 28

        cols = [
            ("file", "File", file_width, "l"),
            ("words", "Words", 8, "r"),
            ("chars", "Chars", 8, "r"),
            ("chars_ns", "NoSp", 8, "r"),
            ("nums", "Nums", 6, "r"),
            ("nums_pct", "%Nums", 7, "r"),
            ("sent", "Sent", 6, "r"),
            ("para", "Para", 6, "r"),
            ("pages", "Pages", 7, "r"),
            ("note", "Note", note_width, "l"),
        ]

        def align(value: str, width: int, how: str) -> str:
            v = fit(value, width)
            return v.rjust(width) if how == "r" else v.ljust(width)

        header = " | ".join(align(title, width, "l") for _, title, width, _ in cols)
        sep = "-+-".join("-" * width for _, _, width, _ in cols)
        body = []
        for row in rows:
            line = " | ".join(align(row[key], width, how) for key, _, width, how in cols)
            body.append(line)

        total_words = sum(r.words for r in self._results)
        total_chars = sum(r.chars for r in self._results)
        total_pages = sum(r.pages_est for r in self._results)

        bill_by = self.bill_by_var.get()
        if bill_by == "Words":
            units = sum(r.words for r in self._results)
        elif bill_by == "Characters":
            units = sum(r.chars for r in self._results)
        else:
            units = sum(r.pages_est for r in self._results)

        rate = float(self.rate_var.get() or 0.0)
        tax = float(self.tax_var.get() or 0.0)
        discount = float(self.discount_var.get() or 0.0)
        currency = (self.currency_var.get() or "").strip() or "GBP"
        subtotal = units * rate
        subtotal_after_discount = subtotal * (1.0 - (discount / 100.0))
        total = subtotal_after_discount * (1.0 + (tax / 100.0))

        lines = [
            f"{APP_NAME} v{APP_VERSION} - Count Report",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "",
            header,
            sep,
            *body,
            "",
            f"Files: {len(self._results)}",
            f"Total words: {total_words}",
            f"Total chars: {total_chars}",
            f"Total pages (est.): {total_pages:.2f}",
            f"Billing: {bill_by} | Rate {currency} {rate:.4f} | Discount {discount:.2f}% | Tax {tax:.2f}%",
            f"Total amount: {currency} {total:.2f}",
        ]
        return "\n".join(lines)

    def copy_report(self):
        if not self._results:
            return
        try:
            report = self._format_clipboard_report()
            self.clipboard_clear()
            self.clipboard_append(report)
            self.update()
            messagebox.showinfo("Copied", "Formatted report copied to clipboard.\nTip: use a fixed-width font in Gmail for best alignment.")
        except Exception as e:
            messagebox.showerror("Copy failed", str(e))

    def export_csv(self):
        if not self._results:
            return
        out = filedialog.asksaveasfilename(
            title="Save CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not out:
            return

        try:
            with open(out, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([
                    "file", "words", "chars", "chars_no_spaces", "numbers", "sentences",
                    "paragraphs", "pages_est", "note"
                ])
                for r in self._results:
                    w.writerow([r.filepath, r.words, r.chars, r.chars_nospace, r.numbers,
                                r.sentences, r.paragraphs, f"{r.pages_est:.2f}", r.note])

                w.writerow([])
                total_words = sum(r.words for r in self._results)
                total_chars = sum(r.chars for r in self._results)
                total_pages = sum(r.pages_est for r in self._results)
                w.writerow(["TOTALS", total_words, total_chars, "", "", "", "", f"{total_pages:.2f}", ""])

            messagebox.showinfo("Exported", f"Saved:\n{out}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))


if __name__ == "__main__":
    App().mainloop()