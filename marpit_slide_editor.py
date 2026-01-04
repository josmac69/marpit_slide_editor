
#!/usr/bin/env python3
"""
Marpit / Marp slide deck editor (single-file GUI)

Features
- Deck overview (left): vertical list showing *all* slides in the file, visibly separated.
- Slide editor (middle): edit one slide at a time with formatting buttons.
- Live preview (right): renders the current slide via Marp CLI into HTML and shows it in an embedded browser.

Requirements
- Python 3.9+
- PySide6 (Qt) + Qt WebEngine
- Marp CLI (recommended): https://github.com/marp-team/marp-cli

Install (typical)
  pip install PySide6
  npm install -g @marp-team/marp-cli

Notes
- Preview/export uses Marp CLI. If it is not found in PATH, set env var MARP_CLI, e.g.:
    export MARP_CLI="marp"
    export MARP_CLI="npx marp"        # if marp is locally installed in a Node project
- PDF/PPTX export requires an installed compatible browser (Chrome/Edge/Firefox).

This tool intentionally keeps slide splitting simple and explicit: slides are separated by Markdown horizontal rulers
(---, ***, ___). YAML front-matter at the top (--- ... ---) is treated as deck preamble.
"""
from __future__ import annotations

import hashlib
import os
import re
import shlex
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QTimer, QUrl, QSize
from PySide6.QtGui import QAction, QFont, QKeySequence, QTextCursor, QSyntaxHighlighter, QTextCharFormat, QColor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStatusBar,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QPlainTextEdit,
)

try:
    from PySide6.QtWebEngineWidgets import QWebEngineView  # type: ignore
    WEBENGINE_AVAILABLE = True
except Exception:
    WEBENGINE_AVAILABLE = False
    QWebEngineView = None  # type: ignore

try:
    from spellchecker import SpellChecker
    SPELLCHECK_AVAILABLE = True
except ImportError:
    SPELLCHECK_AVAILABLE = False

if SPELLCHECK_AVAILABLE:
    class SpellCheckHighlighter(QSyntaxHighlighter):
        def __init__(self, document):
            super().__init__(document)
            self.spell = SpellChecker(language='en')

            self.error_format = QTextCharFormat()
            self.error_format.setUnderlineColor(QColor("red"))
            self.error_format.setUnderlineStyle(QTextCharFormat.SpellCheckUnderline)

        def highlightBlock(self, text):
            # Simple word extraction using regex
            import re
            for match in re.finditer(r'\b[a-zA-Z]+\b', text):
                word = match.group()
                if self.spell.unknown([word]):
                    self.setFormat(match.start(), match.end() - match.start(), self.error_format)



_HR_RE = re.compile(r"^\s{0,3}((-\s*){3,}|(\*\s*){3,}|(_\s*){3,})\s*$")
_YAML_DELIM_RE = re.compile(r"^\s{0,3}---\s*$")


def _is_horizontal_rule(line: str) -> bool:
    return bool(_HR_RE.match(line))


def _strip_bom(text: str) -> str:
    return text[1:] if text.startswith("\ufeff") else text


def parse_marp_markdown(text: str) -> Tuple[str, List[str]]:
    """
    Parse a Marp/Marpit Markdown deck into:
      - preamble: YAML front-matter (including delimiters) if present, else "".
      - slides: list of slide markdown strings (excluding slide separators).
    """
    text = _strip_bom(text).replace("\r\n", "\n").replace("\r", "\n")

    # Detect YAML front-matter at the very top
    preamble = ""
    rest = text
    lines = text.split("\n")

    if lines and _YAML_DELIM_RE.match(lines[0] or ""):
        # Search for the closing delimiter
        for i in range(1, len(lines)):
            if _YAML_DELIM_RE.match(lines[i] or ""):
                preamble = "\n".join(lines[: i + 1]).strip("\n") + "\n\n"
                rest = "\n".join(lines[i + 1 :])
                break

    # Split slides by horizontal rulers
    slides: List[str] = []
    buf: List[str] = []
    for ln in rest.split("\n"):
        if _is_horizontal_rule(ln):
            slide_text = "\n".join(buf).strip("\n")
            slides.append(slide_text)
            buf = []
        else:
            buf.append(ln)

    # Last slide
    slide_text = "\n".join(buf).strip("\n")
    slides.append(slide_text)

    # Ensure at least one slide
    if len(slides) == 0:
        slides = [""]

    # Normalize: never keep None; keep as strings
    slides = [s for s in slides]

    return preamble, slides


def serialize_marp_markdown(preamble: str, slides: List[str]) -> str:
    pre = preamble or ""
    # Ensure preamble ends with exactly one blank line if present
    if pre.strip():
        pre = pre.strip("\n") + "\n\n"
    else:
        pre = ""

    parts = []
    for s in slides:
        parts.append((s or "").strip("\n"))
    return pre + "\n\n---\n\n".join(parts).rstrip() + "\n"


def md_line_transform_remove_leading_syntax(line: str) -> str:
    """Remove common leading Markdown syntax for line-level styling."""
    s = line.lstrip()

    # Remove heading markers
    s = re.sub(r"^#{1,6}\s+", "", s)

    # Remove list markers
    s = re.sub(r"^([-+*])\s+", "", s)
    s = re.sub(r"^\d+\.\s+", "", s)

    # Remove blockquote marker
    s = re.sub(r"^>\s+", "", s)

    return s


def find_marp_cli_command() -> Optional[List[str]]:
    """
    Locate Marp CLI command.

    Preference order:
    1) MARP_CLI environment variable (split like a shell command)
    2) marp in PATH
    3) marp-cli in PATH (rare)
    """
    env = os.getenv("MARP_CLI", "").strip()
    if env:
        cmd = shlex.split(env)
        if cmd:
            return cmd

    # Check local bin directory
    local_bin = Path(__file__).parent / "bin" / "marp"
    if local_bin.exists():
        return [str(local_bin)]

    for cand in ("marp", "marp-cli", "marp.cmd", "marp.exe"):
        path = shutil.which(cand)
        if path:
            return [path]

    return None


@dataclass
class DeckState:
    preamble: str
    slides: List[str]
    file_path: Optional[Path] = None
    dirty: bool = False

    @classmethod
    def new_default(cls) -> "DeckState":
        preamble = (
            "---\n"
            "marp: true\n"
            "theme: default\n"
            "paginate: true\n"
            "html: true\n"
            "---\n\n"
        )
        slides = [
            "# Title\n"
            "## Subtitle\n",
        ]
        return cls(preamble=preamble, slides=[s.strip("\n") for s in slides], file_path=None, dirty=False)


class DeckDirectivesDialog(QDialog):
    def __init__(self, parent: QWidget, preamble_text: str):
        super().__init__(parent)
        self.setWindowTitle("Deck directives (YAML front-matter)")
        self.setModal(True)
        self.resize(700, 500)

        layout = QVBoxLayout(self)

        info = QLabel(
            "This is the deck-level YAML front-matter. It must be the first thing in the file.\n"
            "Example:\n"
            "---\n"
            "marp: true\n"
            "theme: default\n"
            "paginate: true\n"
            "---"
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.editor = QPlainTextEdit()
        self.editor.setPlainText(preamble_text.strip("\n"))
        mono = QFont("Courier New")
        mono.setStyleHint(QFont.Monospace)
        self.editor.setFont(mono)
        layout.addWidget(self.editor, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_preamble(self) -> str:
        txt = self.editor.toPlainText().strip("\n")
        if not txt.strip():
            return ""
        # Ensure it looks like YAML front-matter (starts and ends with ---)
        lines = txt.split("\n")
        if not (lines[0].strip() == "---" and lines[-1].strip() == "---"):
            # If user omitted delimiters, wrap it
            wrapped = ["---"] + lines + ["---"]
            txt = "\n".join(wrapped)
        return txt.strip("\n") + "\n\n"


class ExportDialog(QDialog):
    def __init__(self, parent: QWidget, default_out: str, allow_local: bool):
        super().__init__(parent)
        self.setWindowTitle("Generate slides via Marp CLI")
        self.setModal(True)
        self.resize(700, 180)

        layout = QVBoxLayout(self)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Output file:"))
        self.path_edit = QLineEdit(default_out)
        row1.addWidget(self.path_edit, 1)
        browse = QPushButton("Browse…")
        browse.clicked.connect(self._browse)
        row1.addWidget(browse)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItem("HTML (.html)", "html")
        self.format_combo.addItem("PDF (.pdf)", "pdf")
        self.format_combo.addItem("PowerPoint (.pptx)", "pptx")
        row2.addWidget(self.format_combo)
        self.allow_local_chk = QCheckBox("Allow local files (images, etc.)")
        self.allow_local_chk.setChecked(bool(allow_local))
        row2.addWidget(self.allow_local_chk, 1)
        layout.addLayout(row2)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.format_combo.currentIndexChanged.connect(self._sync_ext)

    def _browse(self):
        fmt = self.format_combo.currentData()
        filt = {
            "html": "HTML (*.html)",
            "pdf": "PDF (*.pdf)",
            "pptx": "PowerPoint (*.pptx)",
        }[fmt]
        p, _ = QFileDialog.getSaveFileName(self, "Select output file", self.path_edit.text(), filt)
        if p:
            self.path_edit.setText(p)

    def _sync_ext(self):
        fmt = self.format_combo.currentData()
        p = Path(self.path_edit.text().strip() or "output")
        ext = {"html": ".html", "pdf": ".pdf", "pptx": ".pptx"}[fmt]
        if p.suffix.lower() != ext:
            p = p.with_suffix(ext)
            self.path_edit.setText(str(p))

    def get_values(self) -> Tuple[Path, str, bool]:
        out_path = Path(self.path_edit.text().strip())
        fmt = str(self.format_combo.currentData())
        allow_local = bool(self.allow_local_chk.isChecked())
        return out_path, fmt, allow_local


class InsertImageDialog(QDialog):
    def __init__(self, parent: QWidget, base_path: Optional[Path] = None):
        super().__init__(parent)
        self.setWindowTitle("Insert Picture")
        self.setModal(True)
        self.resize(500, 300)
        self.base_path = base_path

        layout = QVBoxLayout(self)

        # File selection
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Image Path:"))
        self.path_edit = QLineEdit()
        row1.addWidget(self.path_edit, 1)
        browse = QPushButton("Browse…")
        browse.clicked.connect(self._browse)
        row1.addWidget(browse)
        layout.addLayout(row1)

        # Position
        grp_pos = QWidget()
        l_pos = QHBoxLayout(grp_pos)
        l_pos.setContentsMargins(0, 0, 0, 0)
        l_pos.addWidget(QLabel("X (Left):"))
        self.x_edit = QLineEdit("100px")
        l_pos.addWidget(self.x_edit)
        l_pos.addWidget(QLabel("Y (Top):"))
        self.y_edit = QLineEdit("100px")
        l_pos.addWidget(self.y_edit)
        layout.addWidget(QLabel("Position (absolute):"))
        layout.addWidget(grp_pos)

        # Size
        grp_size = QWidget()
        l_size = QHBoxLayout(grp_size)
        l_size.setContentsMargins(0, 0, 0, 0)
        l_size.addWidget(QLabel("Width:"))
        self.w_edit = QLineEdit("300px")
        l_size.addWidget(self.w_edit)
        l_size.addWidget(QLabel("Height:"))
        self.h_edit = QLineEdit("auto")
        l_size.addWidget(self.h_edit)
        layout.addWidget(QLabel("Size:"))
        layout.addWidget(grp_size)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _browse(self):
        start_dir = str(self.base_path.parent) if self.base_path else ""
        p, _ = QFileDialog.getOpenFileName(self, "Select Image", start_dir, "Images (*.png *.jpg *.jpeg *.svg *.gif);;All files (*)")
        if p:
            # Try to make relative if possible
            if self.base_path:
                try:
                    rel = os.path.relpath(p, self.base_path.parent)
                    p = rel
                except ValueError:
                    pass
            self.path_edit.setText(p)

    def get_html_tag(self) -> str:
        src = self.path_edit.text().strip()
        x = self.x_edit.text().strip()
        y = self.y_edit.text().strip()
        w = self.w_edit.text().strip()
        h = self.h_edit.text().strip()

        style = "position: absolute;"
        if x: style += f" left: {x};"
        if y: style += f" top: {y};"
        if w: style += f" width: {w};"
        if h and h != "auto": style += f" height: {h};"

        return f'<img src="{src}" style="{style}" />'


class MainWindow(QMainWindow):
    def __init__(self, initial_file: Optional[Path] = None):
        super().__init__()
        self.setWindowTitle("Marpit Slide Editor")
        self.resize(1400, 850)

        self.marp_cmd = find_marp_cli_command()

        self.deck = DeckState.new_default()
        self._current_slide_idx = 0
        self._updating_editor = False

        self._tmp_dir = tempfile.TemporaryDirectory(prefix="marpit_editor_")
        self._preview_md_path = Path(self._tmp_dir.name) / "preview.md"
        self._preview_html_path = Path(self._tmp_dir.name) / "preview.html"

        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self.render_preview)

        self._preview_in_flight = False
        self._last_preview_hash = ""

        self.allow_local_files = True

        self._build_ui()

        # Set default splitter sizes (approx 2/3 for editor, 1/3 for preview)
        # Handle initial file load
        if initial_file and initial_file.exists():
            self.load_from_path(initial_file)

        self._refresh_slide_list()
        self._load_slide_into_editor(0)
        self._update_window_title()
        self._schedule_preview()

        if SPELLCHECK_AVAILABLE:
            self.highlighter = SpellCheckHighlighter(self.editor.document())

    # ---------------- UI ----------------
    def _build_ui(self):
        # Main toolbars
        file_tb = QToolBar("File")
        file_tb.setIconSize(QSize(16, 16))
        self.addToolBar(file_tb)

        act_new = QAction("New deck", self)
        act_new.setShortcut(QKeySequence.New)
        act_new.triggered.connect(self.new_deck)
        file_tb.addAction(act_new)

        act_open = QAction("Open…", self)
        act_open.setShortcut(QKeySequence.Open)
        act_open.triggered.connect(self.open_deck)
        file_tb.addAction(act_open)

        act_save = QAction("Save", self)
        act_save.setShortcut(QKeySequence.Save)
        act_save.triggered.connect(self.save_deck)
        file_tb.addAction(act_save)

        act_save_as = QAction("Save as…", self)
        act_save_as.setShortcut(QKeySequence.SaveAs)
        act_save_as.triggered.connect(self.save_deck_as)
        file_tb.addAction(act_save_as)

        file_tb.addSeparator()

        act_directives = QAction("Deck directives…", self)
        act_directives.triggered.connect(self.edit_deck_directives)
        file_tb.addAction(act_directives)

        file_tb.addSeparator()

        self.act_allow_local = QAction("Allow local files", self)
        self.act_allow_local.setCheckable(True)
        self.act_allow_local.setChecked(True)
        self.act_allow_local.triggered.connect(self._toggle_allow_local_files)
        file_tb.addAction(self.act_allow_local)

        file_tb.addSeparator()

        act_export = QAction("Generate slides…", self)
        act_export.triggered.connect(self.export_deck)
        file_tb.addAction(act_export)

        # Slide actions toolbar
        slide_tb = QToolBar("Slides")
        self.addToolBar(slide_tb)

        act_add = QAction("Add slide after", self)
        act_add.setShortcut(QKeySequence("Ctrl+Shift+N"))
        act_add.triggered.connect(self.add_slide_after_current)
        slide_tb.addAction(act_add)

        act_del = QAction("Delete slide", self)
        act_del.setShortcut(QKeySequence("Ctrl+Shift+Del"))
        act_del.triggered.connect(self.delete_current_slide)
        slide_tb.addAction(act_del)

        # Central layout: left half (deck overview + editor) and right half (preview)
        root_split = QSplitter(Qt.Horizontal)
        left_split = QSplitter(Qt.Horizontal)

        # Slide list (overview)
        self.slide_list = QListWidget()
        self.slide_list.setWordWrap(True)
        self.slide_list.setSelectionMode(QListWidget.SingleSelection)
        self.slide_list.currentRowChanged.connect(self._on_slide_selected)

        mono = QFont("Courier New")
        mono.setStyleHint(QFont.Monospace)
        self.slide_list.setFont(mono)
        self.slide_list.setStyleSheet(
            "QListWidget::item {"
            "  border: 1px solid #8a8a8a;"
            "  margin: 6px;"
            "  padding: 6px;"
            "  border-radius: 4px;"
            "}"
        )

        left_split.addWidget(self.slide_list)

        # Editor panel with formatting toolbar above
        editor_panel = QWidget()
        editor_layout = QVBoxLayout(editor_panel)
        editor_layout.setContentsMargins(0, 0, 0, 0)
        editor_layout.setSpacing(4)

        fmt_tb = QToolBar("Formatting")
        fmt_tb.setIconSize(QSize(16, 16))
        editor_layout.addWidget(fmt_tb)

        # Basic line styles (CommonMark)
        fmt_tb.addAction(self._make_action("Title (#)", lambda: self.set_line_heading(1)))
        fmt_tb.addAction(self._make_action("Subtitle (##)", lambda: self.set_line_heading(2)))
        fmt_tb.addAction(self._make_action("Heading (###)", lambda: self.set_line_heading(3)))
        fmt_tb.addSeparator()
        fmt_tb.addAction(self._make_action("Bullet (-)", lambda: self.set_line_list("-")))
        fmt_tb.addAction(self._make_action("Fragment bullet (*)", lambda: self.set_line_list("*")))
        fmt_tb.addAction(self._make_action("Numbered (1.)", self.set_line_numbered))
        fmt_tb.addSeparator()
        fmt_tb.addAction(self._make_action("Quote (>)", self.set_line_quote))
        fmt_tb.addAction(self._make_action("Code block", self.insert_code_block))

        # Image / background menu
        img_btn = QToolButton()
        img_btn.setText("Images ▾")
        img_btn.setPopupMode(QToolButton.InstantPopup)
        img_menu = QMenu(img_btn)
        img_menu.addAction("Inline image", lambda: self.insert_template("![](path-or-url)\n"))
        img_menu.addAction("Inline image (width)", lambda: self.insert_template("![width:200px](path-or-url)\n"))
        img_menu.addAction("Background image", lambda: self.insert_template("![bg](path-or-url)\n"))
        img_menu.addAction("Background (contain)", lambda: self.insert_template("![bg contain](path-or-url)\n"))
        img_menu.addAction("Background (cover)", lambda: self.insert_template("![bg cover](path-or-url)\n"))
        img_menu.addAction("Split background (left)", lambda: self.insert_template("![bg left](path-or-url)\n\n"))
        img_menu.addAction("Split background (right)", lambda: self.insert_template("![bg right](path-or-url)\n\n"))
        img_menu.addAction("Split background (left:33%)", lambda: self.insert_template("![bg left:33%](path-or-url)\n\n"))
        img_menu.addSeparator()
        img_menu.addAction("Insert Picture (absolute)…", self.insert_picture_dialog)
        img_btn.setMenu(img_menu)
        fmt_tb.addSeparator()
        fmt_tb.addWidget(img_btn)

        # Directives menu (Marpit)
        dir_btn = QToolButton()
        dir_btn.setText("Directives ▾")
        dir_btn.setPopupMode(QToolButton.InstantPopup)
        dir_menu = QMenu(dir_btn)
        dir_menu.addAction("Slide class: lead", lambda: self.insert_template("<!-- _class: lead -->\n"))
        dir_menu.addAction("Slide class: invert", lambda: self.insert_template("<!-- _class: invert -->\n"))
        dir_menu.addSeparator()
        dir_menu.addAction("Background color", lambda: self.insert_template("<!-- _backgroundColor: #ffffff -->\n"))
        dir_menu.addAction("Text color", lambda: self.insert_template("<!-- _color: #000000 -->\n"))
        dir_menu.addSeparator()
        dir_menu.addAction("Paginate: true", lambda: self.insert_template("<!-- paginate: true -->\n"))
        dir_menu.addAction("Paginate: hold (show, don't increment)", lambda: self.insert_template("<!-- _paginate: hold -->\n"))
        dir_menu.addAction("Paginate: skip (hide, don't increment)", lambda: self.insert_template("<!-- _paginate: skip -->\n"))
        dir_menu.addSeparator()
        dir_menu.addAction("Header", lambda: self.insert_template("<!-- header: \"\" -->\n"))
        dir_menu.addAction("Footer", lambda: self.insert_template("<!-- footer: \"\" -->\n"))
        dir_btn.setMenu(dir_menu)
        fmt_tb.addWidget(dir_btn)

        fmt_tb.addSeparator()
        fmt_tb.addAction(self._make_action("Presenter note", self.insert_presenter_note))

        self.editor = QPlainTextEdit()
        self.editor.setFont(mono)
        self.editor.textChanged.connect(self._on_editor_text_changed)
        editor_layout.addWidget(self.editor, 1)

        left_split.addWidget(editor_panel)
        left_split.setStretchFactor(0, 1)
        left_split.setStretchFactor(1, 2)

        # Preview
        preview_panel = QWidget()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        self.preview_label = QLabel()
        self.preview_label.setText("Preview")
        self.preview_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        preview_layout.addWidget(self.preview_label)

        if WEBENGINE_AVAILABLE:
            self.preview = QWebEngineView()
            preview_layout.addWidget(self.preview, 1)
        else:
            self.preview = QPlainTextEdit()
            self.preview.setReadOnly(True)
            self.preview.setPlainText(
                "Qt WebEngine is not available.\n\n"
                "Install a PySide6 build that includes QtWebEngine (pip install PySide6).\n"
                "Then restart the application."
            )
            preview_layout.addWidget(self.preview, 1)

        root_split.addWidget(left_split)
        root_split.addWidget(preview_panel)
        # Default ~2/3 for editor+list, ~1/3 for preview
        root_split.setStretchFactor(0, 2)
        root_split.setStretchFactor(1, 1)

        self.setCentralWidget(root_split)

        # Status bar
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._update_status()

    def _make_action(self, text: str, slot):
        act = QAction(text, self)
        act.triggered.connect(slot)
        return act

    # ---------------- State helpers ----------------
    def _update_window_title(self):
        name = self.deck.file_path.name if self.deck.file_path else "Untitled"
        dirty = " *" if self.deck.dirty else ""
        self.setWindowTitle(f"Marpit Slide Editor — {name}{dirty}")

    def _update_status(self, msg: str = ""):
        parts = []
        parts.append(f"Slides: {len(self.deck.slides)}")
        if self.deck.file_path:
            parts.append(str(self.deck.file_path))
        else:
            parts.append("Unsaved deck")

        if self.marp_cmd:
            parts.append(f"Marp CLI: {' '.join(self.marp_cmd)}")
        else:
            parts.append("Marp CLI: NOT FOUND (set MARP_CLI or install @marp-team/marp-cli)")

        if msg:
            parts.append(msg)

        self.statusBar().showMessage(" | ".join(parts))

    def _toggle_allow_local_files(self):
        self.allow_local_files = bool(self.act_allow_local.isChecked())
        self._schedule_preview()

    # ---------------- Slide list / selection ----------------
    def _refresh_slide_list(self):
        self.slide_list.blockSignals(True)
        self.slide_list.clear()

        for i, slide in enumerate(self.deck.slides):
            display = f"Slide {i+1}\n" + ("-" * 24) + "\n"
            display += (slide or "").strip("\n")
            if not display.strip():
                display = f"Slide {i+1}\n" + ("-" * 24) + "\n(empty)"
            item = QListWidgetItem(display)
            item.setData(Qt.UserRole, i)
            # Make items taller if slide has many lines (bounded)
            line_count = max(4, display.count("\n") + 1)
            item.setSizeHint(QSize(220, 18 * line_count))
            self.slide_list.addItem(item)

        self.slide_list.setCurrentRow(self._current_slide_idx)
        self.slide_list.blockSignals(False)

    def _on_slide_selected(self, row: int):
        if row < 0 or row >= len(self.deck.slides):
            return
        if row == self._current_slide_idx:
            return
        if not self._maybe_commit_current_editor():
            # Revert selection if committing failed (rare)
            self.slide_list.setCurrentRow(self._current_slide_idx)
            return

        self._current_slide_idx = row
        self._load_slide_into_editor(row)
        self._schedule_preview()

    def _load_slide_into_editor(self, idx: int):
        self._updating_editor = True
        try:
            self.editor.setPlainText(self.deck.slides[idx] or "")
        finally:
            self._updating_editor = False

    def _maybe_commit_current_editor(self) -> bool:
        # Editor is the source of truth for current slide; ensure deck is updated.
        if self._updating_editor:
            return True
        txt = self.editor.toPlainText().strip("\n")
        if self.deck.slides[self._current_slide_idx] != txt:
            self.deck.slides[self._current_slide_idx] = txt
            self.deck.dirty = True
            self._update_window_title()
            self._refresh_slide_list()
        return True

    def _on_editor_text_changed(self):
        if self._updating_editor:
            return
        # Commit to deck model and refresh list item lazily (cheap enough for now)
        txt = self.editor.toPlainText().strip("\n")
        if self.deck.slides[self._current_slide_idx] != txt:
            self.deck.slides[self._current_slide_idx] = txt
            self.deck.dirty = True
            self._update_window_title()
            # Update only the changed item for responsiveness
            item = self.slide_list.item(self._current_slide_idx)
            if item:
                display = f"Slide {self._current_slide_idx+1}\n" + ("-" * 24) + "\n" + txt
                if not txt.strip():
                    display = f"Slide {self._current_slide_idx+1}\n" + ("-" * 24) + "\n(empty)"
                item.setText(display)
                line_count = max(4, display.count("\n") + 1)
                item.setSizeHint(QSize(220, 18 * line_count))
        self._schedule_preview()

    # ---------------- Deck actions ----------------
    def new_deck(self):
        if not self._confirm_discard_if_dirty():
            return
        self.deck = DeckState.new_default()
        self._current_slide_idx = 0
        self._refresh_slide_list()
        self._load_slide_into_editor(0)
        self._update_window_title()
        self._update_status("New deck created.")
        self._schedule_preview()

    def open_deck(self):
        if not self._confirm_discard_if_dirty():
            return
        path_str, _ = QFileDialog.getOpenFileName(self, "Open Marp/Marpit Markdown", "", "Markdown (*.md *.markdown *.mdown);;All files (*)")
        if not path_str:
            return
        self.load_from_path(Path(path_str))

    def load_from_path(self, p: Path):
        try:
            text = p.read_text(encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Open failed", f"Could not read file:\n{p}\n\n{e}")
            return
        preamble, slides = parse_marp_markdown(text)
        self.deck = DeckState(preamble=preamble, slides=slides, file_path=p, dirty=False)
        self._current_slide_idx = 0
        self._refresh_slide_list()
        self._load_slide_into_editor(0)
        self._update_window_title()
        self._update_status("Deck loaded.")
        self._schedule_preview()

    def save_deck(self):
        if self.deck.file_path is None:
            return self.save_deck_as()
        return self._save_to_path(self.deck.file_path)

    def save_deck_as(self):
        path_str, _ = QFileDialog.getSaveFileName(self, "Save Marp/Marpit Markdown", "", "Markdown (*.md *.markdown);;All files (*)")
        if not path_str:
            return False
        p = Path(path_str)
        if p.suffix.lower() not in (".md", ".markdown", ".mdown"):
            p = p.with_suffix(".md")
        ok = self._save_to_path(p)
        if ok:
            self.deck.file_path = p
            self._update_window_title()
            self._update_status("Saved.")
            self._schedule_preview()
        return ok

    def _save_to_path(self, path: Path) -> bool:
        self._maybe_commit_current_editor()
        text = serialize_marp_markdown(self.deck.preamble, self.deck.slides)
        try:
            path.write_text(text, encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Save failed", f"Could not write file:\n{path}\n\n{e}")
            return False
        self.deck.dirty = False
        self._update_window_title()
        self._update_status("Saved.")
        return True

    def add_slide_after_current(self):
        self._maybe_commit_current_editor()
        insert_at = self._current_slide_idx + 1
        self.deck.slides.insert(insert_at, "# New slide\n")
        self.deck.dirty = True
        self._current_slide_idx = insert_at
        self._refresh_slide_list()
        self._load_slide_into_editor(insert_at)
        self._update_window_title()
        self._update_status("Slide added.")
        self._schedule_preview()

    def delete_current_slide(self):
        if len(self.deck.slides) <= 1:
            # Keep at least one slide
            self.deck.slides[0] = ""
            self._load_slide_into_editor(0)
            self.deck.dirty = True
            self._refresh_slide_list()
            self._update_window_title()
            self._schedule_preview()
            return

        self._maybe_commit_current_editor()
        del self.deck.slides[self._current_slide_idx]
        self.deck.dirty = True
        self._current_slide_idx = max(0, self._current_slide_idx - 1)
        self._refresh_slide_list()
        self._load_slide_into_editor(self._current_slide_idx)
        self._update_window_title()
        self._update_status("Slide deleted.")
        self._schedule_preview()

    def edit_deck_directives(self):
        dlg = DeckDirectivesDialog(self, self.deck.preamble)
        if dlg.exec() == QDialog.Accepted:
            self.deck.preamble = dlg.get_preamble()
            self.deck.dirty = True
            self._update_window_title()
            self._update_status("Deck directives updated.")
            self._schedule_preview()

    # ---------------- Formatting actions ----------------
    def _transform_current_line(self, new_line: str):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        cursor.insertText(new_line)
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)

    def set_line_heading(self, level: int):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText()
        content = md_line_transform_remove_leading_syntax(line)
        prefix = "#" * max(1, min(6, level)) + " "
        self._transform_current_line(prefix + content)

    def set_line_list(self, marker: str):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText()
        content = md_line_transform_remove_leading_syntax(line)
        prefix = f"{marker} "
        self._transform_current_line(prefix + content)

    def set_line_numbered(self):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText()
        content = md_line_transform_remove_leading_syntax(line)
        self._transform_current_line("1. " + content)

    def set_line_quote(self):
        cursor = self.editor.textCursor()
        cursor.select(QTextCursor.LineUnderCursor)
        line = cursor.selectedText()
        content = md_line_transform_remove_leading_syntax(line)
        self._transform_current_line("> " + content)

    def insert_template(self, snippet: str):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.insertText(snippet)
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)

    def insert_code_block(self):
        snippet = "```text\n\n```\n"
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.insertText(snippet)
        # Move cursor to the empty line inside the code block
        cursor.movePosition(QTextCursor.Up)
        cursor.movePosition(QTextCursor.StartOfLine)
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)

    def insert_presenter_note(self):
        snippet = "<!--\nPresenter notes...\n-->\n"
        self.insert_template(snippet)

    def insert_picture_dialog(self):
        if not self.deck.file_path:
            QMessageBox.warning(self, "Save first", "Please save the deck first so we can resolve relative paths.")
            return

        dlg = InsertImageDialog(self, base_path=self.deck.file_path)
        if dlg.exec():
            html_tag = dlg.get_html_tag()
            self.insert_template(html_tag)


    # ---------------- Preview rendering ----------------
    def _schedule_preview(self):
        # Debounce frequent edits
        self._preview_timer.start(350)

    def render_preview(self):
        if not WEBENGINE_AVAILABLE:
            return

        if not self.marp_cmd:
            self._set_preview_error(
                "Marp CLI was not found.\n\n"
                "Install it:\n"
                "  npm install -g @marp-team/marp-cli\n\n"
                "Or set MARP_CLI env var to the marp command.\n"
            )
            return

        self._maybe_commit_current_editor()

        # Build preview markdown (preamble + current slide)
        slide_md = (self.deck.slides[self._current_slide_idx] or "").strip("\n")
        preview_md = (self.deck.preamble or "") + (slide_md + "\n")

        h = hashlib.sha256(preview_md.encode("utf-8")).hexdigest()
        if h == self._last_preview_hash and self._preview_html_path.exists():
            # Nothing changed
            return
        self._last_preview_hash = h

        try:
            self._preview_md_path.write_text(preview_md, encoding="utf-8")
        except Exception as e:
            self._set_preview_error(f"Failed writing preview markdown:\n{e}")
            return

        # Run marp CLI to generate HTML preview
        # marp preview.md -o preview.html
        args = []
        args.extend(self.marp_cmd[1:])
        args.append(str(self._preview_md_path))
        args.extend(["-o", str(self._preview_html_path)])
        if self.allow_local_files:
            args.append("--allow-local-files")

        program = self.marp_cmd[0]

        # Use a simple subprocess to keep this single-file; rendering one slide is fast.
        # If it fails, show stderr.
        import subprocess

        try:
            proc = subprocess.run(
                [program] + args,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=str(self.deck.file_path.parent) if self.deck.file_path else None,
                timeout=10,
            )
        except Exception as e:
            self._set_preview_error(f"Marp CLI failed to run:\n{e}")
            return

        if proc.returncode != 0:
            self._set_preview_error("Marp CLI error:\n\n" + (proc.stderr.strip() or proc.stdout.strip() or "(no output)"))
            return

        if not self._preview_html_path.exists():
            self._set_preview_error("Preview HTML was not generated (unexpected).")
            return

        self.preview.load(QUrl.fromLocalFile(str(self._preview_html_path.resolve())))
        self._update_status("Preview updated.")

    def _set_preview_error(self, message: str):
        self._update_status("Preview error.")
        if WEBENGINE_AVAILABLE:
            html = (
                "<html><body style='font-family:sans-serif;padding:12px;'>"
                "<h3>Preview unavailable</h3>"
                f"<pre style='white-space:pre-wrap'>{self._escape_html(message)}</pre>"
                "</body></html>"
            )
            self.preview.setHtml(html, QUrl.fromLocalFile(str(Path(self._tmp_dir.name).resolve())))
        else:
            self.preview.setPlainText(message)

    @staticmethod
    def _escape_html(s: str) -> str:
        return (
            s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
        )

    # ---------------- Export ----------------
    def export_deck(self):
        if not self.marp_cmd:
            QMessageBox.warning(
                self,
                "Marp CLI not found",
                "Marp CLI was not found.\n\nInstall it:\n  npm install -g @marp-team/marp-cli\n\nOr set MARP_CLI env var.",
            )
            return

        # For best results, export from a real saved file so relative assets resolve.
        temp_input: Optional[Path] = None
        if self.deck.file_path and not self.deck.dirty and self.deck.file_path.exists():
            input_path = self.deck.file_path
        else:
            # Write a temporary input file
            try:
                tmp = Path(self._tmp_dir.name) / "export.md"
                tmp.write_text(serialize_marp_markdown(self.deck.preamble, self.deck.slides), encoding="utf-8")
                input_path = tmp
                temp_input = tmp
            except Exception as e:
                QMessageBox.critical(self, "Export failed", f"Could not create temporary input deck:\n{e}")
                return

        default_out = "deck.html"
        if self.deck.file_path:
            default_out = str(self.deck.file_path.with_suffix(".html"))

        dlg = ExportDialog(self, default_out=default_out, allow_local=self.allow_local_files)
        if dlg.exec() != QDialog.Accepted:
            return
        out_path, fmt, allow_local = dlg.get_values()

        # Build command arguments
        args = []
        args.extend(self.marp_cmd[1:])
        if fmt == "pdf":
            args.append("--pdf")
        elif fmt == "pptx":
            args.append("--pptx")
        # HTML is default
        if allow_local:
            args.append("--allow-local-files")

        args.append(str(input_path))
        args.extend(["-o", str(out_path)])

        program = self.marp_cmd[0]
        import subprocess

        self._update_status("Export started...")

        try:
            proc = subprocess.run(
                [program] + args,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=str(input_path.parent),
                timeout=120,
            )
        except Exception as e:
            QMessageBox.critical(self, "Export failed", f"Marp CLI failed to run:\n{e}")
            self._update_status("Export failed.")
            return

        if proc.returncode != 0:
            QMessageBox.critical(
                self,
                "Export failed",
                "Marp CLI returned a non-zero exit code.\n\n"
                + (proc.stderr.strip() or proc.stdout.strip() or "(no output)"),
            )
            self._update_status("Export failed.")
            return

        QMessageBox.information(self, "Export complete", f"Generated:\n{out_path}")
        self._update_status(f"Exported {fmt.upper()} to {out_path}")

    # ---------------- Close handling ----------------
    def _confirm_discard_if_dirty(self) -> bool:
        if not self.deck.dirty:
            return True
        resp = QMessageBox.question(
            self,
            "Unsaved changes",
            "You have unsaved changes. Do you want to save before continuing?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
            QMessageBox.Save,
        )
        if resp == QMessageBox.Save:
            return bool(self.save_deck())
        if resp == QMessageBox.Discard:
            return True
        return False

    def closeEvent(self, event):  # type: ignore
        if not self._confirm_discard_if_dirty():
            event.ignore()
            return
        try:
            self._tmp_dir.cleanup()
        except Exception:
            pass
        event.accept()


def main():
    app = QApplication(sys.argv)

    initial_file = None
    if len(sys.argv) > 1:
        initial_file = Path(sys.argv[1])

    w = MainWindow(initial_file=initial_file)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
