
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
import shutil
import sys
import tempfile
import json
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from PySide6.QtCore import Qt, QTimer, QUrl, QSize
from PySide6.QtGui import QAction, QFont, QKeySequence, QTextCursor, QSyntaxHighlighter, QTextCharFormat, QColor, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QColorDialog,
    QDialog,
    QDialogButtonBox,
    QDockWidget,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QSpinBox,
    QTabWidget,
    QMessageBox,
    QPushButton,
    QSlider,
    QSplitter,
    QStatusBar,
    QStyle,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QPlainTextEdit,
    QTextEdit,
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


class TemplateManager:
    """Manages loading templates from the templates/ directory."""
    def __init__(self, root_dir: Path):
        self.templates_dir = root_dir / "templates"
    
    def get_available_templates(self) -> List[Path]:
        if not self.templates_dir.exists():
            return []
        return sorted(list(self.templates_dir.glob("*.md")))

    def load_template(self, path: Path) -> Tuple[str, List[str]]:
        try:
            text = path.read_text(encoding="utf-8")
            return parse_marp_markdown(text)
        except Exception as e:
            print(f"Error loading template {path}: {e}")
            return "", [""]



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
    local_bin = (Path(__file__).parent / "bin" / "marp").resolve()
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
        self.format_combo.setCurrentIndex(1) # Select PDF by default
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
        
        # Sync extension on init
        self._sync_ext()

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
        # We start in the CWD (which should mean the deck's directory if loaded/saved)
        # But we can also fallback to base_path if provided.
        start = str(self.base_path.parent) if (self.base_path and self.base_path.parent.exists()) else os.getcwd()

        p, _ = QFileDialog.getOpenFileName(self, "Select Image", start, "Images (*.png *.jpg *.jpeg *.svg *.gif);;All files (*)")
        if p:
            # Always try to make it relative to CWD (which is where the deck is)
            try:
                rel = os.path.relpath(p, os.getcwd())
                p = rel
            except ValueError:
                # If on different drive or fails, keep absolute
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


class GlobalBackgroundDialog(QDialog):
    def __init__(self, parent: QWidget, base_path: Optional[Path] = None):
        super().__init__(parent)
        self.setWindowTitle("Set Global Background Image")
        self.setModal(True)
        self.resize(500, 200)
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

        # Size mode
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Size Mode:"))
        self.size_combo = QComboBox()
        self.size_combo.addItem("Cover (fill slide)", "cover")
        self.size_combo.addItem("Contain (fit inside)", "contain")
        self.size_combo.addItem("Original size", "auto")
        self.size_combo.addItem("Stretch (100% 100%)", "100% 100%")
        row2.addWidget(self.size_combo, 1)

        # Position
        self.pos_combo = QComboBox()
        positions = [
            "center center", "top left", "top center", "top right",
            "bottom left", "bottom center", "bottom right",
            "center left", "center right"
        ]
        for p in positions:
            self.pos_combo.addItem(p, p)

        row2.addWidget(QLabel("Position:"))
        row2.addWidget(self.pos_combo, 1)
        layout.addLayout(row2)
        # Opacity
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Opacity:"))
        self.opacity_spin = QDoubleSpinBox()
        self.opacity_spin.setRange(0.0, 1.0)
        self.opacity_spin.setSingleStep(0.1)
        self.opacity_spin.setValue(1.0)
        row3.addWidget(self.opacity_spin)

        # Slider for convenience
        self.opacity_slider = QSlider(Qt.Horizontal)
        self.opacity_slider.setRange(0, 100)
        self.opacity_slider.setValue(100)
        self.opacity_slider.valueChanged.connect(self._sync_spin_from_slider)
        self.opacity_spin.valueChanged.connect(self._sync_slider_from_spin)
        row3.addWidget(self.opacity_slider, 1)

        layout.addLayout(row3)

        info = QLabel("This will add/update a 'style' block in YAML using 'section::before' for opacity support.")
        info.setStyleSheet("color: gray; font-style: italic;")
        info.setWordWrap(True)
        layout.addWidget(info)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _sync_spin_from_slider(self, val):
        self.opacity_spin.blockSignals(True)
        self.opacity_spin.setValue(val / 100.0)
        self.opacity_spin.blockSignals(False)

    def _sync_slider_from_spin(self, val):
        self.opacity_slider.blockSignals(True)
        self.opacity_slider.setValue(int(val * 100))
        self.opacity_slider.blockSignals(False)

    def _browse(self):
        start = str(self.base_path.parent) if (self.base_path and self.base_path.parent.exists()) else os.getcwd()
        p, _ = QFileDialog.getOpenFileName(self, "Select Image", start, "Images (*.png *.jpg *.jpeg *.svg *.gif);;All files (*)")
        if p:
            try:
                rel = os.path.relpath(p, os.getcwd())
                p = rel
            except ValueError:
                pass
            self.path_edit.setText(p)

    def load_settings(self, src: str, mode: str, opacity: float, position: str = "center center"):
        self.path_edit.setText(src)

        idx = self.size_combo.findData(mode)
        if idx >= 0: self.size_combo.setCurrentIndex(idx)

        idx_pos = self.pos_combo.findData(position)
        if idx_pos >= 0: self.pos_combo.setCurrentIndex(idx_pos)

        self.opacity_spin.setValue(opacity)

    def get_css_content(self) -> str:
        src = self.path_edit.text().strip()
        if not src:
            return ""

        mode = self.size_combo.currentData()
        pos = self.pos_combo.currentData()
        opacity = self.opacity_spin.value()

        # Minimal escape:
        src_escaped = src.replace('"', '%22').replace("'", '%27')

        # Use linear-gradient overlay to simulate opacity (Fading to white)
        # This is more robust than ::before + z-index which often fails in previews.
        # Opacity 1.0 -> Alpha 0.0 (Clear)
        # Opacity 0.0 -> Alpha 1.0 (Solid White)
        alpha = 1.0 - opacity

        # We also enforce white background on body/section to ensure consistency
        css = (
            "body, .marpit {\n"
            "  background-color: white !important;\n"
            "}\n"
            "section {\n"
            "  background-color: white !important;\n"
            f"  background-image: linear-gradient(rgba(255,255,255,{alpha:.2f}), rgba(255,255,255,{alpha:.2f})), url('{src_escaped}');\n"
            "  background-repeat: no-repeat;\n"
            f"  background-position: {pos};\n"
            f"  background-size: {mode};\n"
            "}"
        )
        return css

class GlobalHeaderFooterDialog(QDialog):
    def __init__(self, parent: QWidget, base_path: Optional[Path] = None):
        super().__init__(parent)
        self.setWindowTitle("Set Global Header/Footer")
        self.resize(600, 500)
        self.base_path = base_path

        layout = QVBoxLayout(self)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Header Tab
        self.header_ui = self._create_tab_ui("Header")
        self.tabs.addTab(self.header_ui['widget'], "Header")

        # Footer Tab
        self.footer_ui = self._create_tab_ui("Footer")
        self.tabs.addTab(self.footer_ui['widget'], "Footer")

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _create_tab_ui(self, label: str):
        widget = QWidget()
        lay = QVBoxLayout(widget)

        # Text Content (Markdown)
        lay.addWidget(QLabel(f"{label} Content (Inline Markdown only - no # headings):"))
        text_edit = QTextEdit()
        text_edit.setPlaceholderText(f"Enter {label.lower()} text here. You can use bold, italic, images, etc.")
        lay.addWidget(text_edit, 1)

        # Image Button
        btn_img = QPushButton("Insert Image...")
        btn_img.clicked.connect(lambda: self._insert_image(text_edit))
        lay.addWidget(btn_img)

        # Styling Group
        grp = QGroupBox("Styling & Layout")
        glay = QFormLayout(grp)

        # Height
        h_spin = QSpinBox()
        h_spin.setRange(0, 500)
        h_spin.setValue(100 if label == "Header" else 50)
        h_spin.setSuffix(" px")
        glay.addRow("Height:", h_spin)

        # Offset (gap from edge)
        off_spin = QSpinBox()
        off_spin.setRange(-100, 300)
        off_spin.setValue(0)
        off_spin.setSuffix(" px")
        glay.addRow("Edge Offset:", off_spin)

        # Alignment
        align_combo = QComboBox()
        # Data: CSS property for text-align or special for spread
        align_combo.addItem("Left", "left")
        align_combo.addItem("Center", "center")
        align_combo.addItem("Right", "right")
        align_combo.addItem("Spread (Justify)", "spread")
        align_combo.setCurrentIndex(0) # Default Left
        glay.addRow("Alignment:", align_combo)

        # Font Size
        fs_spin = QSpinBox()
        fs_spin.setRange(8, 100)
        fs_spin.setValue(18) # Default
        fs_spin.setSuffix(" px")
        glay.addRow("Font Size:", fs_spin)

        # Font Family
        font_combo = QComboBox()
        font_combo.addItem("Default", "")
        font_combo.addItem("Arial (Sans)", "Arial, Helvetica, sans-serif")
        font_combo.addItem("Times New Roman (Serif)", '"Times New Roman", Times, serif')
        font_combo.addItem("Courier New (Mono)", '"Courier New", Courier, monospace')
        font_combo.addItem("Georgia (Serif)", "Georgia, serif")
        font_combo.addItem("Verdana (Sans)", "Verdana, Geneva, sans-serif")
        glay.addRow("Font Family:", font_combo)

        # Text Color
        color_btn = QPushButton()
        color_btn.setText("Pick Color...")
        # Store color in a property
        color_btn.setProperty("selected_color", "")
        color_btn.clicked.connect(lambda: self._pick_color(color_btn))
        glay.addRow("Text Color:", color_btn)

        lay.addWidget(grp)

        return {
            'widget': widget,
            'text': text_edit,
            'height': h_spin,
            'offset': off_spin,
            'align': align_combo,
            'font_size': fs_spin,
            'font_family': font_combo,
            'color': color_btn
        }

    def _pick_color(self, btn: QPushButton):
        curr = btn.property("selected_color") or "#000000"
        c = QColorDialog.getColor(QColor(curr), self, "Select Text Color")
        if c.isValid():
            hex_c = c.name()
            btn.setProperty("selected_color", hex_c)
            btn.setStyleSheet(f"background-color: {hex_c}; color: {'white' if c.lightness() < 128 else 'black'}")
            btn.setText(hex_c)

    def _insert_image(self, editor: QTextEdit):
        start = str(self.base_path.parent) if (self.base_path and self.base_path.parent.exists()) else os.getcwd()
        p, _ = QFileDialog.getOpenFileName(self, "Select Image", start, "Images (*.png *.jpg *.jpeg *.svg *.gif);;All files (*)")
        if p:
            try:
                rel = os.path.relpath(p, os.getcwd())
                # Quote path if spaces
                if " " in rel:
                    rel = f"'{rel}'"
                # Insert standard Marp image syntax
                # Typically ![h:50](path)
                editor.insertPlainText(f"![h:50]({rel}) ")
            except ValueError:
                pass

    def get_settings(self):
        return {
            'header': self._get_tab_data(self.header_ui),
            'footer': self._get_tab_data(self.footer_ui)
        }

    def _get_tab_data(self, ui):
        return {
            'content': ui['text'].toPlainText(),
            'height': ui['height'].value(),
            'offset': ui['offset'].value(),
            'offset': ui['offset'].value(),
            'align': ui['align'].currentData(),
            'font_size': ui['font_size'].value(),
            'font_family': ui['font_family'].currentData(),
            'color': ui['color'].property("selected_color")
        }

    def load_settings(self, data: dict):
        if 'header' in data:
            self._set_tab_data(self.header_ui, data['header'])
        if 'footer' in data:
            self._set_tab_data(self.footer_ui, data['footer'])

    def _set_tab_data(self, ui, d):
        ui['text'].setPlainText(d.get('content', ''))
        ui['height'].setValue(int(d.get('height', 100)))
        ui['offset'].setValue(int(d.get('offset', 0)))

        al = d.get('align', 'left')
        idx = ui['align'].findData(al)
        if idx >= 0:
            ui['align'].setCurrentIndex(idx)

        ui['font_size'].setValue(int(d.get('font_size', 18)))

        # Font Family
        fam = d.get('font_family', '')
        idx = ui['font_family'].findData(fam)
        if idx >= 0: ui['font_family'].setCurrentIndex(idx)
        else: ui['font_family'].setCurrentIndex(0)

        # Color
        col = d.get('color', '')
        if col:
             ui['color'].setProperty("selected_color", col)
             ui['color'].setText(col)
             try:
                 c = QColor(col)
                 if c.isValid():
                     ui['color'].setStyleSheet(f"background-color: {col}; color: {'white' if c.lightness() < 128 else 'black'}")
             except: pass
        else:
             ui['color'].setProperty("selected_color", "")
             ui['color'].setText("Pick Color...")
             ui['color'].setStyleSheet("")



class PaginationDialog(QDialog):
    def __init__(self, parent: QWidget, base_path: Optional[Path] = None):
        super().__init__(parent)
        self.setWindowTitle("Pagination Settings")
        self.setModal(True)
        self.resize(400, 300)
        self.base_path = base_path

        layout = QVBoxLayout(self)

        # Enable/Disable
        self.chk_paginate = QCheckBox("Show Page Numbers")
        layout.addWidget(self.chk_paginate)

        grp = QGroupBox("Appearance")
        glay = QFormLayout(grp)

        # Font Size
        self.spin_size = QSpinBox()
        self.spin_size.setRange(8, 100)
        self.spin_size.setValue(18)
        self.spin_size.setSuffix(" px")
        glay.addRow("Font Size:", self.spin_size)

        # Font Family
        self.combo_font = QComboBox()
        self.combo_font.addItem("Default", "")
        self.combo_font.addItem("Arial (Sans)", "Arial, Helvetica, sans-serif")
        self.combo_font.addItem("Times New Roman (Serif)", '"Times New Roman", Times, serif')
        self.combo_font.addItem("Courier New (Mono)", '"Courier New", Courier, monospace')
        self.combo_font.addItem("Verdana (Sans)", "Verdana, Geneva, sans-serif")
        glay.addRow("Font Family:", self.combo_font)

        # Color
        self.btn_color = QPushButton("Pick Color...")
        self.btn_color.setProperty("selected_color", "")
        self.btn_color.clicked.connect(self._pick_color)
        glay.addRow("Color:", self.btn_color)

        # Position (CSS logic is section::after { right: X, bottom: Y } etc)
        self.combo_pos = QComboBox()
        self.combo_pos.addItem("Bottom Right", "bottom-right")
        self.combo_pos.addItem("Bottom Left", "bottom-left")
        self.combo_pos.addItem("Top Right", "top-right")
        self.combo_pos.addItem("Top Left", "top-left")
        glay.addRow("Position:", self.combo_pos)

        layout.addWidget(grp)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        # Disable appearance if unchecked
        self.chk_paginate.toggled.connect(grp.setEnabled)

    def _pick_color(self):
        curr = self.btn_color.property("selected_color") or "#000000"
        c = QColorDialog.getColor(QColor(curr), self, "Select Color")
        if c.isValid():
            hex_c = c.name()
            self.btn_color.setProperty("selected_color", hex_c)
            self.btn_color.setStyleSheet(f"background-color: {hex_c}; color: {'white' if c.lightness() < 128 else 'black'}")
            self.btn_color.setText(hex_c)

    def load_settings(self, enabled: bool, size: int, font: str, color: str, pos: str):
        self.chk_paginate.setChecked(enabled)
        self.spin_size.setValue(size)
        
        idx = self.combo_font.findData(font)
        if idx >= 0: self.combo_font.setCurrentIndex(idx)
        else: self.combo_font.setCurrentIndex(0)

        if color:
             self.btn_color.setProperty("selected_color", color)
             self.btn_color.setText(color)
             self.btn_color.setStyleSheet(f"background-color: {color}; color: auto")
        
        idx_pos = self.combo_pos.findData(pos)
        if idx_pos >= 0: self.combo_pos.setCurrentIndex(idx_pos)

    def get_settings(self):
        return {
            'enabled': self.chk_paginate.isChecked(),
            'size': self.spin_size.value(),
            'font': self.combo_font.currentData(),
            'color': self.btn_color.property("selected_color"),
            'pos': self.combo_pos.currentData()
        }

class PaginationDialog(QDialog):
    def __init__(self, parent: QWidget, base_path: Optional[Path] = None):
        super().__init__(parent)
        self.setWindowTitle("Pagination Settings")
        self.setModal(True)
        self.resize(400, 300)
        self.base_path = base_path

        layout = QVBoxLayout(self)

        # Enable/Disable
        self.chk_paginate = QCheckBox("Show Page Numbers")
        layout.addWidget(self.chk_paginate)

        grp = QGroupBox("Appearance")
        glay = QFormLayout(grp)

        # Font Size
        self.spin_size = QSpinBox()
        self.spin_size.setRange(8, 100)
        self.spin_size.setValue(18)
        self.spin_size.setSuffix(" px")
        glay.addRow("Font Size:", self.spin_size)

        # Font Family
        self.combo_font = QComboBox()
        self.combo_font.addItem("Default", "")
        self.combo_font.addItem("Arial (Sans)", "Arial, Helvetica, sans-serif")
        self.combo_font.addItem("Times New Roman (Serif)", '"Times New Roman", Times, serif')
        self.combo_font.addItem("Courier New (Mono)", '"Courier New", Courier, monospace')
        self.combo_font.addItem("Verdana (Sans)", "Verdana, Geneva, sans-serif")
        glay.addRow("Font Family:", self.combo_font)

        # Color
        self.btn_color = QPushButton("Pick Color...")
        self.btn_color.setProperty("selected_color", "")
        self.btn_color.clicked.connect(self._pick_color)
        glay.addRow("Color:", self.btn_color)

        # Position
        self.combo_pos = QComboBox()
        self.combo_pos.addItem("Bottom Right", "bottom-right")
        self.combo_pos.addItem("Bottom Left", "bottom-left")
        self.combo_pos.addItem("Top Right", "top-right")
        self.combo_pos.addItem("Top Left", "top-left")
        glay.addRow("Position:", self.combo_pos)

        layout.addWidget(grp)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.chk_paginate.toggled.connect(grp.setEnabled)

    def _pick_color(self):
        curr = self.btn_color.property("selected_color") or "#000000"
        c = QColorDialog.getColor(QColor(curr), self, "Select Color")
        if c.isValid():
            hex_c = c.name()
            self.btn_color.setProperty("selected_color", hex_c)
            self.btn_color.setStyleSheet(f"background-color: {hex_c}; color: {'white' if c.lightness() < 128 else 'black'}")
            self.btn_color.setText(hex_c)

    def load_settings(self, enabled: bool, size: int, font: str, color: str, pos: str):
        self.chk_paginate.setChecked(enabled)
        self.spin_size.setValue(size)
        
        idx = self.combo_font.findData(font)
        if idx >= 0: self.combo_font.setCurrentIndex(idx)
        else: self.combo_font.setCurrentIndex(0)

        if color:
             self.btn_color.setProperty("selected_color", color)
             self.btn_color.setText(color)
             self.btn_color.setStyleSheet(f"background-color: {color}; color: auto")
        
        idx_pos = self.combo_pos.findData(pos)
        if idx_pos >= 0: self.combo_pos.setCurrentIndex(idx_pos)

    def get_settings(self):
        return {
            'enabled': self.chk_paginate.isChecked(),
            'size': self.spin_size.value(),
            'font': self.combo_font.currentData(),
            'color': self.btn_color.property("selected_color"),
            'pos': self.combo_pos.currentData()
        }

# -----------------------------------------------------------------------------
# Slide Properties Widget (Dock)
# -----------------------------------------------------------------------------
class SlidePropertiesWidget(QDockWidget):
    def __init__(self, parent=None):
        super().__init__("Slide Properties", parent)
        self.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        
        container = QWidget()
        self.setWidget(container)
        layout = QVBoxLayout(container)
        
        # 1. Directives Group
        grp_dir = QGroupBox("Local Directives")
        form = QFormLayout(grp_dir)
        
        self.chk_paginate = QCheckBox("Show Page Number (paginate)")
        form.addRow(self.chk_paginate)
        
        self.chk_header = QCheckBox("Show Header (undo _header: \"\")")
        self.chk_header.setToolTip("If unchecked, adds <!-- _header: \"\" --> to hide global header.")
        form.addRow(self.chk_header)
        
        self.chk_footer = QCheckBox("Show Footer (undo _footer: \"\")")
        self.chk_footer.setToolTip("If unchecked, adds <!-- _footer: \"\" --> to hide global footer.")
        form.addRow(self.chk_footer)
        
        layout.addWidget(grp_dir)
        
        # 2. Key/Value Directives (Simplified)
        # For now, just a simplified text edit for specific storage if needed, or skip.
        
        # 3. Presenter Notes
        grp_notes = QGroupBox("Presenter Notes")
        notes_layout = QVBoxLayout(grp_notes)
        self.txt_notes = QPlainTextEdit()
        self.txt_notes.setPlaceholderText("Enter speaker notes here...")
        notes_layout.addWidget(self.txt_notes)
        layout.addWidget(grp_notes, 1)

        # Signals
        self.chk_paginate.clicked.connect(self._emit_change)
        self.chk_header.clicked.connect(self._emit_change)
        self.chk_footer.clicked.connect(self._emit_change)
        self.txt_notes.textChanged.connect(self._emit_change)
        
        self.on_change_callback = None
        self._updating_ui = False

    def set_on_change(self, callback):
        self.on_change_callback = callback

    def _emit_change(self):
        if self._updating_ui or not self.on_change_callback:
            return
        self.on_change_callback()

    def load_from_slide_text(self, text: str):
        self._updating_ui = True
        try:
            # Paginate
            # <!-- paginate: true --> or <!-- paginate: false -->
            # We assume default is false in many themes, but often true globally. 
            # This checkbox specifically looks for local `paginate: true` or `paginate: false`
            # For simplicity: Check if `paginate: true` is present locally.
            pag_match = re.search(r'<!--\s*_?paginate:\s*(true|false)\s*-->', text, re.IGNORECASE)
            if pag_match:
                self.chk_paginate.setChecked(pag_match.group(1).lower() == 'true')
            else:
                self.chk_paginate.setChecked(False) # Default assumption or 'inherit' - simplified to False for now

            # Header Hidden? <!-- _header: "" -->
            # If `_header: ""` exists, it means HIDDEN -> Checked = False
            header_hide = re.search(r'<!--\s*_header:\s*""\s*-->', text)
            self.chk_header.setChecked(not bool(header_hide))

            # Footer Hidden? <!-- _footer: "" -->
            footer_hide = re.search(r'<!--\s*_footer:\s*""\s*-->', text)
            self.chk_footer.setChecked(not bool(footer_hide))

            # Notes
            # Pattern: <!-- note: ... --> (single line) OR <!--\nnote:\n...\n-->
            # We'll use a simpler regex that tries to capture the content.
            # Limitation: Multiple comments might exist. We'll grab the first 'note' one.
            note_match = re.search(r'<!--\s*note:\s*(.*?)\s*-->', text, re.DOTALL | re.IGNORECASE)
            if note_match:
                self.txt_notes.setPlainText(note_match.group(1).strip())
            else:
                self.txt_notes.setPlainText("")
        finally:
            self._updating_ui = False

    def process_slide_text(self, text: str) -> str:
        # 1. Update/Add Paginate
        if self.chk_paginate.isChecked():
            # Ensure `<!-- paginate: true -->` exists
            if not re.search(r'<!--\s*paginate:\s*true\s*-->', text, re.IGNORECASE):
                # If `paginate: false` exists, replace it
                if re.search(r'<!--\s*paginate:\s*false\s*-->', text, re.IGNORECASE):
                    text = re.sub(r'<!--\s*paginate:\s*false\s*-->', '<!-- paginate: true -->', text, count=1, flags=re.IGNORECASE)
                else:
                    text = '<!-- paginate: true -->\n' + text
        else:
            # We want `paginate: false` OR remove `paginate: true`?
            # User interface implies "Show Page Number".
            # If unchecked, we can force `paginate: false` locally to be sure.
            if not re.search(r'<!--\s*paginate:\s*false\s*-->', text, re.IGNORECASE):
                if re.search(r'<!--\s*paginate:\s*true\s*-->', text, re.IGNORECASE):
                     text = re.sub(r'<!--\s*paginate:\s*true\s*-->', '<!-- paginate: false -->', text, count=1, flags=re.IGNORECASE)
                else:
                    # Append if missing
                    # text = '<!-- paginate: false -->\n' + text
                    # Actually, better not to spam 'false' if it's default. But ensuring 'false' overrides global.
                    pass # user might just want to remove the 'true'. 
                    # Refinement: If it was true, make it false? Or just remove it?
                    # Let's simple toggle: Checkbox True -> `paginate: true`. Checkbox False -> remove `paginate: true` (revert to global/absent).
                    pass
            # Cleanup for unchecked: remove `paginate: true` if present
            text = re.sub(r'<!--\s*paginate:\s*true\s*-->\n?', '', text, flags=re.IGNORECASE)


        # 2. Header Visibility
        # Checked = Show (Remove `_header: ""`)
        # Unchecked = Hide (Add `_header: ""`)
        if self.chk_header.isChecked():
            text = re.sub(r'<!--\s*_header:\s*""\s*-->\n?', '', text)
        else:
            if not re.search(r'<!--\s*_header:\s*""\s*-->', text):
                text = '<!-- _header: "" -->\n' + text

        # 3. Footer Visibility
        if self.chk_footer.isChecked():
            text = re.sub(r'<!--\s*_footer:\s*""\s*-->\n?', '', text)
        else:
            if not re.search(r'<!--\s*_footer:\s*""\s*-->', text):
                text = '<!-- _footer: "" -->\n' + text

        # 4. Notes
        # Remove existing note
        text = re.sub(r'<!--\s*note:\s*(.*?)\s*-->\n?', '', text, flags=re.DOTALL | re.IGNORECASE)
        
        new_note = self.txt_notes.toPlainText().strip()
        if new_note:
            text = f"{text}\n<!--\nnote:\n{new_note}\n-->"

        return text.strip()


class MainWindow(QMainWindow):


    def __init__(self, initial_file: Optional[Path] = None, launch_cwd: Optional[Path] = None):
        super().__init__()
        self.setWindowTitle("Marpit Slide Editor")
        self.resize(1400, 850)
        
        # Application-wide StyleSheet for cleaner, larger buttons
        self.setStyleSheet("""
            QToolBar {
                spacing: 6px;
                padding: 6px;
                background-color: #e0e0e0;
                border-bottom: 1px solid #c0c0c0;
            }
            QToolButton {
                background-color: #ffffff;
                border: 2px solid #a0a0a0;
                border-radius: 4px;
                padding: 6px 12px;
                min-height: 28px;
                font-size: 14px;
                font-weight: bold;
                color: #333333;
            }
            QToolButton:hover {
                background-color: #f0f8ff;
                border: 2px solid #0078d7;
            }
            QToolButton:pressed {
                background-color: #d0d0d0;
                border: 2px solid #005a9e;
            }
            QToolButton::menu-indicator { 
                image: none; 
            }
        """)

        self.launch_cwd = launch_cwd or Path.cwd()

        self.marp_cmd = find_marp_cli_command()
        
        # Initialize Template Manager
        self.marp_cmd = find_marp_cli_command()
        
        # Initialize Template Manager
        self.template_manager = TemplateManager(Path(__file__).parent)

        # Initialize Properties Widget
        self.props_widget = SlidePropertiesWidget(self)
        self.props_widget.set_on_change(self._on_props_changed)
        self.addDockWidget(Qt.RightDockWidgetArea, self.props_widget)

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
        self.recent_files: List[str] = []

        self._build_ui()

        # Default ~2/3 for editor+list, ~1/3 for preview
        self.root_split.setStretchFactor(0, 2)
        self.root_split.setStretchFactor(1, 1)

        # Restore config if available
        self._load_config()

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
        # ---------------------------------------------------------
        # 1. FILE TOOLBAR
        # ---------------------------------------------------------
        file_tb = QToolBar("File")
        file_tb.setIconSize(QSize(24, 24))
        file_tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(file_tb)
        
        style = self.style()

        act_new = QAction("New deck", self)
        act_new.setToolTip("Create a new slide deck (Ctrl+N)")
        act_new.setIcon(style.standardIcon(QStyle.SP_FileIcon))
        act_new.setShortcut(QKeySequence.New)
        act_new.triggered.connect(self.new_deck)
        file_tb.addAction(act_new)

        act_open = QAction("Open…", self)
        act_open.setToolTip("Open an existing Marp Markdown file (Ctrl+O)")
        act_open.setIcon(style.standardIcon(QStyle.SP_DialogOpenButton))
        act_open.setShortcut(QKeySequence.Open)
        act_open.triggered.connect(self.open_deck)
        file_tb.addAction(act_open)

        # Open Recent
        self.btn_recent = QToolButton()
        self.btn_recent.setText("Open Recent ▾")
        self.btn_recent.setToolTip("Open a recently modified file")
        self.btn_recent.setPopupMode(QToolButton.InstantPopup)
        self.menu_recent = QMenu(self.btn_recent)
        self.btn_recent.setMenu(self.menu_recent)
        file_tb.addWidget(self.btn_recent)

        act_save = QAction("Save", self)
        act_save.setToolTip("Save the current deck (Ctrl+S)")
        act_save.setIcon(style.standardIcon(QStyle.SP_DialogSaveButton))
        act_save.setShortcut(QKeySequence.Save)
        act_save.triggered.connect(self.save_deck)
        file_tb.addAction(act_save)

        act_save_as = QAction("Save as…", self)
        act_save_as.setToolTip("Save the current deck as a new file (Ctrl+Shift+S)")
        act_save_as.setShortcut(QKeySequence.SaveAs)
        act_save_as.triggered.connect(self.save_deck_as)
        file_tb.addAction(act_save_as)

        file_tb.addSeparator()

        # Templates Menu
        btn_tmpl = QToolButton()
        btn_tmpl.setText("Templates ▾")
        btn_tmpl.setToolTip("Create new or insert slides from templates")
        btn_tmpl.setPopupMode(QToolButton.InstantPopup)
        menu_tmpl = QMenu(btn_tmpl)
        
        # Populate templates dynamically
        self.menu_new_tmpl = menu_tmpl.addMenu("New from Template")
        self.menu_ins_tmpl = menu_tmpl.addMenu("Insert Template")
        # We'll populate these on show/init
        self._populate_template_menus()
        
        btn_tmpl.setMenu(menu_tmpl)
        file_tb.addWidget(btn_tmpl)

        file_tb.addSeparator()

        act_export = QAction("Generate slides…", self)
        act_export.setToolTip("Export the deck to PDF, HTML, or PowerPoint")
        act_export.triggered.connect(self.export_deck)
        file_tb.addAction(act_export)

        file_tb.addSeparator()

        self.act_allow_local = QAction("Allow local files", self)
        self.act_allow_local.setToolTip("Allow loading local images and resources in preview")
        self.act_allow_local.setCheckable(True)
        self.act_allow_local.setChecked(True)
        self.act_allow_local.triggered.connect(self._toggle_allow_local_files)
        file_tb.addAction(self.act_allow_local)

        # ---------------------------------------------------------
        # 2. DECK TOOLBAR (Global Settings)
        # ---------------------------------------------------------
        deck_tb = QToolBar("Deck")
        deck_tb.setIconSize(QSize(24, 24))
        deck_tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(deck_tb)

        deck_btn = QToolButton()
        deck_btn.setText("Deck ▾")
        deck_btn.setToolTip("Global presentation settings")
        deck_btn.setPopupMode(QToolButton.InstantPopup)
        deck_menu = QMenu(deck_btn)

        # Global Settings
        deck_menu.addAction("Set global background...", self.set_global_background)
        deck_menu.addAction("Set global header/footer...", self.set_global_header_footer)
        deck_menu.addAction("Pagination...", self.edit_pagination)
        deck_menu.addAction("Edit Front-matter (YAML)...", self.edit_deck_directives)
        deck_menu.addSeparator()

        # Themes
        themes_menu = deck_menu.addMenu("Themes")
        themes_menu.addAction("Default", lambda: self.insert_template("theme: default\n"))
        themes_menu.addAction("Gaia", lambda: self.insert_template("theme: gaia\n"))
        themes_menu.addAction("Uncover", lambda: self.insert_template("theme: uncover\n"))

        deck_btn.setMenu(deck_menu)
        deck_tb.addWidget(deck_btn)

        # ---------------------------------------------------------
        # 3. SLIDE TOOLBAR (Slide Management & Local Style)
        # ---------------------------------------------------------
        slide_tb = QToolBar("Slide")
        slide_tb.setIconSize(QSize(24, 24))
        slide_tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(slide_tb)

        act_add = QAction("Add", self) # Shortened text for toolbar
        act_add.setToolTip("Add a new slide after the current one (Ctrl+Shift+N)")
        act_add.setIcon(style.standardIcon(QStyle.SP_FileDialogNewFolder)) # Best approx for "Add"
        act_add.setShortcut(QKeySequence("Ctrl+Shift+N"))
        act_add.triggered.connect(self.add_slide_after_current)
        slide_tb.addAction(act_add)

        act_del = QAction("Delete", self)
        act_del.setToolTip("Delete the current slide (Ctrl+Shift+Del)")
        act_del.setIcon(style.standardIcon(QStyle.SP_TrashIcon))
        act_del.setShortcut(QKeySequence("Ctrl+Shift+Del"))
        act_del.triggered.connect(self.delete_current_slide)
        slide_tb.addAction(act_del)

        slide_tb.addSeparator()

        slide_btn = QToolButton()
        slide_btn.setText("Slide Options ▾")
        slide_btn.setToolTip("Slide-specific options (backgrounds, styles)")
        slide_btn.setPopupMode(QToolButton.InstantPopup)
        slide_menu = QMenu(slide_btn)

        # Backgrounds
        bg_menu = slide_menu.addMenu("Background Image")
        bg_menu.addAction("Image", lambda: self.insert_template("![bg](path-or-url)\n"))
        bg_menu.addAction("Image (Cover)", lambda: self.insert_template("![bg cover](path-or-url)\n"))
        bg_menu.addAction("Image (Contain)", lambda: self.insert_template("![bg contain](path-or-url)\n"))
        bg_menu.addSeparator()
        bg_menu.addAction("Split Left", lambda: self.insert_template("![bg left](path-or-url)\n\n"))
        bg_menu.addAction("Split Right", lambda: self.insert_template("![bg right](path-or-url)\n\n"))

        # Styles
        style_menu = slide_menu.addMenu("Slide Style")
        style_menu.addAction("Lead (Centered)", lambda: self.insert_template("<!-- _class: lead -->\n"))
        style_menu.addAction("Invert (Dark)", lambda: self.insert_template("<!-- _class: invert -->\n"))

        slide_btn.setMenu(slide_menu)
        slide_tb.addWidget(slide_btn)

        # ---------------------------------------------------------
        # 4. FORMATTING TOOLBAR (Text & Content)
        # ---------------------------------------------------------

        # Central layout: left half (deck overview + editor) and right half (preview)
        self.root_split = QSplitter(Qt.Horizontal)
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
        fmt_tb.setIconSize(QSize(24, 24))
        fmt_tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        editor_layout.addWidget(fmt_tb)

        # Text Formatting (Using Emojis/Text as Icons)
        fmt_tb.addAction(self._make_action("H1", lambda: self.set_line_heading(1), "Heading 1 (Slide Title)"))
        fmt_tb.addAction(self._make_action("H2", lambda: self.set_line_heading(2), "Heading 2"))
        fmt_tb.addAction(self._make_action("H3", lambda: self.set_line_heading(3), "Heading 3"))
        fmt_tb.addSeparator()
        fmt_tb.addAction(self._make_action("📝 List", lambda: self.set_line_list("-"), "Bulleted List"))
        fmt_tb.addAction(self._make_action("🔢 Num", self.set_line_numbered, "Numbered List"))
        fmt_tb.addSeparator()
        
        # Math Support
        fmt_tb.addAction(self._make_action("∑", self.insert_inline_math, "Insert Inline Math ($...$)"))
        fmt_tb.addAction(self._make_action("$$", self.insert_block_math, "Insert Block Math ($$...$$)"))
        fmt_tb.addSeparator()

        # Insert Menu
        ins_btn = QToolButton()
        ins_btn.setText("Insert ▾")
        ins_btn.setToolTip("Insert images, code blocks, quotes, etc.")
        ins_btn.setPopupMode(QToolButton.InstantPopup)
        ins_menu = QMenu(ins_btn)

        ins_menu.addAction("Insert Picture...", self.insert_picture_dialog)
        ins_menu.addSeparator()
        ins_menu.addAction("Inline Image", lambda: self.insert_template("![](path-or-url)\n"))
        ins_menu.addAction("Quote", self.set_line_quote)
        ins_menu.addAction("Code Block", self.insert_code_block)
        ins_menu.addAction("Presenter Note", self.insert_presenter_note)
        ins_menu.addSeparator()
        ins_menu.addAction("New Slide (Separator)", lambda: self.insert_template("\n---\n\n"))

        ins_btn.setMenu(ins_menu)
        fmt_tb.addWidget(ins_btn)

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

        self.root_split.addWidget(left_split)
        self.root_split.addWidget(preview_panel)
        # Default ~2/3 for editor+list, ~1/3 for preview
        self.root_split.setStretchFactor(0, 2)
        self.root_split.setStretchFactor(1, 1)

        self.setCentralWidget(self.root_split)

        # Status bar
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._update_status()

    # ---------------- Props / Sync ----------------
    def _on_props_changed(self):
        """Called when a checkbox or note in the properties panel changes."""
        if self._updating_editor:
            return

        # 1. Get current text from editor
        current_txt = self.editor.toPlainText()
        
        # 2. Process via widget logic to apply changes
        new_txt = self.props_widget.process_slide_text(current_txt)
        
        # 3. Update editor (this will trigger _on_editor_text_changed -> updating deck)
        if new_txt != current_txt:
            scroll = self.editor.verticalScrollBar().value()
            cursor = self.editor.textCursor()
            pos = cursor.position()
            
            self.editor.setPlainText(new_txt)
            
            # Try to restore cursor/scroll
            self.editor.verticalScrollBar().setValue(scroll)
            # cursor.setPosition(min(pos, len(new_txt))) # naive restore
            # self.editor.setTextCursor(cursor)

    def _make_action(self, text: str, slot, tooltip: str = ""):
        act = QAction(text, self)
        act.triggered.connect(slot)
        if tooltip:
            act.setToolTip(tooltip)
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
            txt = self.deck.slides[idx] or ""
            self.editor.setPlainText(txt)
            # Sync properties panel from text
            self.props_widget.load_from_slide_text(txt)
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
        
        # Sync properties panel from text (NEW)
        # We need to guard against recursion if prop widget updates trigger editor updates
        # But load_from_slide_text sets internal flag so it won't emit back
        self.props_widget.load_from_slide_text(txt)

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
        p = p.resolve()
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

        # Switch CWD to the loaded file's directory
        try:
            os.chdir(p.parent)
        except Exception:
            pass

        self._update_window_title()
        self._update_status("Deck loaded.")
        self._add_recent_file(p)
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
            self._add_recent_file(p)
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
        # Switch CWD to the saved file's directory
        try:
            os.chdir(path.parent)
        except Exception:
            pass

        self._update_window_title()
        self._update_status("Saved.")
        self._add_recent_file(path)
        return True

        self._schedule_preview()

    def _populate_template_menus(self):
        self.menu_new_tmpl.clear()
        self.menu_ins_tmpl.clear()
        
        templates = self.template_manager.get_available_templates()
        if not templates:
            non = QAction("(No templates found)", self)
            non.setEnabled(False)
            self.menu_new_tmpl.addAction(non)
            self.menu_ins_tmpl.addAction(non)
            return

        for p in templates:
            name = p.stem.replace("_", " ")
            
            # New from Template
            act_new = QAction(name, self)
            act_new.triggered.connect(lambda checked=False, path=p: self.new_from_template(path))
            self.menu_new_tmpl.addAction(act_new)

            # Insert Template
            act_ins = QAction(name, self)
            act_ins.triggered.connect(lambda checked=False, path=p: self.insert_template_from_file(path))
            self.menu_ins_tmpl.addAction(act_ins)

    def new_from_template(self, path: Path):
        if not self._confirm_discard_if_dirty():
            return
        
        preamble, slides = self.template_manager.load_template(path)
        self.deck = DeckState(preamble=preamble, slides=slides, file_path=None, dirty=True)
        self._current_slide_idx = 0
        self._refresh_slide_list()
        self._load_slide_into_editor(0)
        self._update_window_title()
        self._update_status(f"Created new deck from template: {path.stem}")
        self._schedule_preview()

    def insert_template_from_file(self, path: Path):
        self._maybe_commit_current_editor()
        
        _, new_slides = self.template_manager.load_template(path)
        # Append logic? or Insert at cursor? Let's append for now or insert after current
        insert_at = self._current_slide_idx + 1
        
        for s in reversed(new_slides):
            self.deck.slides.insert(insert_at, s)
        
        self.deck.dirty = True
        self._refresh_slide_list()
        # Jump to first inserted slide
        self._current_slide_idx = insert_at
        self._load_slide_into_editor(insert_at)
        self._update_window_title()
        self._update_status(f"Inserted template slides: {path.stem}")
        self._schedule_preview()
        
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

    def set_global_background(self):
        if not self.deck.file_path:
            QMessageBox.warning(self, "Save first", "Please save the deck first so we can resolve relative paths.")
            return



        dlg = GlobalBackgroundDialog(self, base_path=self.deck.file_path)

        # Check for existing background settings to pre-fill
        # Check for existing background settings to pre-fill
        pre = self.deck.preamble or "---\n\n---\n\n"

        # Robust parsing: Find "style:" line, then grab following indented lines (the block)
        lines = pre.strip().split("\n")
        style_block_lines = []
        in_style = False

        for line in lines:
            if re.match(r"^style\s*:", line):
                in_style = True
                continue
            if in_style:
                # If indented, it's part of the block
                if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                    style_block_lines.append(line)
                else:
                    # End of block
                    break

        import textwrap
        style_block = textwrap.dedent("\n".join(style_block_lines))
        existing_found = False

        if style_block:
            # Try to extract current values
            # Looking for url('...'), background-size: ..., opacity: ...

            # Check for legacy "section::before" style OR new "linear-gradient" style
            if "section::before" in style_block:
                existing_found = True
                # Parse Legacy ::before
                m_url = re.search(r"background-image:\s*url\(['\"](.*?)['\"]\)", style_block)
                if m_url:
                    url = m_url.group(1).replace("%22", '"').replace("%27", "'")
                    m_size = re.search(r"background-size:\s*([^;]+)", style_block)
                    size = m_size.group(1).strip() if m_size else "cover"
                    m_op = re.search(r"opacity:\s*([\d.]+)", style_block)
                    op = float(m_op.group(1)) if m_op else 1.0

                    dlg.load_settings(url, size, op, "center center")
                    m_pos = re.search(r"background-position:\s*([^;]+)", style_block)
                    if m_pos:
                        dlg.load_settings(url, size, op, m_pos.group(1).strip())

            elif "linear-gradient" in style_block:
                existing_found = True
                # Parse New Gradient style
                # background-image: linear-gradient(rgba(255,255,255,0.50), ...), url('...')
                m_url = re.search(r"url\(['\"](.*?)['\"]\)", style_block)
                if m_url:
                    url = m_url.group(1).replace("%22", '"').replace("%27", "'")
                    m_size = re.search(r"background-size:\s*([^;]+)", style_block)
                    size = m_size.group(1).strip() if m_size else "cover"

                    # Opacity is stuck in rgba(..., alpha)
                    # We look for the first rgba(... alpha)
                    m_alpha = re.search(r"rgba\(255,\s*255,\s*255,\s*([\d.]+)\)", style_block)
                    alpha = float(m_alpha.group(1)) if m_alpha else 0.0
                    op = 1.0 - alpha

                    dlg.load_settings(url, size, op, "center center")
                    m_pos = re.search(r"background-position:\s*([^;]+)", style_block)
                    if m_pos:
                        dlg.load_settings(url, size, op, m_pos.group(1).strip())

        if dlg.exec():
            css = dlg.get_css_content()
            if not css:
                return

            # Construct new style block
            new_style_block = "style: |\n"
            for line in css.split("\n"):
                new_style_block += f"  {line}\n"

            # Parse preamble lines
            lines = pre.strip().split("\n")
            if not (len(lines) >= 2 and lines[0].strip() == "---" and lines[-1].strip() == "---"):
                 QMessageBox.warning(self, "Error", "Could not parse YAML front-matter automatically.")
                 return

            # Remove last '---' temporarily
            last_marker = lines.pop()

            # Find existing style block lines to remove
            start_idx = -1
            end_idx = -1

            for i, line in enumerate(lines):
                if re.match(r"^style\s*:", line):
                    start_idx = i
                    break

            if start_idx != -1:
                # Found start, find end.
                # Block ends when indentation returns to 0 or file ends
                end_idx = len(lines)
                for j in range(start_idx + 1, len(lines)):
                    if lines[j].strip() and not lines[j].startswith("  "):
                        # Found a line that is NOT indented (and not empty) -> end of block
                        end_idx = j
                        break

                # Check if it was "our" style block or something else
                if not existing_found and not "section::before" in "\n".join(lines[start_idx:end_idx]):
                     # User has some OTHER style block. Warn.
                     res = QMessageBox.question(
                         self, "Overwrite Styles?",
                         "You have existing custom styles defined in the YAML header.\n"
                         "Overwrite them with the new background settings?",
                         QMessageBox.Yes | QMessageBox.No
                     )
                     if res != QMessageBox.Yes:
                         # Revert to copy-paste fallback.
                         QApplication.clipboard().setText(css)
                         self.edit_deck_directives()
                         return

                # Remove existing block
                del lines[start_idx:end_idx]

            # Append new block
            lines.append(new_style_block.rstrip())
            lines.append(last_marker)

            self.deck.preamble = "\n".join(lines) + "\n\n"
            self.deck.dirty = True

            self._update_window_title()
            self._update_status("Global background updated.")
            self._schedule_preview()


    def set_global_header_footer(self):
        if not self.deck.file_path:
            QMessageBox.warning(self, "Save first", "Please save the deck first so we can resolve relative paths.")
            return

        dlg = GlobalHeaderFooterDialog(self, base_path=self.deck.file_path)

        # Parse existing header/footer
        pre = self.deck.preamble or "---\n\n---\n\n"

        # 1. Extract header/footer strings
        header_content = ""
        footer_content = ""
        m_head = re.search(r"^header:\s*(.*)$", pre, re.MULTILINE)
        if m_head:
            header_content = m_head.group(1).strip().strip("'\"")
        m_foot = re.search(r"^footer:\s*(.*)$", pre, re.MULTILINE)
        if m_foot:
            footer_content = m_foot.group(1).strip().strip("'\"")

        # 2. Extract CSS for header/footer
        # We look for the style block again
        style_block_lines = []
        in_style = False
        lines = pre.strip().split("\n")

        import textwrap

        for line in lines:
            if re.match(r"^style\s*:", line):
                in_style = True
                continue
            if in_style:
                if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                    style_block_lines.append(line)
                else:
                    break

        raw_block = "\n".join(style_block_lines)
        style_block = textwrap.dedent(raw_block)

        # Helpers to extract CSS props
        def extract_css(selector, prop, default):
            m_sel = re.search(rf"{selector}\s*\{{([^}}]*)\}}", style_block, re.DOTALL)
            if m_sel:
                block_content = m_sel.group(1)
                m_prop = re.search(rf"{prop}:\s*([^;]+)", block_content)
                if m_prop:
                    val = m_prop.group(1).strip()
                    if val.endswith("px"):
                        return int(val[:-2])
                    return val
            return default

        # Header CSS
        h_height = extract_css("header", "height", 100)
        h_top = extract_css("header", "top", 0)
        h_size = extract_css("header", "font-size", 18)
        h_color = extract_css("header", "color", "")
        h_fam = extract_css("header", "font-family", "")
        if isinstance(h_top, str): h_top = 0

        h_align = extract_css("header", "text-align", "left")
        h_display = extract_css("header", "display", "")
        if "flex" in str(h_display):
             h_align = "spread"

        # Footer CSS
        f_height = extract_css("footer", "height", 50)
        f_bottom = extract_css("footer", "bottom", 0)
        f_size = extract_css("footer", "font-size", 18)
        f_color = extract_css("footer", "color", "")
        f_fam = extract_css("footer", "font-family", "")
        if isinstance(f_bottom, str): f_bottom = 0

        f_align = extract_css("footer", "text-align", "left")
        f_display = extract_css("footer", "display", "")
        if "flex" in str(f_display):
             f_align = "spread"

        dlg.load_settings({
            'header': {'content': header_content, 'height': h_height, 'offset': h_top, 'align': h_align, 'font_size': h_size, 'font_family': h_fam, 'color': h_color},
            'footer': {'content': footer_content, 'height': f_height, 'offset': f_bottom, 'align': f_align, 'font_size': f_size, 'font_family': f_fam, 'color': f_color}
        })

        if dlg.exec():
            data = dlg.get_settings()

            existing_lines = pre.strip().split("\n")
            if existing_lines and existing_lines[0] == "---": existing_lines.pop(0)
            if existing_lines and existing_lines[-1] == "---": existing_lines.pop()

            new_lines = []
            in_style = False
            
            for line in existing_lines:
                if re.match(r"^header\s*:", line) or re.match(r"^footer\s*:", line):
                    continue
                if re.match(r"^style\s*:", line):
                    in_style = True
                    continue
                if in_style:
                    if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                        continue
                    else:
                        in_style = False
                new_lines.append(line)

            common_css = style_block
            common_css = re.sub(r"header\s*\{[^}]*\}", "", common_css)
            common_css = re.sub(r"footer\s*\{[^}]*\}", "", common_css)
            common_css = re.sub(r"\n{3,}", "\n\n", common_css).strip()

            if data['header']['content']:
                 safe_h = data['header']['content'].replace('"', '\\"')
                 new_lines.append(f'header: "{safe_h}"')
            
            if data['footer']['content']:
                 safe_f = data['footer']['content'].replace('"', '\\"')
                 new_lines.append(f'footer: "{safe_f}"')

            h_css = ""
            if data['header']['content']: 
                d = data['header']
                align = d['align']
                disp = "block"
                if align == "spread":
                    align = "left"
                    disp = "flex; justify-content: space-between"
                
                h_css = f"header {{\n  height: {d['height']}px;\n  top: {d['offset']}px;\n  font-size: {d['font_size']}px;\n  text-align: {align};\n"
                if disp != "block": h_css += f"  display: {disp};\n"
                if d['font_family']: h_css += f"  font-family: {d['font_family']};\n"
                if d['color']: h_css += f"  color: {d['color']};\n"
                h_css += "}\n"

            f_css = ""
            if data['footer']['content']:
                d = data['footer']
                align = d['align']
                disp = "block"
                if align == "spread":
                    align = "left"
                    disp = "flex; justify-content: space-between"

                f_css = f"footer {{\n  height: {d['height']}px;\n  bottom: {d['offset']}px;\n  font-size: {d['font_size']}px;\n  text-align: {align};\n"
                if disp != "block": f_css += f"  display: {disp};\n"
                if d['font_family']: f_css += f"  font-family: {d['font_family']};\n"
                if d['color']: f_css += f"  color: {d['color']};\n"
                f_css += "}\n"

            final_style = common_css
            if h_css: final_style += "\n" + h_css
            if f_css: final_style += "\n" + f_css
            
            if final_style.strip():
                new_lines.append("style: |")
                for l in final_style.strip().split("\n"):
                    new_lines.append(f"  {l}")

        # Helper to dedent block
        import textwrap

        for line in lines:
            if re.match(r"^style\s*:", line):
                in_style = True
                continue
            if in_style:
                if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                    style_block_lines.append(line)
                else:
                    break

        # Join then dedent to normalize
        # Note: preserve empty lines
        raw_block = "\n".join(style_block_lines)
        style_block = textwrap.dedent(raw_block)

        # Helpers to extract CSS props
        def extract_css(selector, prop, default):
            # Regex for "selector { ... prop: val; ... }"
            # Simplified: look for "selector {" then scan for prop
            m_sel = re.search(rf"{selector}\s*\{{([^}}]*)\}}", style_block, re.DOTALL)
            if m_sel:
                block_content = m_sel.group(1)
                m_prop = re.search(rf"{prop}:\s*([^;]+)", block_content)
                if m_prop:
                    val = m_prop.group(1).strip()
                    # clean optional "!important" for reading back
                    val = val.replace("!important", "").strip()
                    # clean "px"
                    if val.endswith("px"):
                        return int(val[:-2])
                    return val
            return default

        # Header CSS
        # Header CSS
        h_height = extract_css("header", "height", 100)
        h_top = extract_css("header", "top", 0)
        h_size = extract_css("header", "font-size", 18)
        h_color = extract_css("header", "color", "")
        h_fam = extract_css("header", "font-family", "")
        if isinstance(h_top, str): h_top = 0 # Safety

        # Header Align
        # Check text-align first
        h_align = extract_css("header", "text-align", "left")
        # Check if spread (display: flex)
        h_display = extract_css("header", "display", "")
        if "flex" in str(h_display):
             h_align = "spread"

        # Footer CSS
        f_height = extract_css("footer", "height", 50)
        f_bottom = extract_css("footer", "bottom", 0)
        f_size = extract_css("footer", "font-size", 18)
        f_color = extract_css("footer", "color", "")
        f_fam = extract_css("footer", "font-family", "")
        if isinstance(f_bottom, str): f_bottom = 0

        f_align = extract_css("footer", "text-align", "left")
        f_display = extract_css("footer", "display", "")
        if "flex" in str(f_display):
             f_align = "spread"

        dlg.load_settings({
            'header': {'content': header_content, 'height': h_height, 'offset': h_top, 'align': h_align, 'font_size': h_size, 'font_family': h_fam, 'color': h_color},
            'footer': {'content': footer_content, 'height': f_height, 'offset': f_bottom, 'align': f_align, 'font_size': f_size, 'font_family': f_fam, 'color': f_color}
        })

        if dlg.exec():
            data = dlg.get_settings()

            # Reconstruct Preamble

            # Remove existing header/footer keys
            existing_lines = pre.strip().split("\n")
            if existing_lines and existing_lines[0] == "---": existing_lines.pop(0)
            if existing_lines and existing_lines[-1] == "---": existing_lines.pop()

            # Filter out header/footer/style

            # Let's rebuild the style block
            # Retrieve existing style content MINUS header/footer rules

            clean_style_block = style_block
            # Remove entire header { ... } block
            clean_style_block = re.sub(r"header\s*\{[^}]*\}", "", clean_style_block)
            clean_style_block = re.sub(r"footer\s*\{[^}]*\}", "", clean_style_block)
            clean_style_block = "\n".join([l for l in clean_style_block.split("\n") if l.strip()])

            # Generate new CSS for header/footer
            def gen_css(sel, d, is_top=True):
                css = f"{sel} {{\n"
                css += f"  height: {d['height']}px;\n"
                css += f"  font-size: {d['font_size']}px;\n"
                if d.get('color'):
                    css += f"  color: {d['color']};\n"
                if d.get('font_family'):
                    css += f"  font-family: {d['font_family']};\n"
                if d['align'] == 'spread':
                    css += "  display: flex !important;\n"
                    css += "  justify-content: space-between !important;\n"
                    css += "  align-items: center;\n"
                    css += "  text-align: left;\n" # fallback
                else:
                    css += f"  text-align: {d['align']} !important;\n"
                    css += "  display: block;\n" # reset

                # Position
                if d['offset'] != 0:
                    prop = "top" if is_top else "bottom"
                    css += f"  {prop}: {d['offset']}px;\n"
                else:
                    # Explicit reset if user sets 0 but maybe was not 0 before?
                    # Good practice to set 0 if 0
                    pass

                css += "}\n"
                return css

            new_css = ""
            if data['header']['content']:
                new_css += gen_css("header", data['header'], True)
            if data['footer']['content']:
                new_css += gen_css("footer", data['footer'], False)

            final_style_block = clean_style_block + "\n" + new_css

            # Parse lines to keep (non-header/footer/style)
            final_lines = ["---"]

            skip_indent = False
            for line in existing_lines:
                if re.match(r"^(header|footer)\s*:", line):
                    continue
                if re.match(r"^style\s*:", line):
                    skip_indent = True
                    continue
                if skip_indent:
                    if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                        continue
                    else:
                        skip_indent = False

                final_lines.append(line)

            # Append new keys
            if data['header']['content']:
                c = data['header']['content'].replace('"', '\\"')
                final_lines.append(f'header: "{c}"')
            if data['footer']['content']:
                c = data['footer']['content'].replace('"', '\\"')
                final_lines.append(f'footer: "{c}"')

            # Append Style
            if final_style_block.strip():
                final_lines.append("style: |")
                for l in final_style_block.split("\n"):
                    if l.strip():
                        final_lines.append(f"  {l}")

            # Re-assemble
            self.deck.preamble = "---\n" + "\n".join(final_lines).strip() + "\n---\n\n"
            self.deck.dirty = True
            self._update_window_title()
            self._update_status("Global header/footer updated.")
            self._schedule_preview()






    def edit_pagination(self):
        if not self.deck.file_path:
            QMessageBox.warning(self, "Save first", "Please save the deck first.")
            return

        pre = self.deck.preamble or "---\n\n---\n\n"
        
        is_paginated = False
        if re.search(r"^paginate:\s*true", pre, re.MULTILINE | re.IGNORECASE):
            is_paginated = True
            
        style_block = ""
        in_style = False
        for line in pre.split("\n"):
            if re.match(r"^style\s*:", line):
                in_style = True
                continue
            if in_style:
                if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                    style_block += line + "\n"
                else:
                    break
        
        import textwrap
        style_block = textwrap.dedent(style_block)

        s_size = 18
        s_font = ""
        s_color = ""
        s_pos = "bottom-right"
        
        m_sect = re.search(r"section::after\s*\{([^}]*)\}", style_block, re.DOTALL)
        if m_sect:
            blk = m_sect.group(1)
            m = re.search(r"font-size:\s*(\d+)px", blk)
            if m: s_size = int(m.group(1))
            m = re.search(r"font-family:\s*([^;]+)", blk)
            if m: s_font = m.group(1).strip()
            m = re.search(r"color:\s*([^;]+)", blk)
            if m: s_color = m.group(1).strip()
            
            is_top = "top:" in blk
            is_left = "left:" in blk
            if is_top and is_left: s_pos = "top-left"
            elif is_top: s_pos = "top-right"
            elif is_left: s_pos = "bottom-left"
            else: s_pos = "bottom-right"

        dlg = PaginationDialog(self, self.deck.file_path)
        dlg.load_settings(is_paginated, s_size, s_font, s_color, s_pos)
        
        if dlg.exec():
            d = dlg.get_settings()
            
            existing_lines = pre.strip().split("\n")
            if existing_lines and existing_lines[0] == "---": existing_lines.pop(0)
            if existing_lines and existing_lines[-1] == "---": existing_lines.pop()
            
            new_lines = []
            in_style = False
            
            for line in existing_lines:
                if re.match(r"^paginate\s*:", line): continue
                if re.match(r"^style\s*:", line): 
                    in_style = True
                    continue
                if in_style:
                    if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                        continue
                    else:
                        in_style = False
                new_lines.append(line)
            
            if d['enabled']:
                new_lines.append("paginate: true")
            else:
                new_lines.append("paginate: false")

            common_css = style_block
            common_css = re.sub(r"section::after\s*\{[^}]*\}", "", common_css)
            common_css = re.sub(r"\n{3,}", "\n\n", common_css).strip()
            
            pag_css = ""
            if d['enabled']:
                pos_css = ""
                if d['pos'] == "bottom-right":
                    pos_css = "bottom: 10px; right: 20px; top: auto; left: auto;"
                elif d['pos'] == "bottom-left":
                    pos_css = "bottom: 10px; left: 20px; top: auto; right: auto;"
                elif d['pos'] == "top-right":
                    pos_css = "top: 10px; right: 20px; bottom: auto; left: auto;"
                elif d['pos'] == "top-left":
                    pos_css = "top: 10px; left: 20px; bottom: auto; right: auto;"
                
                font_css = ""
                if d['font']: font_css += f"  font-family: {d['font']};\n"
                col_css = ""
                if d['color']: col_css += f"  color: {d['color']};\n"
                
                pag_css = f"section::after {{\n  font-size: {d['size']}px;\n{font_css}{col_css}  {pos_css}\n}}"

            final_style = common_css
            if pag_css: final_style += "\n" + pag_css
            
            if final_style.strip():
                new_lines.append("style: |")
                for l in final_style.strip().split("\n"):
                    new_lines.append(f"  {l}")

            self.deck.preamble = "---\n" + "\n".join(new_lines).strip() + "\n---\n\n"

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

        # Determine where to write the preview file
        # If we have a real file, write a hidden preview file in the same dir so relative paths work.
        if self.deck.file_path:
            self._preview_md_path = self.deck.file_path.parent / ".marp_preview.md"
            self._preview_html_path = self.deck.file_path.parent / ".marp_preview.html"
        else:
            # Fallback to temp dir
            self._preview_md_path = Path(self._tmp_dir.name) / "preview.md"
            self._preview_html_path = Path(self._tmp_dir.name) / "preview.html"

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
        args.append("--html")
        if self.allow_local_files:
            args.append("--allow-local-files")

        program = self.marp_cmd[0]

        # Use a simple subprocess to keep this single-file; rendering one slide is fast.
        # If it fails, show stderr.
        import subprocess

        try:
            # We want to run in the same dir as the preview file
            cwd = self._preview_md_path.parent

            proc = subprocess.run(
                [program] + args,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=str(cwd),
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
        args.append("--html")

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

    # ---------------- Config Persistence ----------------
    def _get_config_path(self) -> Path:
        return self.launch_cwd / ".marpit_editor_config.json"

    def _load_config(self):
        cfg_path = self._get_config_path()
        if not cfg_path.exists():
            self._update_recent_menu() # init empty
            return

        try:
            data = json.loads(cfg_path.read_text(encoding="utf-8"))
            w = data.get("window_width")
            h = data.get("window_height")
            if w and h:
                self.resize(w, h)

            sizes = data.get("splitter_sizes")
            if sizes and isinstance(sizes, list) and len(sizes) == 2:
                self.root_split.setSizes(sizes)

            # Load recent files
            self.recent_files = data.get("recent_files", [])
            # Filter non-existing files? Maybe not, network drives might be offline.
            self._update_recent_menu()

        except Exception:
            pass # Ignore config errors

    def _save_config(self):
        cfg_path = self._get_config_path()
        data = {
            "window_width": self.width(),
            "window_height": self.height(),
            "splitter_sizes": self.root_split.sizes(),
            "recent_files": self.recent_files
        }
        try:
            cfg_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _add_recent_file(self, path: Path):
        p_str = str(path.resolve())
        # Remove if exists
        if p_str in self.recent_files:
            self.recent_files.remove(p_str)
        # Add to top
        self.recent_files.insert(0, p_str)
        # Cap at 10
        self.recent_files = self.recent_files[:10]
        self._update_recent_menu()
        self._save_config()

    def _update_recent_menu(self):
        self.menu_recent.clear()
        if not self.recent_files:
            act = QAction("No recent files", self)
            act.setEnabled(False)
            self.menu_recent.addAction(act)
            return

        for fpath in self.recent_files:
             # Use path as text (checking if valid?)
             # Truncate if too long?
             fname = Path(fpath).name
             # Show name, toolip full path
             # Or show full path if ambiguous?
             # Let's show "Name (Path)" or just Path
             # Path is clearer for now
             act = QAction(fname, self)
             act.setToolTip(fpath)
             act.setData(fpath)
             # Use lambda with default arg to capture fpath properly
             act.triggered.connect(lambda checked=False, p=fpath: self.load_from_path(Path(p)))
             self.menu_recent.addAction(act)

        self.menu_recent.addSeparator()
        act_clear = QAction("Clear recent files", self)
        act_clear.triggered.connect(self._clear_recent)
        self.menu_recent.addAction(act_clear)

    def _clear_recent(self):
        self.recent_files = []
        self._update_recent_menu()
        self._save_config()

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
        self._save_config()
        if not self._confirm_discard_if_dirty():
            event.ignore()
            return
        try:
            self._tmp_dir.cleanup()
        except Exception:
            pass
        event.accept()

    def insert_inline_math(self):
        """Insert $  $ and place cursor in middle."""
        cursor = self.editor.textCursor()
        cursor.insertText("$  $")
        cursor.movePosition(QTextCursor.Left, QTextCursor.MoveAnchor, 2)
        self.editor.setTextCursor(cursor)
        self.editor.setFocus()

    def insert_block_math(self):
        """Insert block math $$ ... $$ and place cursor in middle."""
        cursor = self.editor.textCursor()
        # Check if we are at start of line, else insert newline
        # Simple approach: Insert \n$$\n\n$$\n
        text = "\n$$\n\n$$\n"
        cursor.insertText(text)
        # Move back up 2 lines
        cursor.movePosition(QTextCursor.Up, QTextCursor.MoveAnchor, 2)
        self.editor.setTextCursor(cursor)
        self.editor.setFocus()

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

        # Determine where to write the preview file
        # If we have a real file, write a hidden preview file in the same dir so relative paths work.
        if self.deck.file_path:
            self._preview_md_path = self.deck.file_path.parent / ".marp_preview.md"
            self._preview_html_path = self.deck.file_path.parent / ".marp_preview.html"
        else:
            # Fallback to temp dir
            self._preview_md_path = Path(self._tmp_dir.name) / "preview.md"
            self._preview_html_path = Path(self._tmp_dir.name) / "preview.html"

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
        args.append("--html")
        if self.allow_local_files:
            args.append("--allow-local-files")

        program = self.marp_cmd[0]

        # Use a simple subprocess to keep this single-file; rendering one slide is fast.
        # If it fails, show stderr.
        import subprocess

        try:
            # We want to run in the same dir as the preview file
            cwd = self._preview_md_path.parent

            proc = subprocess.run(
                [program] + args,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=str(cwd),
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
        args.append("--html")

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

    # ---------------- Config Persistence ----------------
    def _get_config_path(self) -> Path:
        return self.launch_cwd / ".marpit_editor_config.json"

    def _load_config(self):
        cfg_path = self._get_config_path()
        if not cfg_path.exists():
            self._update_recent_menu() # init empty
            return

        try:
            data = json.loads(cfg_path.read_text(encoding="utf-8"))
            w = data.get("window_width")
            h = data.get("window_height")
            if w and h:
                self.resize(w, h)

            sizes = data.get("splitter_sizes")
            if sizes and isinstance(sizes, list) and len(sizes) == 2:
                self.root_split.setSizes(sizes)

            # Load recent files
            self.recent_files = data.get("recent_files", [])
            # Filter non-existing files? Maybe not, network drives might be offline.
            self._update_recent_menu()

        except Exception:
            pass # Ignore config errors

    def _save_config(self):
        cfg_path = self._get_config_path()
        data = {
            "window_width": self.width(),
            "window_height": self.height(),
            "splitter_sizes": self.root_split.sizes(),
            "recent_files": self.recent_files
        }
        try:
            cfg_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _add_recent_file(self, path: Path):
        p_str = str(path.resolve())
        # Remove if exists
        if p_str in self.recent_files:
            self.recent_files.remove(p_str)
        # Add to top
        self.recent_files.insert(0, p_str)
        # Cap at 10
        self.recent_files = self.recent_files[:10]
        self._update_recent_menu()
        self._save_config()

    def _update_recent_menu(self):
        self.menu_recent.clear()
        if not self.recent_files:
            act = QAction("No recent files", self)
            act.setEnabled(False)
            self.menu_recent.addAction(act)
            return

        for fpath in self.recent_files:
             # Use path as text (checking if valid?)
             # Truncate if too long?
             fname = Path(fpath).name
             # Show name, toolip full path
             # Or show full path if ambiguous?
             # Let's show "Name (Path)" or just Path
             # Path is clearer for now
             act = QAction(fname, self)
             act.setToolTip(fpath)
             act.setData(fpath)
             # Use lambda with default arg to capture fpath properly
             act.triggered.connect(lambda checked=False, p=fpath: self.load_from_path(Path(p)))
             self.menu_recent.addAction(act)

        self.menu_recent.addSeparator()
        act_clear = QAction("Clear recent files", self)
        act_clear.triggered.connect(self._clear_recent)
        self.menu_recent.addAction(act_clear)

    def _clear_recent(self):
        self.recent_files = []
        self._update_recent_menu()
        self._save_config()

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
        self._save_config()
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

    # Capture directory from where script was run (before any chdir happens inside app)
    launch_cwd = Path.cwd()

    w = MainWindow(initial_file=initial_file, launch_cwd=launch_cwd)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
