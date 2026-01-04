# Marpit Slide Editor

A single-file GUI editor for Marpit / Marp slide decks, built with Python and PySide6.

## Features
- **Deck Overview**: Vertical list showing all slides.
- **Slide Editor**: Edit one slide at a time with formatting buttons.
- **Live Preview**: Real-time rendering using Marp CLI and Qt WebEngine (defaults to 1/3 window width).
- **Spell Check**: Automatic English spell checking (misspelled words underlined in red).

## Dependencies

### Python Dependencies
The project requires **Python 3.9+**.

1. Create a virtual environment:
   ```bash
   python3 -m venv venv
   ```
2. Activate the environment:
   ```bash
   source venv/bin/activate
   ```
3. Install dependencies (PySide6, pyspellchecker):
   ```bash
   pip install -r requirements.txt
   ```

### Marp CLI
The application requires the **Marp CLI** to render previews.

- **Automatic (Bundled)**: The application can use a local binary if placed in `bin/marp`.
  - *Note: If you followed the setup, a script may have already downloaded this for you.*
- **System-wide**: Alternatively, install it via npm:
  ```bash
  npm install -g @marp-team/marp-cli
  ```
  Or set the `MARP_CLI` environment variable to the path of your marp executable.

## Usage

Run the editor from your virtual environment:

```bash
./venv/bin/python marpit_slide_editor.py
```

You can also open a specific file directly:

```bash
./venv/bin/python marpit_slide_editor.py path/to/presentation.md
```