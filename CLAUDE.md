# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Single-file Python application (`convert_docx_to_pdf.py`) that batch-converts `.docx` and `.xlsx` files to PDF. Supports both a tkinter GUI (when a display is available) and an interactive CLI mode (headless or `--cli` flag). Also includes duplicate file detection and removal.

## Platform Behavior

- **Windows**: Uses COM automation via `comtypes` (Word and Excel must be installed)
- **Linux/macOS**: Uses LibreOffice headless mode (`soffice` must be in PATH)
- Display detection (`_display_available()`) auto-selects GUI vs CLI mode

## Running the App

```bash
python convert_docx_to_pdf.py         # GUI if display available, else CLI
python convert_docx_to_pdf.py --cli   # Force CLI mode
```

## Running Tests

```bash
pytest                          # Run all tests
pytest tests/test_find_files.py # Run a single test file
pytest -k test_success          # Run tests matching a name pattern
```

Tests live in `tests/` and use `pytest` with fixtures defined in `tests/conftest.py`. Three shared fixtures: `sample_tree` (directory with .docx/.xlsx files), `duplicate_tree` (files with known duplicates), and `empty_dir`. Tests mock COM objects and LibreOffice subprocess calls — no actual Office applications are needed to run the test suite.

## Architecture

All application code is in the single `convert_docx_to_pdf.py` module, organized into sections:

1. **Conversion logic** — `find_convertible_files()`, `convert_one_word()`, `convert_one_excel()`, `convert_one_libreoffice()`
2. **Duplicate detection** — `_file_hash()`, `find_duplicates()` (size pre-filter then SHA-256)
3. **Display detection** — `_display_available()`
4. **CLI mode** — `main_cli()`, `_cli_remove_duplicates()`
5. **GUI mode** — `main_gui()` containing the `ConverterApp` tkinter class (uses background threads with `self.after()` for thread-safe UI updates)
6. **Entry point** — `main()` dispatches to GUI or CLI

Each `convert_one_*` function returns a `(pdf_path, error|None)` tuple. Conversion skips files that already have a corresponding `.pdf` and ignores temp files (`~$` prefix).
