"""
DOCX / XLSX to PDF Converter — GUI + CLI

Converts every .docx and .xlsx file in a directory (and subdirectories)
to PDF.

When a graphical display is available the app opens a tkinter GUI with a
directory browser, progress bar, and report table.  On headless servers
(no $DISPLAY) it falls back automatically to an interactive CLI mode.
You can also force CLI mode with the --cli flag.

Requirements:
    - Python 3.8+  (tkinter is included with the standard installer)
    - Windows: Microsoft Word & Excel installed + pip install comtypes
    - Linux/macOS: LibreOffice installed (sudo apt install libreoffice)
"""

import hashlib
import os
import shutil
import subprocess
import sys
import threading
from collections import defaultdict
from pathlib import Path


# ---------------------------------------------------------------------------
# Conversion logic
# ---------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = {".docx", ".xlsx"}


def find_convertible_files(root: Path):
    """Return a sorted list of .docx and .xlsx files, skipping temp files (~$)."""
    files = []
    for ext in SUPPORTED_EXTENSIONS:
        files.extend(
            f for f in root.rglob(f"*{ext}") if not f.name.startswith("~$")
        )
    return sorted(files)


def convert_one_word(docx_path: Path, word_app):
    """Convert a single .docx → .pdf using Word COM. Returns (pdf_path, error|None)."""
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        doc = word_app.Documents.Open(str(docx_path))
        doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        return pdf_path, None
    except Exception as exc:
        return None, str(exc)


def convert_one_excel(xlsx_path: Path, excel_app):
    """Convert a single .xlsx → .pdf using Excel COM. Returns (pdf_path, error|None)."""
    pdf_path = xlsx_path.with_suffix(".pdf")
    try:
        wb = excel_app.Workbooks.Open(str(xlsx_path))
        wb.ExportAsFixedFormat(0, str(pdf_path))  # 0 = xlTypePDF
        wb.Close(False)
        return pdf_path, None
    except Exception as exc:
        return None, str(exc)


def convert_one_libreoffice(file_path: Path):
    """Convert a single .docx/.xlsx → .pdf using LibreOffice. Returns (pdf_path, error|None)."""
    pdf_path = file_path.with_suffix(".pdf")
    try:
        result = subprocess.run(
            [
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", str(file_path.parent),
                str(file_path),
            ],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            return None, result.stderr.strip() or f"soffice exited with code {result.returncode}"
        if pdf_path.exists():
            return pdf_path, None
        return None, "PDF file was not created"
    except Exception as exc:
        return None, str(exc)


# ---------------------------------------------------------------------------
# Duplicate file detection / removal
# ---------------------------------------------------------------------------

def _file_hash(path: Path, buf_size: int = 65536) -> str:
    """Return the SHA-256 hex digest of a file."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            data = f.read(buf_size)
            if not data:
                break
            h.update(data)
    return h.hexdigest()


def find_duplicates(root: Path, progress_cb=None):
    """Scan *root* recursively and return duplicate groups.

    Returns a list of (kept_path, [duplicate_paths, ...]) tuples.
    *progress_cb*, if provided, is called with (files_scanned, total_files).
    """
    # 1. Collect every file and group by size (cheap pre-filter).
    all_files = sorted(f for f in root.rglob("*") if f.is_file() and not f.name.startswith("~$"))
    total = len(all_files)

    size_groups = defaultdict(list)
    for f in all_files:
        try:
            size_groups[f.stat().st_size].append(f)
        except OSError:
            pass

    # 2. For groups with the same size, compute hashes.
    hash_groups = defaultdict(list)
    scanned = 0
    for same_size in size_groups.values():
        for f in same_size:
            scanned += 1
            if progress_cb:
                progress_cb(scanned, total)
            if len(same_size) < 2:
                continue
            try:
                digest = _file_hash(f)
                hash_groups[digest].append(f)
            except OSError:
                pass

    # Count remaining files that were in unique-size groups.
    if progress_cb:
        progress_cb(total, total)

    # 3. Build result — keep the first path (sorted), rest are duplicates.
    results = []
    for paths in hash_groups.values():
        if len(paths) < 2:
            continue
        kept, *dupes = sorted(paths)
        results.append((kept, dupes))
    return results


# ---------------------------------------------------------------------------
# Display detection
# ---------------------------------------------------------------------------

def _display_available():
    """Return True if a graphical display is available for tkinter."""
    if sys.platform == "win32":
        return True  # Windows always has a desktop when logged in
    if not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY"):
        return False
    try:
        import tkinter as tk
        root = tk.Tk()
        root.destroy()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# CLI mode (headless)
# ---------------------------------------------------------------------------

def main_cli():
    """Interactive CLI mode for headless environments."""
    print("DOCX / XLSX to PDF Converter (CLI mode)\n")

    directory = input("Enter the directory path containing .docx / .xlsx files: ").strip()
    directory = directory.strip('"').strip("'")

    root = Path(directory)
    if not root.is_dir():
        print(f"Error: '{directory}' is not a valid directory.")
        sys.exit(1)

    files = find_convertible_files(root)
    if not files:
        print("No .docx or .xlsx files found.")
        sys.exit(0)

    total = len(files)
    print(f"Found {total} file(s).\n")

    use_com = sys.platform == "win32"

    # --- Platform-specific setup --------------------------------------------
    word = None
    excel = None
    if use_com:
        try:
            import comtypes
            import comtypes.client
        except ImportError:
            print("Missing dependency. Install it with:  pip install comtypes")
            sys.exit(1)
        comtypes.CoInitialize()
        has_docx = any(f.suffix.lower() == ".docx" for f in files)
        has_xlsx = any(f.suffix.lower() == ".xlsx" for f in files)
        try:
            if has_docx:
                word = comtypes.client.CreateObject("Word.Application")
                word.Visible = False
            if has_xlsx:
                excel = comtypes.client.CreateObject("Excel.Application")
                excel.Visible = False
        except Exception as exc:
            if word:
                word.Quit()
            if excel:
                excel.Quit()
            comtypes.CoUninitialize()
            print(f"Error: Could not start Office application: {exc}")
            sys.exit(1)
    else:
        if shutil.which("soffice") is None:
            print("Error: LibreOffice is not installed.")
            print("Install it with:  sudo apt install libreoffice")
            sys.exit(1)

    # --- Convert each file --------------------------------------------------
    converted = 0
    skipped = 0
    failed = 0

    try:
        for idx, src_file in enumerate(files, start=1):
            abs_path = src_file.resolve()
            pdf_path = abs_path.with_suffix(".pdf")

            try:
                rel = src_file.relative_to(root)
            except ValueError:
                rel = src_file

            if pdf_path.exists():
                skipped += 1
                print(f"  [{idx}/{total}] {rel} ... Skipped (PDF exists)")
            else:
                ext = abs_path.suffix.lower()
                if use_com:
                    if ext == ".docx":
                        pdf_path, error = convert_one_word(abs_path, word)
                    else:
                        pdf_path, error = convert_one_excel(abs_path, excel)
                else:
                    pdf_path, error = convert_one_libreoffice(abs_path)

                if error is None:
                    converted += 1
                    print(f"  [{idx}/{total}] {rel} ... Success")
                else:
                    failed += 1
                    print(f"  [{idx}/{total}] {rel} ... FAILED: {error}")
    finally:
        if use_com:
            if word:
                word.Quit()
            if excel:
                excel.Quit()
            comtypes.CoUninitialize()

    print(f"\nTotal: {total} | Converted: {converted} | Skipped: {skipped} | Failed: {failed}")

    # --- Duplicate removal --------------------------------------------------
    answer = input("\nScan for and remove exact duplicate files? [y/N]: ").strip().lower()
    if answer in ("y", "yes"):
        _cli_remove_duplicates(root)


def _cli_remove_duplicates(root: Path):
    """Scan a directory for duplicates and interactively remove them (CLI)."""
    print("\nScanning for duplicate files ...")
    duplicates = find_duplicates(root)

    if not duplicates:
        print("No duplicate files found.")
        return

    total_dupes = sum(len(dupes) for _, dupes in duplicates)
    print(f"Found {total_dupes} duplicate(s) across {len(duplicates)} group(s).\n")

    for kept, dupes in duplicates:
        try:
            rel_kept = kept.relative_to(root)
        except ValueError:
            rel_kept = kept
        print(f"  Keeping: {rel_kept}")
        for d in dupes:
            try:
                rel_d = d.relative_to(root)
            except ValueError:
                rel_d = d
            print(f"  Removing: {rel_d}")
        print()

    confirm = input(f"Delete {total_dupes} duplicate file(s)? [y/N]: ").strip().lower()
    if confirm not in ("y", "yes"):
        print("Aborted — no files were deleted.")
        return

    removed = 0
    for _, dupes in duplicates:
        for d in dupes:
            try:
                d.unlink()
                removed += 1
            except OSError as exc:
                print(f"  Could not delete {d}: {exc}")
    print(f"\nRemoved {removed} duplicate file(s).")


# ---------------------------------------------------------------------------
# GUI mode (tkinter)
# ---------------------------------------------------------------------------

def main_gui():
    """Launch the tkinter GUI."""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class ConverterApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("DOCX / XLSX to PDF Converter")
            self.minsize(780, 520)
            self._build_ui()
            self._converting = False

        # ---- UI construction -----------------------------------------------

        def _build_ui(self):
            pad = {"padx": 10, "pady": 5}

            # -- Directory row -----------------------------------------------
            dir_frame = ttk.Frame(self)
            dir_frame.pack(fill="x", **pad)

            ttk.Label(dir_frame, text="Directory:").pack(side="left")
            self.dir_var = tk.StringVar()
            self.dir_entry = ttk.Entry(dir_frame, textvariable=self.dir_var)
            self.dir_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
            ttk.Button(dir_frame, text="Browse", command=self._browse).pack(side="left")

            # -- Action buttons -------------------------------------------------
            btn_frame = ttk.Frame(self)
            btn_frame.pack(**pad)

            self.convert_btn = ttk.Button(btn_frame, text="Convert", command=self._start_conversion)
            self.convert_btn.pack(side="left", padx=(0, 5))

            self.dedup_btn = ttk.Button(btn_frame, text="Remove Duplicates", command=self._start_dedup)
            self.dedup_btn.pack(side="left")

            # -- Progress bar + counter --------------------------------------
            prog_frame = ttk.Frame(self)
            prog_frame.pack(fill="x", **pad)

            ttk.Label(prog_frame, text="Progress:").pack(side="left")
            self.progress = ttk.Progressbar(prog_frame, length=400, mode="determinate")
            self.progress.pack(side="left", fill="x", expand=True, padx=(5, 5))
            self.prog_label = ttk.Label(prog_frame, text="0 / 0")
            self.prog_label.pack(side="left")

            # -- Report table (Treeview) -------------------------------------
            table_frame = ttk.LabelFrame(self, text="Conversion Report")
            table_frame.pack(fill="both", expand=True, **pad)

            columns = ("num", "file", "status", "pdf_path")
            self.tree = ttk.Treeview(
                table_frame, columns=columns, show="headings", selectmode="browse"
            )
            self.tree.heading("num", text="#")
            self.tree.heading("file", text="File")
            self.tree.heading("status", text="Status")
            self.tree.heading("pdf_path", text="PDF Path")

            self.tree.column("num", width=40, stretch=False, anchor="center")
            self.tree.column("file", width=220, anchor="w")
            self.tree.column("status", width=140, anchor="w")
            self.tree.column("pdf_path", width=340, anchor="w")

            scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=scrollbar.set)
            self.tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # -- Summary label -----------------------------------------------
            self.summary_var = tk.StringVar(value="Total: 0 | Converted: 0 | Skipped: 0 | Failed: 0")
            ttk.Label(self, textvariable=self.summary_var, font=("", 10, "bold")).pack(**pad)

        # ---- Callbacks -----------------------------------------------------

        def _browse(self):
            folder = filedialog.askdirectory(title="Select folder with .docx / .xlsx files")
            if folder:
                self.dir_var.set(folder)

        def _start_conversion(self):
            if self._converting:
                return

            directory = self.dir_var.get().strip().strip('"').strip("'")
            if not directory:
                messagebox.showwarning("No directory", "Please select a directory first.")
                return

            root = Path(directory)
            if not root.is_dir():
                messagebox.showerror("Invalid directory", f"'{directory}' is not a valid directory.")
                return

            files = find_convertible_files(root)
            if not files:
                messagebox.showinfo("Nothing to do", "No .docx or .xlsx files found in that directory.")
                return

            # Reset UI
            self.tree.delete(*self.tree.get_children())
            total = len(files)
            self.progress["maximum"] = total
            self.progress["value"] = 0
            self.prog_label.config(text=f"0 / {total}")
            self.summary_var.set(f"Total: {total} | Converted: 0 | Skipped: 0 | Failed: 0")
            self.convert_btn.state(["disabled"])
            self.dedup_btn.state(["disabled"])
            self._converting = True

            thread = threading.Thread(
                target=self._convert_thread, args=(root, files), daemon=True
            )
            thread.start()

        # ---- Background conversion thread ----------------------------------

        def _convert_thread(self, root: Path, files: list):
            use_com = sys.platform == "win32"

            word = None
            excel = None
            if use_com:
                import comtypes
                import comtypes.client
                comtypes.CoInitialize()
                has_docx = any(f.suffix.lower() == ".docx" for f in files)
                has_xlsx = any(f.suffix.lower() == ".xlsx" for f in files)
                try:
                    if has_docx:
                        word = comtypes.client.CreateObject("Word.Application")
                        word.Visible = False
                    if has_xlsx:
                        excel = comtypes.client.CreateObject("Excel.Application")
                        excel.Visible = False
                except Exception as exc:
                    if word:
                        word.Quit()
                    if excel:
                        excel.Quit()
                    comtypes.CoUninitialize()
                    self.after(0, self._thread_error, f"Could not start Office application:\n{exc}")
                    return
            else:
                if shutil.which("soffice") is None:
                    self.after(
                        0, self._thread_error,
                        "LibreOffice is not installed.\n\n"
                        "Install it with:  sudo apt install libreoffice",
                    )
                    return

            total = len(files)
            converted = 0
            skipped = 0
            failed = 0

            try:
                for idx, src_file in enumerate(files, start=1):
                    abs_path = src_file.resolve()
                    pdf_path = abs_path.with_suffix(".pdf")

                    if pdf_path.exists():
                        skipped += 1
                        status = "Skipped — PDF exists"
                        pdf_display = str(pdf_path)
                    else:
                        ext = abs_path.suffix.lower()
                        if use_com:
                            if ext == ".docx":
                                pdf_path, error = convert_one_word(abs_path, word)
                            else:
                                pdf_path, error = convert_one_excel(abs_path, excel)
                        else:
                            pdf_path, error = convert_one_libreoffice(abs_path)

                        if error is None:
                            converted += 1
                            status = "Success"
                            pdf_display = str(pdf_path)
                        else:
                            failed += 1
                            status = f"Failed — {error}"
                            pdf_display = "—"

                    try:
                        rel = src_file.relative_to(root)
                    except ValueError:
                        rel = src_file

                    self.after(
                        0, self._update_row,
                        idx, str(rel), status, pdf_display,
                        idx, total, converted, skipped, failed,
                    )
            finally:
                if use_com:
                    if word:
                        word.Quit()
                    if excel:
                        excel.Quit()
                    comtypes.CoUninitialize()
                self.after(0, self._conversion_done)

        # ---- Duplicate removal -----------------------------------------------

        def _start_dedup(self):
            if self._converting:
                return

            directory = self.dir_var.get().strip().strip('"').strip("'")
            if not directory:
                messagebox.showwarning("No directory", "Please select a directory first.")
                return

            root = Path(directory)
            if not root.is_dir():
                messagebox.showerror("Invalid directory", f"'{directory}' is not a valid directory.")
                return

            # Reset UI
            self.tree.delete(*self.tree.get_children())
            self.progress["maximum"] = 1
            self.progress["value"] = 0
            self.prog_label.config(text="Scanning ...")
            self.summary_var.set("Scanning for duplicate files ...")
            self.convert_btn.state(["disabled"])
            self.dedup_btn.state(["disabled"])
            self._converting = True

            thread = threading.Thread(
                target=self._dedup_thread, args=(root,), daemon=True
            )
            thread.start()

        def _dedup_thread(self, root: Path):
            def on_progress(scanned, total):
                self.after(0, self._dedup_scan_progress, scanned, total)

            duplicates = find_duplicates(root, progress_cb=on_progress)

            if not duplicates:
                self.after(0, self._dedup_no_dupes)
                return

            # Show what will be removed, then delete.
            total_dupes = sum(len(dupes) for _, dupes in duplicates)
            row = 0
            removed = 0
            errors = 0

            for kept, dupes in duplicates:
                # Show the kept file.
                row += 1
                try:
                    rel_kept = kept.relative_to(root)
                except ValueError:
                    rel_kept = kept
                self.after(0, self._dedup_row, row, str(rel_kept), "Kept (original)", "")

                # Delete each duplicate.
                for d in dupes:
                    row += 1
                    try:
                        rel_d = d.relative_to(root)
                    except ValueError:
                        rel_d = d
                    try:
                        d.unlink()
                        removed += 1
                        self.after(0, self._dedup_row, row, str(rel_d), "Removed (duplicate)", "")
                    except OSError as exc:
                        errors += 1
                        self.after(0, self._dedup_row, row, str(rel_d), f"Error — {exc}", "")

            self.after(
                0, self._dedup_done,
                len(duplicates), total_dupes, removed, errors,
            )

        # ---- Thread-safe UI updates ----------------------------------------

        def _dedup_scan_progress(self, scanned, total):
            self.progress["maximum"] = total
            self.progress["value"] = scanned
            self.prog_label.config(text=f"{scanned} / {total}")

        def _dedup_row(self, num, filename, status, _detail):
            if status.startswith("Kept"):
                tag = "skipped"
            elif status.startswith("Removed"):
                tag = "success"
            else:
                tag = "fail"
            self.tree.insert("", "end", values=(num, filename, status, ""), tags=(tag,))
            self.tree.yview_moveto(1.0)

        def _dedup_no_dupes(self):
            self._converting = False
            self.convert_btn.state(["!disabled"])
            self.dedup_btn.state(["!disabled"])
            self.summary_var.set("No duplicate files found.")
            self.prog_label.config(text="Done")
            messagebox.showinfo("No duplicates", "No duplicate files found in that directory.")

        def _dedup_done(self, groups, total_dupes, removed, errors):
            self._converting = False
            self.convert_btn.state(["!disabled"])
            self.dedup_btn.state(["!disabled"])
            self.summary_var.set(
                f"Groups: {groups} | Duplicates: {total_dupes} | Removed: {removed} | Errors: {errors}"
            )
            self.prog_label.config(text="Done")

        def _update_row(self, num, filename, status, pdf_path,
                        current, total, converted, skipped, failed):
            if status == "Success":
                tag = "success"
            elif status.startswith("Skipped"):
                tag = "skipped"
            else:
                tag = "fail"
            self.tree.insert("", "end", values=(num, filename, status, pdf_path), tags=(tag,))
            self.tree.yview_moveto(1.0)
            self.progress["value"] = current
            self.prog_label.config(text=f"{current} / {total}")
            self.summary_var.set(
                f"Total: {total} | Converted: {converted} | Skipped: {skipped} | Failed: {failed}"
            )

        def _conversion_done(self):
            self._converting = False
            self.convert_btn.state(["!disabled"])
            self.dedup_btn.state(["!disabled"])

        def _thread_error(self, message):
            self._converting = False
            self.convert_btn.state(["!disabled"])
            self.dedup_btn.state(["!disabled"])
            messagebox.showerror("Error", message)

    app = ConverterApp()
    app.tree.tag_configure("success", foreground="green")
    app.tree.tag_configure("skipped", foreground="gray")
    app.tree.tag_configure("fail", foreground="red")
    app.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    force_cli = "--cli" in sys.argv

    if force_cli or not _display_available():
        main_cli()
    else:
        main_gui()


if __name__ == "__main__":
    main()
