"""
DOCX-to-PDF Converter — GUI Edition

Browse to a folder (including mapped Google Drive letters), convert every
.docx file to PDF, and view a live progress bar plus a detailed report table.

Requirements:
    - Python 3.8+  (tkinter is included with the standard installer)
    - Windows: Microsoft Word installed + pip install comtypes
    - Linux/macOS: LibreOffice installed (sudo apt install libreoffice)
"""

import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ---------------------------------------------------------------------------
# Conversion logic (runs on a background thread)
# ---------------------------------------------------------------------------

def find_docx_files(root: Path):
    """Return a sorted list of .docx files, skipping Word temp files (~$)."""
    return sorted(
        f for f in root.rglob("*.docx") if not f.name.startswith("~$")
    )


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


def convert_one_libreoffice(docx_path: Path):
    """Convert a single .docx → .pdf using LibreOffice. Returns (pdf_path, error|None)."""
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        result = subprocess.run(
            [
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", str(docx_path.parent),
                str(docx_path),
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
# Main application window
# ---------------------------------------------------------------------------

class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOCX to PDF Converter")
        self.minsize(780, 520)
        self._build_ui()
        self._converting = False

    # ---- UI construction ---------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # -- Directory row ---------------------------------------------------
        dir_frame = ttk.Frame(self)
        dir_frame.pack(fill="x", **pad)

        ttk.Label(dir_frame, text="Directory:").pack(side="left")
        self.dir_var = tk.StringVar()
        self.dir_entry = ttk.Entry(dir_frame, textvariable=self.dir_var)
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        ttk.Button(dir_frame, text="Browse", command=self._browse).pack(side="left")

        # -- Convert button --------------------------------------------------
        self.convert_btn = ttk.Button(self, text="Convert", command=self._start_conversion)
        self.convert_btn.pack(**pad)

        # -- Progress bar + counter ------------------------------------------
        prog_frame = ttk.Frame(self)
        prog_frame.pack(fill="x", **pad)

        ttk.Label(prog_frame, text="Progress:").pack(side="left")
        self.progress = ttk.Progressbar(prog_frame, length=400, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(5, 5))
        self.prog_label = ttk.Label(prog_frame, text="0 / 0")
        self.prog_label.pack(side="left")

        # -- Report table (Treeview) -----------------------------------------
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

        # -- Summary label ---------------------------------------------------
        self.summary_var = tk.StringVar(value="Total: 0 | Converted: 0 | Skipped: 0 | Failed: 0")
        ttk.Label(self, textvariable=self.summary_var, font=("", 10, "bold")).pack(**pad)

    # ---- Callbacks ---------------------------------------------------------

    def _browse(self):
        folder = filedialog.askdirectory(title="Select folder with .docx files")
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

        docx_files = find_docx_files(root)
        if not docx_files:
            messagebox.showinfo("Nothing to do", "No .docx files found in that directory.")
            return

        # Reset UI
        self.tree.delete(*self.tree.get_children())
        total = len(docx_files)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        self.prog_label.config(text=f"0 / {total}")
        self.summary_var.set(f"Total: {total} | Converted: 0 | Skipped: 0 | Failed: 0")
        self.convert_btn.state(["disabled"])
        self._converting = True

        # Launch background thread
        thread = threading.Thread(
            target=self._convert_thread, args=(root, docx_files), daemon=True
        )
        thread.start()

    # ---- Background conversion thread --------------------------------------

    def _convert_thread(self, root: Path, docx_files: list):
        use_word = sys.platform == "win32"

        # --- Platform-specific setup ----------------------------------------
        word = None
        if use_word:
            import comtypes
            import comtypes.client
            comtypes.CoInitialize()
            try:
                word = comtypes.client.CreateObject("Word.Application")
                word.Visible = False
            except Exception as exc:
                comtypes.CoUninitialize()
                self.after(0, self._thread_error, f"Could not start Word:\n{exc}")
                return
        else:
            if shutil.which("soffice") is None:
                self.after(
                    0, self._thread_error,
                    "LibreOffice is not installed.\n\n"
                    "Install it with:  sudo apt install libreoffice",
                )
                return

        # --- Convert each file ----------------------------------------------
        total = len(docx_files)
        converted = 0
        skipped = 0
        failed = 0

        try:
            for idx, docx_file in enumerate(docx_files, start=1):
                abs_path = docx_file.resolve()
                pdf_path = abs_path.with_suffix(".pdf")

                # Skip if a PDF already exists
                if pdf_path.exists():
                    skipped += 1
                    status = "Skipped — PDF exists"
                    pdf_display = str(pdf_path)
                else:
                    if use_word:
                        pdf_path, error = convert_one_word(abs_path, word)
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

                # Compute path relative to the chosen root for display
                try:
                    rel = docx_file.relative_to(root)
                except ValueError:
                    rel = docx_file

                # Schedule UI update on the main thread
                self.after(
                    0, self._update_row,
                    idx, str(rel), status, pdf_display,
                    idx, total, converted, skipped, failed,
                )
        finally:
            if use_word:
                word.Quit()
                comtypes.CoUninitialize()
            self.after(0, self._conversion_done)

    # ---- Thread-safe UI updates --------------------------------------------

    def _update_row(self, num, filename, status, pdf_path,
                    current, total, converted, skipped, failed):
        if status == "Success":
            tag = "success"
        elif status.startswith("Skipped"):
            tag = "skipped"
        else:
            tag = "fail"
        self.tree.insert("", "end", values=(num, filename, status, pdf_path), tags=(tag,))
        self.tree.yview_moveto(1.0)  # auto-scroll to bottom
        self.progress["value"] = current
        self.prog_label.config(text=f"{current} / {total}")
        self.summary_var.set(
            f"Total: {total} | Converted: {converted} | Skipped: {skipped} | Failed: {failed}"
        )

    def _conversion_done(self):
        self._converting = False
        self.convert_btn.state(["!disabled"])

    def _thread_error(self, message):
        self._converting = False
        self.convert_btn.state(["!disabled"])
        messagebox.showerror("Error", message)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = ConverterApp()

    # Colour-code rows
    app.tree.tag_configure("success", foreground="green")
    app.tree.tag_configure("skipped", foreground="gray")
    app.tree.tag_configure("fail", foreground="red")

    app.mainloop()


if __name__ == "__main__":
    main()
