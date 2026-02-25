"""
DOCX-to-PDF Converter — GUI Edition

Browse to a folder (including mapped Google Drive letters), convert every
.docx file to PDF using Microsoft Word COM automation, and view a live
progress bar plus a detailed report table.

Requirements:
    - Windows with Microsoft Word installed
    - pip install comtypes
    - Python 3.8+  (tkinter is included with the standard installer)
"""

import os
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


def convert_one(docx_path: Path, word_app):
    """Convert a single .docx → .pdf. Returns (pdf_path, error_string|None)."""
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        doc = word_app.Documents.Open(str(docx_path))
        doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        return pdf_path, None
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
        self.summary_var = tk.StringVar(value="Total: 0 | Converted: 0 | Failed: 0")
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
        self.summary_var.set(f"Total: {total} | Converted: 0 | Failed: 0")
        self.convert_btn.state(["disabled"])
        self._converting = True

        # Launch background thread
        thread = threading.Thread(
            target=self._convert_thread, args=(root, docx_files), daemon=True
        )
        thread.start()

    # ---- Background conversion thread --------------------------------------

    def _convert_thread(self, root: Path, docx_files: list):
        # COM must be initialised on the thread that uses it.
        import comtypes
        import comtypes.client
        comtypes.CoInitialize()

        try:
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
        except Exception as exc:
            self.after(0, self._thread_error, f"Could not start Word:\n{exc}")
            return

        total = len(docx_files)
        converted = 0
        failed = 0

        try:
            for idx, docx_file in enumerate(docx_files, start=1):
                abs_path = docx_file.resolve()
                pdf_path, error = convert_one(abs_path, word)

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
                    idx, total, converted, failed,
                )
        finally:
            word.Quit()
            comtypes.CoUninitialize()
            self.after(0, self._conversion_done)

    # ---- Thread-safe UI updates --------------------------------------------

    def _update_row(self, num, filename, status, pdf_path,
                    current, total, converted, failed):
        tag = "success" if status == "Success" else "fail"
        self.tree.insert("", "end", values=(num, filename, status, pdf_path), tags=(tag,))
        self.tree.yview_moveto(1.0)  # auto-scroll to bottom
        self.progress["value"] = current
        self.prog_label.config(text=f"{current} / {total}")
        self.summary_var.set(
            f"Total: {total} | Converted: {converted} | Failed: {failed}"
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
    app.tree.tag_configure("fail", foreground="red")

    app.mainloop()


if __name__ == "__main__":
    main()
