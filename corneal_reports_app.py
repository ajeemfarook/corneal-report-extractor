"""
Corneal Reports PDF Extractor — GUI wrapper.

A small Tkinter front-end around `corneal_reports_core.py` that lets the user
pick the input/output folders and run the extraction with one click.

Run directly:
    python corneal_reports_app.py

Or build a single-file Windows .exe (see build_windows.bat):
    pyinstaller --noconfirm corneal_reports_app.spec
"""

from __future__ import annotations

import io
import os
import sys
import threading
import traceback
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Make sure we can import the core module whether we're run as a script,
# imported as a module, or frozen by PyInstaller.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
    sys.path.insert(0, str(BASE_DIR))
else:
    BASE_DIR = Path(__file__).resolve().parent
    sys.path.insert(0, str(BASE_DIR))

import corneal_reports_core as core  # noqa: E402


APP_TITLE = "Corneal Reports — PDF Extractor"


# ============================================================
# LOG CAPTURE
# ============================================================

class TextRedirector:
    """Redirect writes to a Tk Text widget, thread-safely."""

    def __init__(self, text_widget: tk.Text):
        self.text_widget = text_widget

    def write(self, message: str) -> None:
        if not message:
            return
        def _append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", message)
            self.text_widget.see("end")
            self.text_widget.configure(state="disabled")
        try:
            self.text_widget.after(0, _append)
        except Exception:
            # Window may already be closed during shutdown.
            pass

    def flush(self) -> None:  # noqa: D401
        pass


# ============================================================
# APP
# ============================================================

class App(tk.Tk):

    def __init__(self) -> None:
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("780x600")
        self.minsize(680, 520)

        # Default paths live next to the .exe (or this script, in dev).
        default_in = BASE_DIR / "reports"
        default_out = BASE_DIR / "output"

        self.input_dir = tk.StringVar(value=str(default_in))
        self.output_dir = tk.StringVar(value=str(default_out))
        self.status_text = tk.StringVar(value="Ready.")
        self.progress_value = tk.DoubleVar(value=0.0)

        self._build_ui()
        self._build_menu()

        # Window icon (best effort).
        try:
            self.iconbitmap(default=BASE_DIR / "app.ico")
        except Exception:
            pass

        # Center the window on screen.
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        # On close, ensure any running worker is asked to stop.
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._worker: threading.Thread | None = None

    # --------------------------------------------------------
    # UI
    # --------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        # ---- Input folder row ----
        ttk.Label(frm, text="Reports folder (PDFs):").grid(
            row=0, column=0, sticky="w", **pad
        )
        ttk.Entry(frm, textvariable=self.input_dir).grid(
            row=1, column=0, columnspan=2, sticky="ew", **pad
        )
        ttk.Button(frm, text="Browse…", command=self._pick_input).grid(
            row=1, column=2, **pad
        )

        # ---- Output folder row ----
        ttk.Label(frm, text="Output folder:").grid(
            row=2, column=0, sticky="w", **pad
        )
        ttk.Entry(frm, textvariable=self.output_dir).grid(
            row=3, column=0, columnspan=2, sticky="ew", **pad
        )
        ttk.Button(frm, text="Browse…", command=self._pick_output).grid(
            row=3, column=2, **pad
        )

        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)

        # ---- Action buttons ----
        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        btns.columnconfigure(2, weight=1)

        self.run_btn = ttk.Button(
            btns, text="▶  Process PDFs", command=self._on_run
        )
        self.run_btn.grid(row=0, column=0, sticky="ew", padx=4, pady=4)

        self.open_out_btn = ttk.Button(
            btns, text="Open output folder",
            command=self._open_output,
        )
        self.open_out_btn.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Button(
            btns, text="Clear log", command=self._clear_log
        ).grid(row=0, column=2, sticky="ew", padx=4, pady=4)

        # ---- Progress bar ----
        self.progress = ttk.Progressbar(
            frm, mode="indeterminate", variable=self.progress_value
        )
        self.progress.grid(
            row=5, column=0, columnspan=3, sticky="ew", **pad
        )

        # ---- Log area ----
        ttk.Label(frm, text="Log:").grid(
            row=6, column=0, sticky="w", **pad
        )
        log_frame = ttk.Frame(frm)
        log_frame.grid(
            row=7, column=0, columnspan=3, sticky="nsew", **pad
        )
        frm.rowconfigure(7, weight=1)
        frm.columnconfigure(0, weight=1)

        self.log = tk.Text(
            log_frame, wrap="word", state="disabled",
            font=("Consolas", 9), bg="#111", fg="#eee",
        )
        scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log.yview
        )
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # ---- Status bar ----
        status = ttk.Frame(self)
        status.pack(fill="x", side="bottom")
        ttk.Label(
            status, textvariable=self.status_text, anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=10, pady=4)

    def _build_menu(self) -> None:
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(
            label="Open output folder", command=self._open_output
        )
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    # --------------------------------------------------------
    # Picker actions
    # --------------------------------------------------------

    def _pick_input(self) -> None:
        path = filedialog.askdirectory(
            title="Select folder containing PDF reports",
            initialdir=self._safe_initial(self.input_dir.get()),
        )
        if path:
            self.input_dir.set(path)

    def _pick_output(self) -> None:
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self._safe_initial(self.output_dir.get()),
        )
        if path:
            self.output_dir.set(path)

    def _safe_initial(self, p: str) -> str:
        try:
            pp = Path(p)
            return pp if pp.exists() else str(BASE_DIR)
        except Exception:
            return str(BASE_DIR)

    # --------------------------------------------------------
    # Run / log / open
    # --------------------------------------------------------

    def _on_run(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo(
                APP_TITLE, "A run is already in progress. Please wait."
            )
            return

        in_path = Path(self.input_dir.get()).expanduser()
        out_path = Path(self.output_dir.get()).expanduser()

        if not in_path.exists():
            messagebox.showerror(
                APP_TITLE, f"Input folder does not exist:\n{in_path}"
            )
            return

        try:
            out_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror(
                APP_TITLE, f"Cannot create output folder:\n{e}"
            )
            return

        # Wire the core module to the chosen folders.
        core.INPUT_DIR = in_path
        core.OUTPUT_DIR = out_path
        core.OUTPUT_XLSX = out_path / "corneal_reports.xlsx"
        core.OUTPUT_CSV = out_path / "corneal_reports.csv"
        core.ERROR_CSV = out_path / "corneal_reports_errors.csv"

        self.run_btn.configure(state="disabled")
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.status_text.set("Processing…")

        self._worker = threading.Thread(
            target=self._run_core, daemon=True
        )
        self._worker.start()

    def _run_core(self) -> None:
        out = TextRedirector(self.log)
        err = TextRedirector(self.log)
        try:
            with redirect_stdout(out), redirect_stderr(err):
                core.main()
            self._finish(success=True)
        except Exception:
            tb = traceback.format_exc()
            err.write("\n" + tb + "\n")
            self._finish(success=False, message=traceback.format_exc())
        finally:
            try:
                sys.stdout = sys.__stdout__
                sys.stderr = sys.__stderr__
            except Exception:
                pass

    def _finish(self, success: bool, message: str = "") -> None:
        def _update():
            self.progress.stop()
            self.progress.configure(mode="determinate", value=100)
            self.run_btn.configure(state="normal")
            if success:
                self.status_text.set("Done. Output saved to the chosen folder.")
                self._log_builtin(
                    f"\n[GUI] Finished. Files saved under:\n"
                    f"      {self.output_dir.get()}\n"
                )
            else:
                self.status_text.set("Failed — see log.")
                self._log_builtin(f"\n[GUI] FAILED.\n{message}\n")
        self.after(0, _update)

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _log_builtin(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _open_output(self) -> None:
        out_path = self.output_dir.get()
        try:
            if sys.platform.startswith("win"):
                os.startfile(out_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{out_path}"')
            else:
                os.system(f'xdg-open "{out_path}"')
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Cannot open folder:\n{e}")

    def _show_about(self) -> None:
        messagebox.showinfo(
            APP_TITLE,
            "Corneal Reports PDF Extractor\n\n"
            "Reads PDFs from the reports folder, extracts clinical fields "
            "(Sim-K, cornea, pupil, AK, keratoconus, axial length, device), "
            "and writes Excel/CSV outputs.\n\n"
            "Built with PyMuPDF + pandas.",
        )

    def _on_close(self) -> None:
        if self._worker and self._worker.is_alive():
            if not messagebox.askyesno(
                APP_TITLE, "A run is still in progress. Quit anyway?"
            ):
                return
        self.destroy()


# ============================================================
# ENTRY POINT
# ============================================================

def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
