#!/usr/bin/env python3
"""
Procurement Categorization — Desktop Launcher
Drag-and-drop or browse to select one or more CSV/Excel files,
then click Run to categorize and generate the Excel report.

No extra dependencies beyond what run_categorization.py already needs.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)


def _read_source_csv(path: str):
    """Read a source CSV, skipping phantom columns during parse.

    PO data CSVs have unescaped commas in description fields, so pandas
    detects thousands of phantom columns named 'Unnamed: N'.  We use
    usecols with a callable to reject them during parsing so they never
    get allocated in memory.
    """
    import pandas as pd

    # Only keep columns that are NOT auto-generated phantom names
    _keep = lambda c: not str(c).startswith("Unnamed:")

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False,
                               on_bad_lines="skip", usecols=_keep)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(path, encoding="latin-1", low_memory=False,
                       on_bad_lines="skip", usecols=_keep)


# ── Colour palette ────────────────────────────────────────────────────────────
BG       = "#1e1e2e"
SURFACE  = "#2a2a3e"
ACCENT   = "#6c63ff"
ACCENT2  = "#48cfad"
TEXT     = "#e0e0e0"
MUTED    = "#888"
SUCCESS  = "#48cfad"
WARNING  = "#f9ca24"
ERROR    = "#ff6b6b"
BORDER   = "#3a3a5c"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Procurement Categorization")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(700, 460)

        self.files: list[str] = []
        self.running = False

        self._build_ui()
        self._center()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Procurement Spend Categorization",
                 font=("Helvetica", 16, "bold"), bg=ACCENT, fg="white").pack()
        tk.Label(hdr, text="Load one or more CSV / Excel files and generate the strategic sourcing report",
                 font=("Helvetica", 9), bg=ACCENT, fg="#ddd").pack()

        # ── File list area ───────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG, padx=18, pady=12)
        body.pack(fill="both", expand=True)

        lbl_row = tk.Frame(body, bg=BG)
        lbl_row.pack(fill="x")
        tk.Label(lbl_row, text="Input Files", font=("Helvetica", 11, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Label(lbl_row, text="(CSV or Excel)", font=("Helvetica", 9),
                 bg=BG, fg=MUTED).pack(side="left", padx=6)

        # Listbox with scrollbar
        list_frame = tk.Frame(body, bg=BORDER, padx=1, pady=1)
        list_frame.pack(fill="both", expand=True, pady=(6, 0))

        self.listbox = tk.Listbox(
            list_frame, bg=SURFACE, fg=TEXT, selectbackground=ACCENT,
            selectforeground="white", font=("Courier", 9),
            activestyle="none", relief="flat", borderwidth=0,
            highlightthickness=0
        )
        scroll = tk.Scrollbar(list_frame, command=self.listbox.yview,
                              bg=SURFACE, troughcolor=SURFACE)
        self.listbox.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.listbox.pack(fill="both", expand=True)

        self._placeholder()

        # ── File buttons + Run button ─────────────────────────────────────────
        btn_row = tk.Frame(body, bg=BG, pady=6)
        btn_row.pack(fill="x")

        self._btn(btn_row, "＋  Add Files", self._add_files, ACCENT).pack(side="left")
        self._btn(btn_row, "✕  Remove Selected", self._remove_selected,
                  SURFACE, fg=ERROR).pack(side="left", padx=8)
        self._btn(btn_row, "Clear All", self._clear_all,
                  SURFACE, fg=MUTED).pack(side="left")

        # Run button — always visible, right-aligned on the same row
        self.run_btn = self._btn(btn_row, "▶  Run Categorization",
                                 self._run, ACCENT2, fg="#000",
                                 font=("Helvetica", 11, "bold"), padx=20, pady=6)
        self.run_btn.pack(side="right")

        # Excel-only button — skip categorization, just build Excel from existing CSV
        self.excel_btn = self._btn(btn_row, "📊  Excel from CSV",
                                   self._run_excel_only, ACCENT, fg="white",
                                   font=("Helvetica", 9), padx=12, pady=6)
        self.excel_btn.pack(side="right", padx=(0, 8))

        # ── Output folder ────────────────────────────────────────────────────
        out_row = tk.Frame(body, bg=BG, pady=4)
        out_row.pack(fill="x")
        tk.Label(out_row, text="Output Folder:", bg=BG, fg=TEXT,
                 font=("Helvetica", 9)).pack(side="left")

        self.out_var = tk.StringVar(value=script_dir)
        out_entry = tk.Entry(out_row, textvariable=self.out_var,
                             bg=SURFACE, fg=TEXT, insertbackground=TEXT,
                             relief="flat", font=("Courier", 9), width=52)
        out_entry.pack(side="left", padx=6, ipady=4)
        self._btn(out_row, "Browse", self._pick_output_dir,
                  SURFACE, fg=ACCENT2).pack(side="left")

        # ── Progress bar ─────────────────────────────────────────────────────
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=SURFACE, background=ACCENT2,
                        bordercolor=BORDER, lightcolor=ACCENT2, darkcolor=ACCENT2)

        self.progress = ttk.Progressbar(body, mode="indeterminate", length=400,
                                        style="Custom.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(6, 0))

        # ── Log console ──────────────────────────────────────────────────────
        tk.Label(body, text="Log", font=("Helvetica", 9, "bold"),
                 bg=BG, fg=MUTED).pack(anchor="w", pady=(6, 2))

        self.log = scrolledtext.ScrolledText(
            body, height=7, bg="#0d0d1a", fg=TEXT,
            font=("Courier", 9), relief="flat", state="disabled",
            insertbackground=TEXT
        )
        self.log.pack(fill="both", expand=False)
        self.log.tag_config("ok",   foreground=SUCCESS)
        self.log.tag_config("warn", foreground=WARNING)
        self.log.tag_config("err",  foreground=ERROR)
        self.log.tag_config("hdr",  foreground=ACCENT2, font=("Courier", 9, "bold"))

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _btn(self, parent, text, cmd, bg, fg=TEXT, font=("Helvetica", 9), **kw):
        return tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                         activebackground=bg, activeforeground=fg,
                         relief="flat", cursor="hand2", font=font,
                         padx=kw.pop("padx", 12), pady=kw.pop("pady", 5), **kw)

    def _center(self):
        self.update_idletasks()
        w, h = 720, 500
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _placeholder(self):
        if not self.files:
            self.listbox.configure(fg=MUTED)
            self.listbox.delete(0, "end")
            self.listbox.insert("end", "  No files added yet — click '＋ Add Files' to browse")

    def _refresh_list(self):
        self.listbox.delete(0, "end")
        if not self.files:
            self._placeholder()
            return
        self.listbox.configure(fg=TEXT)
        for i, f in enumerate(self.files, 1):
            size_kb = os.path.getsize(f) / 1024
            label = f"  {i:>2}.  {os.path.basename(f)}   ({size_kb:,.0f} KB)   {os.path.dirname(f)}"
            self.listbox.insert("end", label)

    # ── File management ───────────────────────────────────────────────────────

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select spend data files",
            filetypes=[("Spreadsheets", "*.csv *.xlsx *.xls"),
                       ("CSV files", "*.csv"),
                       ("Excel files", "*.xlsx *.xls"),
                       ("All files", "*.*")]
        )
        added = 0
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                added += 1
        if added:
            self._refresh_list()
            self._log(f"Added {added} file(s). Total: {len(self.files)}", "ok")

    def _remove_selected(self):
        selected = self.listbox.curselection()
        if not selected or not self.files:
            return
        for i in reversed(selected):
            if i < len(self.files):
                removed = self.files.pop(i)
                self._log(f"Removed: {os.path.basename(removed)}", "warn")
        self._refresh_list()

    def _clear_all(self):
        self.files.clear()
        self._refresh_list()
        self._log("File list cleared.", "warn")

    def _pick_output_dir(self):
        d = filedialog.askdirectory(title="Select output folder",
                                    initialdir=self.out_var.get())
        if d:
            self.out_var.set(d)

    # ── Logging ───────────────────────────────────────────────────────────────

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── Run ───────────────────────────────────────────────────────────────────

    def _run(self):
        if self.running:
            return
        if not self.files:
            self._log("No files selected. Add at least one CSV or Excel file.", "err")
            return

        self.running = True
        self.run_btn.configure(text="⏳  Running…", state="disabled", bg=MUTED)
        self.progress.start(12)
        self._log("─" * 60, "hdr")
        self._log(f"Starting categorization of {len(self.files)} file(s)…", "hdr")

        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            import pandas as pd
            from categorization import categorize_dataframe

            out_dir = self.out_var.get() or script_dir

            # ── Read all files ──────────────────────────────────────────────
            dfs = []
            for f in self.files:
                self._log(f"Reading: {os.path.basename(f)}")
                try:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in (".xlsx", ".xls"):
                        df = pd.read_excel(f)
                    else:
                        df = _read_source_csv(f)
                    self._log(f"  → {len(df):,} rows, {len(df.columns)} cols", "ok")
                    dfs.append(df)
                except Exception as e:
                    self._log(f"  ✗ Failed to read {os.path.basename(f)}: {e}", "err")

            if not dfs:
                self._log("No data could be read. Aborting.", "err")
                return

            combined = pd.concat(dfs, ignore_index=True)
            self._log(f"Combined: {len(combined):,} total rows across {len(dfs)} file(s)", "ok")

            # ── Categorize ──────────────────────────────────────────────────
            self._log("Categorizing…")
            result = categorize_dataframe(combined)

            # Bucket summary
            summary = result["master_bucket"].value_counts()
            self._log("─" * 40, "hdr")
            for bucket, count in summary.items():
                pct = count / len(result) * 100
                self._log(f"  {bucket:<42} {count:>6,}  ({pct:.1f}%)")

            unc = (result["master_bucket"] == "Uncategorized").sum()
            if unc:
                self._log(f"  ⚠  {unc:,} rows uncategorized", "warn")
            else:
                self._log("  ✓  All rows categorized", "ok")

            avg_conf = result["confidence_score"].mean()
            self._log(f"  Avg confidence: {avg_conf:.0%}", "ok")

            # ── Write CSV (only real columns, not phantom ones) ──────────
            csv_path = os.path.join(out_dir, "categorized_output.csv")
            result.to_csv(csv_path, index=False, encoding="utf-8-sig")
            self._log(f"CSV saved: {csv_path}", "ok")

            # ── Build Excel directly from in-memory DataFrame ──────────
            xlsx_path = os.path.join(out_dir, "Procurement_Detail_Breakdown.xlsx")
            self._log("Building Excel report…")
            from build_detail_excel_v2 import build_excel_from_df
            build_excel_from_df(result, xlsx_path, log=self._log)
            self._log(f"Excel saved: {xlsx_path}", "ok")
            self._log("Done! Open Procurement_Detail_Breakdown.xlsx to view results.", "ok")
            self._log("─" * 60, "hdr")

        except Exception as e:
            import traceback
            self._log(f"ERROR: {e}", "err")
            self._log(traceback.format_exc(), "err")
        finally:
            self.after(0, self._done)

    # ── Excel-only mode ────────────────────────────────────────────────────

    def _run_excel_only(self):
        if self.running:
            return
        csv_path = filedialog.askopenfilename(
            title="Select categorized CSV file",
            initialdir=self.out_var.get() or script_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not csv_path:
            return

        self.running = True
        self.run_btn.configure(state="disabled", bg=MUTED)
        self.excel_btn.configure(text="⏳  Building…", state="disabled", bg=MUTED)
        self.progress.start(12)
        self._log("─" * 60, "hdr")
        self._log(f"Building Excel from: {os.path.basename(csv_path)}", "hdr")

        threading.Thread(target=self._excel_worker, args=(csv_path,), daemon=True).start()

    def _excel_worker(self, csv_path):
        try:
            import pandas as pd
            from build_detail_excel_v2 import build_excel_from_df, _read_csv_robust

            out_dir = self.out_var.get() or script_dir
            xlsx_path = os.path.join(out_dir, "Procurement_Detail_Breakdown.xlsx")

            self._log("Reading categorized data…")
            df = _read_csv_robust(csv_path)
            build_excel_from_df(df, xlsx_path, log=self._log)
            self._log(f"Excel saved: {xlsx_path}", "ok")
            self._log("Done! Open Procurement_Detail_Breakdown.xlsx to view results.", "ok")
            self._log("─" * 60, "hdr")
        except Exception as e:
            import traceback
            self._log(f"ERROR: {e}", "err")
            self._log(traceback.format_exc(), "err")
        finally:
            self.after(0, self._done)

    def _done(self):
        self.running = False
        self.progress.stop()
        self.run_btn.configure(text="▶  Run Categorization",
                               state="normal", bg=ACCENT2)
        self.excel_btn.configure(text="📊  Excel from CSV",
                                 state="normal", bg=ACCENT)


if __name__ == "__main__":
    app = App()
    app.mainloop()
