from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .generator import generate
from .settings import Settings
from .workbook import (ValidationError, WorkbookError, discover_sheets, export_issues,
                       open_xlsx, validate_reference)

BG = "#080B14"
SURFACE = "#101625"
TEXT = "#F4F6FF"
MUTED = "#9AA7BC"
PRIMARY = "#7C5CFF"


class Application:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.settings = Settings.load()
        self.events = queue.Queue()
        self.busy = False
        self.result = None
        self.raw_path = tk.StringVar()
        self.sheet = tk.StringVar()
        self.status = tk.StringVar(value="Choose your customer open order export to begin.")
        self.summary = tk.StringVar(value="Your completed report will appear here.")
        self.reference_label = tk.StringVar()
        self.folder_label = tk.StringVar()
        self.update_settings_labels()
        root.title("MTC Open Order")
        root.geometry("1000x700")
        root.minsize(860, 680)
        root.configure(bg=BG)
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.configure_styles()
        self.build_ui()
        root.after(100, self.poll)

    def configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=SURFACE)
        style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=SURFACE)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED)
        style.configure("CardMuted.TLabel", background=SURFACE, foreground=MUTED)
        style.configure("Title.TLabel", font=("Segoe UI", 26, "bold"))
        style.configure("Heading.TLabel", background=SURFACE, font=("Segoe UI", 13, "bold"))
        style.configure("TButton", font=("Segoe UI", 10), padding=(16, 9),
                        background="#222B3E", foreground=TEXT, borderwidth=0)
        style.map("TButton", background=[("active", "#35425D"), ("disabled", "#192030")],
                  foreground=[("disabled", "#69778D")])
        style.configure("Primary.TButton", background=PRIMARY, font=("Segoe UI", 11, "bold"))
        style.map("Primary.TButton", background=[("active", "#9277FF"), ("disabled", "#30284F")])
        style.configure("TEntry", fieldbackground="#1B2435", foreground=TEXT, padding=8)
        style.configure("TCombobox", fieldbackground="#1B2435", foreground=TEXT, padding=6)
        style.map("TCombobox", fieldbackground=[("readonly", "#1B2435")],
                  foreground=[("readonly", TEXT)])
        style.configure("Horizontal.TProgressbar", troughcolor="#222B3E", background=PRIMARY,
                        borderwidth=0, thickness=7)
        style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE,
                        foreground=TEXT, rowheight=28, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background="#222B3E", foreground=TEXT)

    def build_ui(self):
        page = ttk.Frame(self.root, padding=30)
        page.pack(fill="both", expand=True)
        header = ttk.Frame(page)
        header.pack(fill="x")
        ttk.Label(header, text="MTC Open Order", style="Title.TLabel").pack(side="left")
        self.settings_button = ttk.Button(header, text="Settings", command=self.open_settings)
        self.settings_button.pack(side="right")
        ttk.Label(page, text="Turn your raw export into a ready-to-use cluster report.",
                  style="Muted.TLabel").pack(anchor="w", pady=(5, 24))

        card = ttk.Frame(page, style="Card.TFrame", padding=22)
        card.pack(fill="x")
        ttk.Label(card, text="1   Select raw Excel", style="Heading.TLabel").pack(anchor="w")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", pady=(14, 10))
        self.path_entry = ttk.Entry(row, textvariable=self.raw_path, state="readonly")
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 12))
        self.browse_button = ttk.Button(row, text="Browse…", command=self.browse)
        self.browse_button.pack(side="right")
        sheets = ttk.Frame(card, style="Card.TFrame")
        sheets.pack(fill="x")
        ttk.Label(sheets, text="Worksheet", style="CardMuted.TLabel").pack(side="left", padx=(0, 12))
        self.sheet_picker = ttk.Combobox(sheets, textvariable=self.sheet, state="disabled", width=36)
        self.sheet_picker.pack(side="left")
        ttk.Label(card, textvariable=self.reference_label, style="CardMuted.TLabel",
                  wraplength=830).pack(anchor="w", pady=(14, 0))

        actions = ttk.Frame(page)
        actions.pack(fill="x", pady=20)
        self.generate_button = ttk.Button(actions, text="2   Generate Output", style="Primary.TButton",
                                          command=self.start_generation, state="disabled")
        self.generate_button.pack(side="left")
        ttk.Label(actions, text="Uses Microsoft Excel on this PC", style="Muted.TLabel").pack(side="left", padx=18)
        self.progress = ttk.Progressbar(page, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        ttk.Label(page, textvariable=self.status, style="Muted.TLabel", wraplength=880).pack(anchor="w", pady=(10, 20))

        result_card = ttk.Frame(page, style="Card.TFrame", padding=22)
        result_card.pack(fill="both", expand=True)
        ttk.Label(result_card, text="3   Your generated report", style="Heading.TLabel").pack(anchor="w")
        ttk.Label(result_card, textvariable=self.summary, style="Card.TLabel", wraplength=820,
                  justify="left").pack(anchor="w", pady=(12, 14))
        buttons = ttk.Frame(result_card, style="Card.TFrame")
        buttons.pack(anchor="w")
        self.open_file_button = ttk.Button(buttons, text="Open File", command=self.open_file, state="disabled")
        self.open_file_button.pack(side="left", padx=(0, 10))
        self.open_folder_button = ttk.Button(buttons, text="Open Folder", command=self.open_folder, state="disabled")
        self.open_folder_button.pack(side="left")
        ttk.Label(page, textvariable=self.folder_label, style="Muted.TLabel",
                  wraplength=880).pack(anchor="w", pady=(16, 0))

    def update_settings_labels(self):
        self.reference_label.set("References: " + (self.settings.reference_path or "Bundled updated CENSUS and SKU data"))
        self.folder_label.set(f"Save to: {self.settings.output_directory}")

    def set_busy(self, busy: bool):
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.browse_button.configure(state=state)
        self.settings_button.configure(state=state)
        self.generate_button.configure(state="normal" if not busy and self.sheet.get() else "disabled")
        self.sheet_picker.configure(state="readonly" if not busy and self.sheet.get() else "disabled")

    def browse(self):
        path = filedialog.askopenfilename(parent=self.root, title="Select customer open order export",
                                          filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        self.result = None
        self.summary.set("Your completed report will appear here.")
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.raw_path.set(path)
        self.sheet.set("")
        self.sheet_picker.configure(values=())
        self.progress["value"] = 0
        self.status.set("Checking worksheets…")
        self.set_busy(True)
        def work():
            try:
                sheets = discover_sheets(Path(path))
                if not sheets:
                    raise WorkbookError("No worksheet has all 35 required raw headers on row 1. Select the original raw export.")
                self.events.put(("sheets", sheets))
            except Exception as exc:
                self.events.put(("error", exc))
        threading.Thread(target=work, daemon=True).start()

    def start_generation(self):
        if self.busy or not self.sheet.get():
            return
        raw, sheet = Path(self.raw_path.get()), self.sheet.get()
        ref = Path(self.settings.reference_path) if self.settings.reference_path else None
        folder = self.settings.output_directory
        self.result = None
        self.summary.set("Preparing your report…")
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.progress["value"] = 0
        self.set_busy(True)
        def work():
            try:
                result = generate(raw, sheet, folder, ref,
                                  lambda value, message: self.events.put(("progress", (value, message))))
                self.events.put(("done", result))
            except Exception as exc:
                self.events.put(("error", exc))
        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "progress":
                    self.progress["value"], message = value
                    self.status.set(message)
                elif kind == "sheets":
                    self.sheet_picker.configure(values=value)
                    self.sheet.set(value[0])
                    self.set_busy(False)
                    self.status.set("Ready to generate." if len(value) == 1 else "Multiple raw worksheets found. Choose the worksheet to process.")
                elif kind == "done":
                    self.result = value
                    self.set_busy(False)
                    self.status.set("Completed. Your Excel report is ready.")
                    self.summary.set(f"{value.input_rows:,} raw rows  →  {value.output_rows:,} report rows\n"
                                     f"{value.excluded_rows:,} zero-served rows excluded\n\n{value.path.name}")
                    self.open_file_button.configure(state="normal")
                    self.open_folder_button.configure(state="normal")
                elif kind == "error":
                    self.set_busy(False)
                    self.progress["value"] = 0
                    self.status.set("Unable to complete. Review the details and try again.")
                    self.summary.set("No new report was saved.")
                    if isinstance(value, ValidationError):
                        self.show_issues(value)
                    else:
                        messagebox.showerror("MTC Open Order", str(value), parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def show_issues(self, error: ValidationError):
        window = tk.Toplevel(self.root)
        window.title("Check raw data and references")
        window.geometry("1080x520")
        window.configure(bg=BG)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=str(error), wraplength=1000).pack(anchor="w", pady=(0, 12))
        grid = ttk.Frame(frame)
        grid.pack(fill="both", expand=True)
        tree = ttk.Treeview(grid, columns=("sheet", "row", "field", "value", "issue"), show="headings")
        for name, width in (("sheet", 130), ("row", 65), ("field", 140), ("value", 260), ("issue", 360)):
            tree.heading(name, text={"row": "Source row", "issue": "Issue"}.get(name, name.title()))
            tree.column(name, width=width, minwidth=60)
        scroll = ttk.Scrollbar(grid, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(grid, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
        grid.rowconfigure(0, weight=1)
        grid.columnconfigure(0, weight=1)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        for issue in error.issues:
            tree.insert("", "end", values=(issue.sheet, issue.row, issue.field, issue.value, issue.message))
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(15, 0))
        def save():
            name = filedialog.asksaveasfilename(parent=window, title="Save issue list", defaultextension=".csv",
                                               initialfile="MTC validation issues.csv", filetypes=[("CSV", "*.csv")])
            if name:
                try:
                    export_issues(error.issues, Path(name))
                except OSError as exc:
                    messagebox.showerror("Cannot save issue list", str(exc), parent=window)
        ttk.Button(buttons, text="Save Issue List…", command=save).pack(side="left")
        def settings():
            window.destroy()
            self.open_settings()
        ttk.Button(buttons, text="Change References…", command=settings).pack(side="left", padx=10)
        ttk.Button(buttons, text="Close", command=window.destroy).pack(side="right")

    def open_settings(self):
        window = tk.Toplevel(self.root)
        window.title("MTC Open Order — Settings")
        window.geometry("800x365")
        window.configure(bg=BG)
        window.transient(self.root)
        window.grab_set()
        frame = ttk.Frame(window, padding=24)
        frame.pack(fill="both", expand=True)
        reference = tk.StringVar(value=self.settings.reference_path)
        folder = tk.StringVar(value=str(self.settings.output_directory))
        ttk.Label(frame, text="Reference workbook (blank = bundled updated references)").pack(anchor="w")
        row = ttk.Frame(frame)
        row.pack(fill="x", pady=(8, 6))
        ttk.Entry(row, textvariable=reference).pack(side="left", fill="x", expand=True)
        def choose_ref():
            path = filedialog.askopenfilename(parent=window, filetypes=[("Excel workbook", "*.xlsx")])
            if path:
                reference.set(path)
        ttk.Button(row, text="Browse…", command=choose_ref).pack(side="right", padx=(10, 0))
        ttk.Button(frame, text="Use Bundled References", command=lambda: reference.set("")).pack(anchor="w")
        ttk.Label(frame, text="Output folder").pack(anchor="w", pady=(20, 8))
        row = ttk.Frame(frame)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=folder).pack(side="left", fill="x", expand=True)
        def choose_folder():
            path = filedialog.askdirectory(parent=window)
            if path:
                folder.set(path)
        ttk.Button(row, text="Browse…", command=choose_folder).pack(side="right", padx=(10, 0))
        def save():
            wb = None
            try:
                if reference.get().strip():
                    wb = open_xlsx(Path(reference.get().strip()), read_only=True)
                    validate_reference(wb)
                new_settings = Settings(reference.get().strip(), folder.get().strip())
                new_settings.save()
                self.settings = new_settings
                self.update_settings_labels()
                window.destroy()
            except (WorkbookError, OSError) as exc:
                messagebox.showerror("Cannot save settings", str(exc), parent=window)
            finally:
                if wb:
                    wb.close()
        ttk.Button(frame, text="Save Settings", style="Primary.TButton", command=save).pack(anchor="e", pady=(20, 0))

    def open_file(self):
        if self.result:
            self.launch(self.result.path)

    def open_folder(self):
        if self.result:
            self.launch(self.result.path.parent)

    def launch(self, path: Path):
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror("Cannot open", str(exc), parent=self.root)

    def close(self):
        if self.busy:
            messagebox.showinfo("Processing", "Please wait for the current operation to finish before closing.", parent=self.root)
            return
        self.root.destroy()


def main():
    root = tk.Tk()
    Application(root)
    root.mainloop()
