"""
*******************************************************************************
*** Program Name:   SASPyStudio.py                                          ***
***                                                                         ***
*** Application:    SASPy Studio                                            ***
*** Version:        1.0                                                     ***
***                                                                         ***
*** Purpose:        Provide a lightweight desktop interface for executing   ***
***                 local SAS programs using SAS OnDemand for Academics     ***
***                 through SASPy.                                          ***
***-------------------------------------------------------------------------***
*** Programmed By:  Manivannan Mathialagan                                  ***
*** Created On:     30Sep2026                                               ***
***-------------------------------------------------------------------------***
*** Process:        1. Connect to SAS OnDemand.                             ***
***                 2. Create and synchronize the remote study workspace.   ***
***                 3. Execute the study AUTOEXEC program.                  ***
***                 4. Assign RAW, SDTM and ADAM libraries.                 ***
***                 5. Load global and study SAS macros.                    ***
***                 6. Execute the selected SAS program/program plan.       ***
***                 7. Check and save SAS logs and results.                 ***
***                 8. Download generated SAS datasets and outputs.         ***
***                 9. Synchronize outputs to the local study structure.    ***
***                10. End the SAS session.                                 ***
***-------------------------------------------------------------------------***
*** Security:       SAS OnDemand credentials are read from the private      ***
***                 _authinfo file configured in sas_config.json.           ***
***                 Authentication files must not be committed to source    ***
***                 control.                                                ***
***-------------------------------------------------------------------------***
*** Notes:          Local SAS programs remain platform independent.         ***
***                 Windows study paths are translated as required when     ***
***                 programs are submitted to SAS OnDemand/Linux.           ***
***                                                                         ***
***                 Remote study files are temporarily synchronized to      ***
***                 the SAS WORK location during program execution.         ***
***-------------------------------------------------------------------------***
*** Change History:                                                         ***
***                                                                         ***
*** Version   Date         Description                                      ***
*** -------   -----------  -------------------------------------------------***
*** 1.0       30Sep2026    Initial release of SASPy Studio with SASPy       ***
***                        connectivity, remote study synchronization,      ***
***                        AUTOEXEC execution, RAW/SDTM/ADAM library        ***
***                        assignment, global/study macro loading, program  ***
***                        execution, log/result handling and output        ***
***                        synchronization.                                 ***
*******************************************************************************
"""


import json
import os
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import saspy
from openpyxl import load_workbook


APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "sas_config.json"
RUNTIME_CFG_FILE = APP_DIR / ".saspy_runtime_cfg.py"

DEFAULT_AUTOEXEC = APP_DIR / "autoexec.sas"
DEFAULT_GLOBAL_MACROS = APP_DIR / "macros"
DEFAULT_STUDY_MACROS = ""
DEFAULT_LOG_DIR = APP_DIR / "logs"
DEFAULT_RESULT_DIR = APP_DIR / "results"
DEFAULT_OUTPUT_DIR = APP_DIR / "output"


def load_sas_config():
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Missing configuration file:\n{CONFIG_FILE}\n\n"
            "Keep sas_config.json beside SASPyStudio.py."
        )

    cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))

    required = (
        "config_name",
        "java_path",
        "iom_host",
        "iom_port",
        "authkey",
        "authinfo",
    )
    missing = [name for name in required if name not in cfg]
    if missing:
        raise ValueError("sas_config.json is missing: " + ", ".join(missing))

    return cfg


def resolve_authinfo_path(cfg):
    authinfo_value = str(cfg["authinfo"]).strip()
    authinfo_path = Path(authinfo_value).expanduser()

    if not authinfo_path.is_absolute():
        authinfo_path = APP_DIR / authinfo_path

    authinfo_path = authinfo_path.resolve()

    if not authinfo_path.exists():
        raise FileNotFoundError(
            f"SAS authinfo file not found:\n{authinfo_path}\n\n"
            "Check the 'authinfo' value in sas_config.json."
        )

    return authinfo_path


def build_runtime_config(cfg):
    cfgname = str(cfg["config_name"])
    java = str(cfg["java_path"])
    host = cfg["iom_host"]
    port = int(cfg["iom_port"])
    authkey = str(cfg["authkey"])
    encoding = str(cfg.get("encoding", "utf-8"))

    if isinstance(host, list):
        host_repr = repr([str(x) for x in host])
    else:
        host_repr = repr(str(host))

    config_text = f"""SAS_config_names = [{cfgname!r}]

{cfgname} = {{
    'java'      : {java!r},
    'iomhost'   : {host_repr},
    'iomport'   : {port},
    'authkey'   : {authkey!r},
    'encoding'  : {encoding!r}
}}
"""
    RUNTIME_CFG_FILE.write_text(config_text, encoding="utf-8")
    return cfgname


def flatten_submit_part(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        lines = []
        for item in value:
            if isinstance(item, dict):
                lines.append(str(item.get("line", "")))
            else:
                lines.append(str(item))
        return "\n".join(lines)
    return str(value)


PROBLEM_NOTE_PATTERNS = (
    "uninitialized",
    "invalid data",
    "missing values were generated",
    "mathematical operations could not be performed",
    "division by zero",
    "numeric values have been converted to character",
    "character values have been converted to numeric",
    "at least one w.d format was too small",
    "merge statement has more than one data set with repeats of by values",
    "variable is uninitialized",
    "lost card",
)


def check_sas_log(log_text):
    errors = []
    warnings = []
    notes = []

    for raw in (log_text or "").splitlines():
        line = raw.strip()
        upper = line.upper()

        if upper.startswith("ERROR:"):
            errors.append(line)
        elif upper.startswith("WARNING:"):
            warnings.append(line)
        elif upper.startswith("NOTE:"):
            low = line.lower()
            if any(pattern in low for pattern in PROBLEM_NOTE_PATTERNS):
                notes.append(line)

    return errors, warnings, notes


class SASPyStudio(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("SASPy Studio")
        self.geometry("1120x820")
        self.minsize(980, 700)

        # Project locations. The finalized repository/study structure is derived from these roots.
        self.study_root = tk.StringVar()
        self.programs_root = tk.StringVar()
        self.global_macros_path = tk.StringVar(value=str(DEFAULT_GLOBAL_MACROS))
        self.program_path = tk.StringVar()
        self.program_queue = []
        self.use_program_plan = tk.BooleanVar(value=False)
        self.download_outputs = tk.BooleanVar(value=True)

        # Derived paths are refreshed whenever a run/preview starts.
        self.autoexec_path = tk.StringVar()
        self.study_macros_path = tk.StringVar()
        self.program_plan_path = tk.StringVar()
        self.log_dir = tk.StringVar()
        self.result_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.use_autoexec = tk.BooleanVar(value=True)
        self.use_global_macros = tk.BooleanVar(value=True)
        self.use_study_macros = tk.BooleanVar(value=True)
        self.sync_study = tk.BooleanVar(value=True)

        self.status = tk.StringVar(value="Ready")
        self.info = tk.StringVar(
            value="Fresh SAS session • XLSX program plans • validation • study synchronization"
        )
        self.log_counts = tk.StringVar(
            value="Errors: 0   Warnings: 0   Important Notes: 0"
        )

        self.running = False
        self.build()

    def build(self):
        BG = "#F4F5F7"
        PANEL = "#FFFFFF"
        TEXT = "#202124"
        MUTED = "#667085"
        BORDER = "#D9DDE5"
        BLUE = "#2563EB"
        BLUE_ACTIVE = "#1D4ED8"
        PURPLE = "#7C3AED"
        PURPLE_ACTIVE = "#6D28D9"
        GREEN = "#16A34A"
        AMBER = "#F59E0B"
        AMBER_ACTIVE = "#D97706"
        RED = "#DC2626"
        RED_ACTIVE = "#B91C1C"
        SOFT = "#EEF1F5"
        SOFT_ACTIVE = "#E1E6EC"

        self.configure(bg=BG)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Card.TLabelframe",
            background=PANEL,
            bordercolor=BORDER,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=BG,
            foreground=TEXT,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure("Card.TLabel", background=PANEL, foreground=TEXT)

        class RoundedButton(tk.Canvas):
            def __init__(
                self, parent, text, command, bg, active_bg,
                fg="white", width=120, height=36, radius=10,
                font=("Segoe UI", 9, "bold")
            ):
                super().__init__(
                    parent, width=width, height=height,
                    bg=parent.cget("bg"), highlightthickness=0,
                    bd=0, cursor="hand2"
                )
                self.command = command
                self.normal_bg = bg
                self.active_bg = active_bg
                self.fg = fg
                self.radius = radius
                self.w = width
                self.h = height
                self.font_value = font
                self.text_value = text
                self.enabled = True
                self.draw(bg)
                self.bind(
                    "<Enter>",
                    lambda e: self.draw(self.active_bg) if self.enabled else None
                )
                self.bind(
                    "<Leave>",
                    lambda e: self.draw(self.normal_bg) if self.enabled else None
                )
                self.bind("<Button-1>", self._click)

            def rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
                points = [
                    x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
                    x2, y2-r, x2, y2, x2-r, y2, x1+r, y2,
                    x1, y2, x1, y2-r, x1, y1+r, x1, y1
                ]
                return self.create_polygon(points, smooth=True, **kwargs)

            def draw(self, color):
                self.delete("all")
                self.rounded_rect(
                    1, 1, self.w-1, self.h-1, self.radius,
                    fill=color, outline=color
                )
                self.create_text(
                    self.w/2, self.h/2,
                    text=self.text_value, fill=self.fg,
                    font=self.font_value
                )

            def _click(self, _event):
                if self.enabled and self.command:
                    self.command()

            def configure(self, cnf=None, **kwargs):
                state = kwargs.pop("state", None)
                if state is not None:
                    self.enabled = str(state) not in ("disabled", "DISABLED")
                    super().configure(cursor="hand2" if self.enabled else "arrow")
                    self.draw(self.normal_bg if self.enabled else "#B8C0CC")
                if kwargs or cnf:
                    return super().configure(cnf, **kwargs)

            config = configure

        class GreenCheck(tk.Frame):
            def __init__(self, parent, text, variable):
                super().__init__(parent, bg=PANEL)
                self.variable = variable
                self.box = tk.Canvas(
                    self, width=20, height=20, bg=PANEL,
                    highlightthickness=0, bd=0, cursor="hand2"
                )
                self.box.pack(side="left")
                self.label = tk.Label(
                    self, text=text, bg=PANEL, fg=TEXT,
                    font=("Segoe UI", 9), cursor="hand2"
                )
                self.label.pack(side="left", padx=(5, 0))
                self.box.bind("<Button-1>", self.toggle)
                self.label.bind("<Button-1>", self.toggle)
                self.variable.trace_add("write", lambda *_: self.draw())
                self.draw()

            def toggle(self, _event=None):
                self.variable.set(not self.variable.get())

            def draw(self):
                self.box.delete("all")
                if self.variable.get():
                    self.box.create_rectangle(
                        2, 2, 18, 18, fill=GREEN, outline=GREEN, width=1
                    )
                    self.box.create_line(
                        5, 10, 9, 14, 16, 6,
                        fill="white", width=2.2,
                        capstyle="round", joinstyle="round"
                    )
                else:
                    self.box.create_rectangle(
                        2, 2, 18, 18,
                        fill="white", outline="#AAB2BF", width=1
                    )

        self._RoundedButton = RoundedButton

        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=20, pady=(14, 7))

        tk.Label(
            header, text="SASPy Studio",
            bg=BG, fg="#172554",
            font=("Segoe UI", 20, "bold")
        ).pack(side="left")

        tk.Label(
            header, text="SAS OnDemand",
            bg=BG, fg=PURPLE,
            font=("Segoe UI", 9, "bold")
        ).pack(side="left", padx=(12, 0), pady=(6, 0))

        # ---------- Project ----------
        project_card = ttk.LabelFrame(
            self, text="Project", style="Card.TLabelframe"
        )
        project_card.pack(fill="x", padx=20, pady=(0, 8))
        project_card.columnconfigure(1, weight=1)

        project_rows = [
            ("Study Location:", self.study_root, lambda: self.browse_folder(self.study_root)),
            ("Programs Location:", self.programs_root, lambda: self.browse_folder(self.programs_root)),
            ("Global Macros:", self.global_macros_path, lambda: self.browse_folder(self.global_macros_path)),
        ]
        for row, (label, variable, command) in enumerate(project_rows):
            ttk.Label(project_card, text=label, style="Card.TLabel").grid(
                row=row, column=0, padx=(12, 8), pady=5, sticky="w"
            )
            ttk.Entry(project_card, textvariable=variable).grid(
                row=row, column=1, padx=5, pady=5, sticky="ew"
            )
            wrap = tk.Frame(project_card, bg=PANEL)
            wrap.grid(row=row, column=2, padx=(5, 12), pady=4)
            RoundedButton(
                wrap, "Browse", command, PURPLE, PURPLE_ACTIVE,
                width=88, height=30, radius=9, font=("Segoe UI", 8, "bold")
            ).pack()

        tk.Label(
            project_card,
            text="Programs Location is the SAS folder. autoexec.sas, macros, program_plan.xlsx, production and validation folders are derived automatically.",
            bg=PANEL, fg=MUTED, font=("Segoe UI", 8)
        ).grid(row=3, column=1, columnspan=2, padx=5, pady=(0, 7), sticky="w")

        # ---------- Execution ----------
        exec_card = ttk.LabelFrame(
            self, text="Execution", style="Card.TLabelframe"
        )
        exec_card.pack(fill="x", padx=20, pady=(0, 8))
        exec_card.columnconfigure(1, weight=1)

        ttk.Label(exec_card, text="Program:", style="Card.TLabel").grid(
            row=0, column=0, padx=(12, 8), pady=6, sticky="w"
        )
        ttk.Entry(exec_card, textvariable=self.program_path).grid(
            row=0, column=1, padx=5, pady=6, sticky="ew"
        )
        program_wrap = tk.Frame(exec_card, bg=PANEL)
        program_wrap.grid(row=0, column=2, padx=(5, 12), pady=5)
        RoundedButton(
            program_wrap, "Browse", self.browse_manual_program,
            PURPLE, PURPLE_ACTIVE, width=88, height=30, radius=9,
            font=("Segoe UI", 8, "bold")
        ).pack()

        plan_wrap = tk.Frame(exec_card, bg=PANEL)
        plan_wrap.grid(row=1, column=0, padx=(12, 8), pady=5, sticky="w")
        GreenCheck(plan_wrap, "Use Program Plan", self.use_program_plan).pack()

        self.plan_display = ttk.Label(exec_card, text="program_plan.xlsx (derived)", style="Card.TLabel")
        self.plan_display.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        preview_wrap = tk.Frame(exec_card, bg=PANEL)
        preview_wrap.grid(row=1, column=2, padx=(5, 12), pady=4)
        RoundedButton(
            preview_wrap, "Preview", self.preview_program_plan,
            SOFT, SOFT_ACTIVE, fg="#374151", width=88, height=30,
            radius=9, font=("Segoe UI", 8, "bold")
        ).pack()

        options_wrap = tk.Frame(exec_card, bg=PANEL)
        options_wrap.grid(row=2, column=1, columnspan=2, padx=5, pady=(2, 8), sticky="w")
        GreenCheck(options_wrap, "Download WORK outputs", self.download_outputs).pack(side="left")
        tk.Label(
            options_wrap,
            text="   Study synchronization is automatic for the selected Study Location.",
            bg=PANEL, fg=MUTED, font=("Segoe UI", 8)
        ).pack(side="left")

        # ---------- Toolbar ----------
        toolbar = tk.Frame(
            self, bg=PANEL,
            highlightbackground=BORDER, highlightthickness=1
        )
        toolbar.pack(fill="x", padx=20, pady=(0, 8))

        left = tk.Frame(toolbar, bg=PANEL)
        left.pack(side="left", padx=10, pady=8)

        self.runbtn = RoundedButton(
            left, "▶  Run SAS", self.start_run,
            BLUE, BLUE_ACTIVE, width=126, height=40,
            radius=11, font=("Segoe UI", 10, "bold")
        )
        self.runbtn.pack(side="left")

        self.restartbtn = RoundedButton(
            left, "↻  Restart SAS", self.restart_sas,
            AMBER, AMBER_ACTIVE, fg="#3F2D00",
            width=126, height=36, radius=10
        )
        self.restartbtn.pack(side="left", padx=(7, 0))

        self.closebtn = RoundedButton(
            left, "Close Studio", self.close_app,
            RED, RED_ACTIVE, width=112, height=36, radius=10
        )
        self.closebtn.pack(side="left", padx=(7, 0))

        right = tk.Frame(toolbar, bg=PANEL)
        right.pack(side="right", padx=10, pady=8)
        RoundedButton(
            right, "Clear Window", self.clear_execution_window,
            SOFT, SOFT_ACTIVE, fg="#374151", width=104, height=32,
            radius=9, font=("Segoe UI", 8, "bold")
        ).pack(side="left")

        # ---------- Run information ----------
        info_card = ttk.LabelFrame(
            self, text="Run Information", style="Card.TLabelframe"
        )
        info_card.pack(fill="x", padx=20, pady=(0, 8))

        self.status_label = tk.Label(
            info_card, textvariable=self.status,
            bg=PANEL, fg=BLUE,
            font=("Segoe UI", 10, "bold")
        )
        self.status_label.pack(anchor="w", padx=12, pady=(8, 2))

        tk.Label(
            info_card, textvariable=self.info,
            bg=PANEL, fg=MUTED, font=("Segoe UI", 9)
        ).pack(anchor="w", padx=12, pady=(0, 2))

        tk.Label(
            info_card, textvariable=self.log_counts,
            bg=PANEL, fg=TEXT, font=("Segoe UI", 9, "bold")
        ).pack(anchor="w", padx=12, pady=(0, 8))

        # ---------- Execution ----------
        output_card = ttk.LabelFrame(
            self, text="Execution / Log Check", style="Card.TLabelframe"
        )
        output_card.pack(
            fill="both", expand=True, padx=20, pady=(0, 14)
        )

        self.box = ScrolledText(
            output_card,
            font=("Consolas", 9),
            bg="#FCFCFD", fg="#263238",
            relief="flat", bd=0
        )
        self.box.pack(fill="both", expand=True, padx=8, pady=8)
        self.box.configure(state="disabled")

    def refresh_derived_paths(self):
        programs_text = self.programs_root.get().strip()
        study_text = self.study_root.get().strip()
        if programs_text:
            programs = Path(programs_text).expanduser()
            self.autoexec_path.set(str(programs / "autoexec.sas"))
            self.study_macros_path.set(str(programs / "macros"))
            self.program_plan_path.set(str(programs / "program_plan.xlsx"))
            if hasattr(self, "plan_display"):
                self.plan_display.configure(text=str(programs / "program_plan.xlsx"))
        if study_text:
            study = Path(study_text).expanduser()
            self.log_dir.set(str(study / "logs" / "sas"))
            self.result_dir.set(str(study / "results" / "sas"))
            self.output_dir.set(str(study / "output" / "sas"))

    def browse_manual_program(self):
        current = self.program_path.get().strip()
        programs = self.programs_root.get().strip()
        initial = Path(programs) if programs else APP_DIR
        if current and Path(current).exists():
            initial = Path(current).parent
        filename = filedialog.askopenfilename(
            title="Select SAS Program",
            initialdir=str(initial) if initial.exists() else str(APP_DIR),
            filetypes=[("SAS programs", "*.sas"), ("All files", "*.*")],
        )
        if filename:
            self.program_path.set(filename)

    def clear_execution_window(self):
        self.box.configure(state="normal")
        self.box.delete("1.0", "end")
        self.box.configure(state="disabled")
        self.status.set("Ready")
        self.log_counts.set("Errors: 0   Warnings: 0   Important Notes: 0")

    def add_programs(self):
        files = filedialog.askopenfilenames(
            title="Add SAS Programs",
            initialdir=str(APP_DIR),
            filetypes=[("SAS programs", "*.sas"), ("All files", "*.*")],
        )
        for filename in files:
            path = str(Path(filename).resolve())
            if path not in self.program_queue:
                self.program_queue.append(path)
                self.queue_list.insert("end", path)

    def remove_programs(self):
        selected = list(self.queue_list.curselection())
        for index in reversed(selected):
            del self.program_queue[index]
            self.queue_list.delete(index)

    def clear_program_queue(self):
        self.program_queue.clear()
        self.queue_list.delete(0, "end")

    def move_program(self, direction):
        selected = list(self.queue_list.curselection())
        if len(selected) != 1:
            messagebox.showinfo(
                "Program Order",
                "Select one program at a time to move it up or down."
            )
            return
        index = selected[0]
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.program_queue):
            return
        self.program_queue[index], self.program_queue[new_index] = (
            self.program_queue[new_index], self.program_queue[index]
        )
        self.queue_list.delete(0, "end")
        for item in self.program_queue:
            self.queue_list.insert("end", item)
        self.queue_list.selection_set(new_index)
        self.queue_list.see(new_index)

    def browse_sas_file(self, variable):
        current = variable.get().strip()
        initial = APP_DIR
        if current:
            p = Path(current)
            initial = p.parent if p.suffix else p

        filename = filedialog.askopenfilename(
            title="Select SAS Program",
            initialdir=str(initial) if Path(initial).exists() else str(APP_DIR),
            filetypes=[("SAS programs", "*.sas"), ("All files", "*.*")],
        )
        if filename:
            variable.set(filename)

    def browse_program_plan(self):
        current = self.program_plan_path.get().strip()
        initial = Path(current).parent if current else APP_DIR
        filename = filedialog.askopenfilename(
            title="Select Program Plan",
            initialdir=str(initial) if Path(initial).exists() else str(APP_DIR),
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.program_plan_path.set(filename)

    @staticmethod
    def _yes(value):
        return str(value or "").strip().upper() in {"Y", "YES", "TRUE", "1", "X"}

    @staticmethod
    def _plan_type(value):
        text = str(value or "").strip().upper()
        aliases = {
            "ADAM": "adam", "ADaM": "adam", "SDTM": "sdtm",
            "TABLE": "tables", "TABLES": "tables",
            "LISTING": "listings", "LISTINGS": "listings",
            "FIGURE": "figures", "FIGURES": "figures",
        }
        return aliases.get(text, "")

    def read_program_plan(self, plan_path, code_root):
        # Read the last saved XLSX from disk. Excel may remain open.
        wb = load_workbook(plan_path, data_only=True, read_only=True)
        try:
            ws = wb["Programs"] if "Programs" in wb.sheetnames else wb[wb.sheetnames[0]]
            rows = ws.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                raise ValueError("Program plan is empty.")

            def key(value):
                return " ".join(str(value or "").strip().lower().replace("_", " ").split())

            columns = {key(value): idx for idx, value in enumerate(header)}
            required = ["type", "program name", "order", "run production", "run validation"]
            missing = [name for name in required if name not in columns]
            if missing:
                raise ValueError("Program plan is missing column(s): " + ", ".join(missing))

            plan = []
            for excel_row, values in enumerate(rows, start=2):
                ptype_raw = values[columns["type"]] if columns["type"] < len(values) else None
                program_raw = values[columns["program name"]] if columns["program name"] < len(values) else None
                if not ptype_raw and not program_raw:
                    continue

                ptype = self._plan_type(ptype_raw)
                if not ptype:
                    raise ValueError(f"Unsupported Type at Excel row {excel_row}: {ptype_raw}")

                program_name = str(program_raw or "").strip()
                if not program_name:
                    raise ValueError(f"Program Name is blank at Excel row {excel_row}.")
                if not program_name.lower().endswith(".sas"):
                    program_name += ".sas"

                order_value = values[columns["order"]] if columns["order"] < len(values) else None
                try:
                    order = float(order_value)
                except (TypeError, ValueError):
                    raise ValueError(f"Invalid Order at Excel row {excel_row}: {order_value}")

                run_prod = self._yes(values[columns["run production"]])
                run_val = self._yes(values[columns["run validation"]])
                if not run_prod and not run_val:
                    continue

                prod_path = code_root / ptype / program_name
                val_path = code_root / "validation" / ptype / ("v_" + program_name)

                plan.append({
                    "type": ptype, "program_name": program_name, "order": order,
                    "run_production": run_prod, "run_validation": run_val,
                    "production_path": prod_path, "validation_path": val_path,
                    "excel_row": excel_row,
                })

            type_order = {"sdtm": 1, "adam": 2, "tables": 3, "listings": 4, "figures": 5}
            plan.sort(key=lambda x: (type_order[x["type"]], x["order"], x["excel_row"]))
            return plan
        finally:
            wb.close()

    def preview_program_plan(self):
        try:
            self.refresh_derived_paths()
            programs_root = self._required_folder(self.programs_root.get(), "Programs location")
            plan_path = self._required_file(str(programs_root / "program_plan.xlsx"), "Program plan")
            plan = self.read_program_plan(plan_path, programs_root)
            lines = ["PROGRAM PLAN", "-" * 68]
            prod = val = 0
            for row in plan:
                if row["run_production"]:
                    prod += 1
                if row["run_validation"]:
                    val += 1
                lines.append(
                    f'{row["type"].upper():9} {row["order"]:>5g}  {row["program_name"]:<28} '
                    f'Prod={"Y" if row["run_production"] else "N"}  '
                    f'Val={"Y" if row["run_validation"] else "N"}'
                )
            lines.extend(["-" * 68, f"Production: {prod}   Validation: {val}"])
            messagebox.showinfo("Program Plan Preview", "\n".join(lines) if plan else "No programs are selected for execution.")
        except Exception as exc:
            messagebox.showerror("Program Plan", str(exc))

    def browse_folder(self, variable):
        current = variable.get().strip()
        initial = current if current and Path(current).exists() else str(APP_DIR)

        folder = filedialog.askdirectory(
            title="Select Folder",
            initialdir=initial,
        )
        if folder:
            variable.set(folder)

    def restart_sas(self):
        if self.running:
            messagebox.showinfo(
                "SAS is running",
                "Wait for the current run to finish. The next Run SAS "
                "automatically starts a fresh SAS session."
            )
            return

        self.status.set("Ready - next run starts a fresh SAS session")
        self.write_box(
            "Restart SAS: next run will start a fresh SAS session."
        )

    def close_app(self):
        if self.running:
            messagebox.showwarning(
                "SAS is running",
                "SAS is currently executing. Studio will not force-close the "
                "active IOM connection. Wait for the current submit to return, "
                "then close Studio."
            )
            return
        self.destroy()

    @staticmethod
    def get_work_path(sas):
        """Return the resolved remote SAS WORK path, never the echoed SAS code."""
        token = "__SASPYSTUDIO_WORK__="
        result = sas.submit(
            f"%put {token}%sysfunc(pathname(work));"
        )
        log = flatten_submit_part(result.get("LOG"))

        # SASPy logs can contain both the submitted source statement and the
        # value written by %PUT.  Only accept a line that STARTS with our
        # marker; otherwise the echoed source can be mistaken for the path.
        for raw_line in log.splitlines():
            line = raw_line.strip()
            for prefix in (token, "NOTE: " + token):
                if line.startswith(prefix):
                    path = line[len(prefix):].strip().replace("\\", "/")
                    if path and "%sysfunc" not in path.lower():
                        return path.rstrip("/")

        raise RuntimeError(
            "Could not determine the resolved remote SAS WORK path.\n\n" + log
        )

    @staticmethod
    def list_remote_files(sas, remote_dir):
        if not remote_dir:
            return set()

        token = "__SPYFILE__="
        quoted = remote_dir.replace('"', '""')
        sas_code = (
            'filename _spydir "' + quoted + '";\n'
            'data _null_;\n'
            '  length name $512;\n'
            "  did=dopen('_spydir');\n"
            '  if did > 0 then do i=1 to dnum(did);\n'
            '    name=dread(did,i);\n'
            '    put "' + token + '" name;\n'
            '  end;\n'
            '  rc=dclose(did);\n'
            'run;\n'
            'filename _spydir clear;\n'
        )

        result = sas.submit(sas_code)
        log = flatten_submit_part(result.get("LOG"))

        return {
            line.split(token, 1)[1].strip()
            for line in log.splitlines()
            if token in line
        }

    def download_work_outputs(self, sas, remote_dir, before, run_dir):
        allowed = {
            ".csv", ".txt", ".rtf", ".pdf", ".xlsx", ".xls",
            ".xml", ".json", ".zip", ".png", ".jpg", ".jpeg",
            ".svg", ".html", ".htm", ".lst", ".xpt"
        }

        after = self.list_remote_files(sas, remote_dir)
        names = sorted(
            name for name in (after - before)
            if Path(name).suffix.lower() in allowed
        )

        downloaded = []
        if names:
            run_dir.mkdir(parents=True, exist_ok=True)

        for name in names:
            remote = remote_dir.rstrip("/\\") + "/" + name
            local = run_dir / name
            try:
                sas.download(str(local), remote, overwrite=True)
                downloaded.append(local)
            except Exception as exc:
                self.write_box(
                    f"Output download warning for {name}: {exc}"
                )

        return downloaded

    @staticmethod
    def _sas_quote(value):
        return str(value).replace('"', '""')

    def ensure_remote_dir(self, sas, remote_dir):
        """Create and verify one remote directory level on the SAS host."""
        remote_dir = remote_dir.replace("\\", "/").rstrip("/")
        parent, leaf = remote_dir.rsplit("/", 1)
        qparent = self._sas_quote(parent)
        qleaf = self._sas_quote(leaf)
        qdir = self._sas_quote(remote_dir)
        marker = "__SASPYSTUDIO_DIR__="

        # DCREATE expects the actual parent-directory pathname as its second
        # argument.  Passing a fileref name here can silently fail on ODA.
        code = (
            'data _null_;\n'
            '  length newdir $2048;\n'
            f'  if fileexist("{qdir}") then newdir="{qdir}";\n'
            f'  else newdir=dcreate("{qleaf}","{qparent}");\n'
            f'  if fileexist("{qdir}") then put "{marker}OK|{qdir}";\n'
            f'  else put "{marker}FAIL|{qdir}";\n'
            'run;\n'
        )
        result = sas.submit(code)
        log = flatten_submit_part(result.get("LOG"))

        # SAS can wrap long log lines, splitting the remote pathname across
        # physical lines.  Verify only the short status token; the DATA step
        # has already confirmed FILEEXIST(qdir) before emitting OK.
        ok_marker = f"{marker}OK|"
        if ok_marker not in log:
            raise RuntimeError(
                "SASPyStudio could not create/verify remote directory:\n"
                f"{remote_dir}\n\nSAS log:\n{log}"
            )
        return True

    def prepare_remote_study(self, sas, local_root, remote_work):
        """Mirror standard study inputs into a temporary remote study tree."""
        remote_root = remote_work.rstrip("/\\") + "/SASPyStudio_Study"
        rel_dirs = [
            "data", "data/raw", "data/sas", "data/sas/raw",
            "data/sas/sdtm", "data/sas/adam",
            "output", "output/sas", "output/sas/sdtm", "output/sas/adam",
            "output/sas/tables", "output/sas/listings", "output/sas/figures",
            "logs", "logs/sas", "results", "results/sas",
        ]

        self.ensure_remote_dir(sas, remote_root)
        for rel in rel_dirs:
            self.ensure_remote_dir(sas, remote_root + "/" + rel)

        uploaded = 0
        upload_roots = [
            local_root / "data" / "raw",
            local_root / "data" / "sas",
        ]
        for base in upload_roots:
            if not base.exists():
                continue
            for local_file in sorted(p for p in base.rglob("*") if p.is_file()):
                rel = local_file.relative_to(local_root).as_posix()
                remote_file = remote_root + "/" + rel
                remote_parent = remote_file.rsplit("/", 1)[0]
                self.ensure_remote_dir(sas, remote_parent)
                sas.upload(str(local_file), remote_file, overwrite=True)
                uploaded += 1

        return remote_root, uploaded

    def assign_remote_libraries(self, sas, remote_root):
        """Assign remote study librefs using literal server paths and verify them."""
        root = remote_root.replace("\\", "/").rstrip("/")
        paths = {
            "RAW": root + "/data/sas/raw",
            "SDTM": root + "/data/sas/sdtm",
            "ADAM": root + "/data/sas/adam",
        }

        self.write_box("")
        self.write_box("Assigning remote libraries...")

        logs = []
        for libref, path in paths.items():
            qpath = self._sas_quote(path)
            marker = f"__SASPYSTUDIO_{libref}_RC__="
            code = (
                f'libname {libref.lower()} "{qpath}";\n'
                f'%put {marker}%sysfunc(libref({libref.lower()}));\n'
            )
            result = sas.submit(code)
            log = flatten_submit_part(result.get("LOG"))
            logs.append(f"/* {libref} */\n{log}")

            rc = None
            for raw_line in log.splitlines():
                line = raw_line.strip()
                if line.startswith(marker):
                    rc = line[len(marker):].strip()
                    break

            if rc == "0":
                self.write_box(f"  {libref:<4} ..................... OK")
            else:
                self.write_box(f"  {libref:<4} ..................... FAILED (RC={rc or 'unknown'})")
                raise RuntimeError(
                    f"Remote {libref} library could not be assigned to:\n{path}\n\nSAS log:\n{log}"
                )

        return "\n".join(logs)

    @staticmethod
    def rewrite_study_paths(source, local_root, remote_root):
        """Replace the configured local study root with its remote mirror."""
        local_text = str(local_root)
        variants = {
            local_text,
            local_text.replace("\\", "/"),
            local_text.replace("/", "\\"),
        }
        for value in sorted(variants, key=len, reverse=True):
            source = source.replace(value, remote_root)

        # SAS OnDemand runs on a Unix-like host. Normalize separators only in
        # known study-path macro assignments in the submitted source. This keeps
        # the physical/local SAS files Windows-friendly and avoids changing
        # unrelated SAS strings, escape characters, or program logic.
        import re

        path_macros = (
            "studyroot", "rawsrc", "rawpath", "sdtmpath", "adampath",
            "tblpath", "lstpath", "figpath", "logpath", "respath",
        )
        macro_pattern = re.compile(
            r"(?im)^(\s*%let\s+(?:" + "|".join(path_macros) + r")\s*=)(.*?)(;\s*)$"
        )

        def normalize_path_assignment(match):
            prefix, value, suffix = match.groups()
            return prefix + value.replace("\\", "/") + suffix

        source = macro_pattern.sub(normalize_path_assignment, source)
        return source

    def sync_remote_study_back(self, sas, local_root, remote_root):
        """Download permanent SAS data and final SAS outputs to local study folders."""
        rel_dirs = [
            "data/sas/raw", "data/sas/sdtm", "data/sas/adam",
            "output/sas/tables", "output/sas/listings", "output/sas/figures",
        ]
        downloaded = []
        for rel in rel_dirs:
            remote_dir = remote_root.rstrip("/") + "/" + rel
            local_dir = local_root / Path(rel)
            names = sorted(self.list_remote_files(sas, remote_dir))
            if not names:
                continue
            local_dir.mkdir(parents=True, exist_ok=True)
            for name in names:
                remote_file = remote_dir.rstrip("/") + "/" + name
                local_file = local_dir / name
                try:
                    sas.download(str(local_file), remote_file, overwrite=True)
                    downloaded.append(local_file)
                except Exception as exc:
                    self.write_box(
                        f"Study sync download warning for {rel}/{name}: {exc}"
                    )
        return downloaded

    def write_box(self, text):
        def update():
            self.box.configure(state="normal")
            self.box.insert("end", text.rstrip() + "\n")
            self.box.see("end")
            self.box.configure(state="disabled")
        self.after(0, update)

    def set_status(self, text):
        self.after(0, lambda: self.status.set(text))

    def clear_selected_folder(self, variable, label):
        if self.running:
            messagebox.showinfo(
                "SAS is running",
                "Wait for the current SAS run to finish before clearing files."
            )
            return

        value = variable.get().strip()
        if not value:
            messagebox.showwarning(
                f"Clear {label.title()}",
                f"No {label} folder has been selected."
            )
            return

        folder = Path(value)
        if not folder.exists():
            messagebox.showwarning(
                f"Clear {label.title()}",
                f"Folder does not exist:\n{folder}"
            )
            return

        if not messagebox.askyesno(
            f"Clear {label.title()}",
            f"Delete all files directly inside this folder?\n\n{folder}"
        ):
            return

        count = 0
        for item in folder.iterdir():
            if item.is_file():
                item.unlink()
                count += 1

        self.status.set(f"Cleared {label} ({count} file(s))")

    @staticmethod
    def _required_file(value, label):
        if not value.strip():
            raise ValueError(f"{label} has not been selected.")
        path = Path(value).expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"{label} not found:\n{path}")
        return path.resolve()

    @staticmethod
    def _required_folder(value, label):
        if not value.strip():
            raise ValueError(f"{label} folder has not been selected.")
        path = Path(value).expanduser()
        if not path.is_dir():
            raise FileNotFoundError(f"{label} folder not found:\n{path}")
        return path.resolve()

    @staticmethod
    def _destination_folder(value, label):
        if not value.strip():
            raise ValueError(f"{label} folder has not been selected.")
        path = Path(value).expanduser().resolve()
        path.mkdir(parents=True, exist_ok=True)
        return path

    def start_run(self):
        if self.running:
            return

        try:
            self.refresh_derived_paths()
            study_root = self._required_folder(self.study_root.get(), "Study location")
            programs_root = self._required_folder(self.programs_root.get(), "Programs location")
            autoexec = self._required_file(str(programs_root / "autoexec.sas"), "Autoexec")
            global_macros = self._required_folder(self.global_macros_path.get(), "Global macros")
            study_macros_path = programs_root / "macros"
            study_macros = study_macros_path.resolve() if study_macros_path.is_dir() else None

            plan_items = []
            if self.use_program_plan.get():
                plan_path = self._required_file(str(programs_root / "program_plan.xlsx"), "Program plan")
                plan = self.read_program_plan(plan_path, programs_root)
                for row in plan:
                    if row["run_production"]:
                        path = self._required_file(str(row["production_path"]), "Production program")
                        plan_items.append({"path": path, "type": row["type"], "kind": "production", "order": row["order"]})
                    if row["run_validation"]:
                        path = self._required_file(str(row["validation_path"]), "Validation program")
                        plan_items.append({"path": path, "type": row["type"], "kind": "validation", "order": row["order"]})
                if not plan_items:
                    raise ValueError("No programs are selected in program_plan.xlsx.")
                programs = [item["path"] for item in plan_items]
            else:
                program = self._required_file(self.program_path.get(), "SAS program")
                try:
                    relative = program.relative_to(programs_root)
                    manual_type = relative.parts[0].lower() if len(relative.parts) > 1 else "setup"
                except ValueError:
                    manual_type = "manual"
                programs = [program]
                plan_items = [{"path": program, "type": manual_type, "kind": "production", "order": 1}]

            program = programs[0]
            manual_type = plan_items[0]["type"] if not self.use_program_plan.get() else "plan"
            log_dir = study_root / "logs" / "sas" / manual_type if manual_type != "plan" else study_root / "logs" / "sas"
            result_dir = study_root / "results" / "sas" / manual_type if manual_type != "plan" else study_root / "results" / "sas"
            output_dir = study_root / "output" / "sas" / manual_type if manual_type != "plan" else study_root / "output" / "sas"
            for folder in (log_dir, result_dir, output_dir):
                folder.mkdir(parents=True, exist_ok=True)

            settings = {
                "program": program,
                "programs": programs,
                "plan_items": plan_items,
                "use_program_plan": self.use_program_plan.get(),
                "use_autoexec": True,
                "use_global_macros": True,
                "use_study_macros": study_macros is not None,
                "download_outputs": self.download_outputs.get(),
                "sync_study": True,
                "study_root": study_root,
                "programs_root": programs_root,
                "autoexec": autoexec,
                "global_macros": global_macros,
                "study_macros": study_macros,
                "log_dir": log_dir,
                "result_dir": result_dir,
                "output_dir": output_dir,
            }

        except Exception as exc:
            messagebox.showerror("SASPy Studio", str(exc))
            return

        self.running = True
        self.log_counts.set("Errors: 0   Warnings: 0   Important Notes: 0")
        self.runbtn.configure(state="disabled")
        self.box.configure(state="normal")
        self.box.delete("1.0", "end")
        self.box.configure(state="disabled")

        threading.Thread(
            target=self.run_worker, args=(settings,), daemon=True
        ).start()

    def run_worker(self, settings):
        sas = None
        started = time.perf_counter()

        all_log_parts = []
        final_lst = ""

        program = settings["program"]
        run_id = time.strftime("%Y%m%d_%H%M%S")

        run_name = "program_plan" if settings["use_program_plan"] else (program.stem if len(settings["programs"]) == 1 else "batch_run")
        log_file = (
            settings["log_dir"] /
            f"{run_name}_{run_id}.log"
        )
        result_file = (
            settings["result_dir"] /
            f"{run_name}_{run_id}.html"
        )
        run_output_dir = settings["output_dir"] / run_id

        try:
            self.set_status("Connecting to SAS OnDemand...")
            self.write_box("Connecting to SAS OnDemand...")

            cfg = load_sas_config()
            cfgname = build_runtime_config(cfg)
            authinfo_path = resolve_authinfo_path(cfg)

            self.write_box(f"Authinfo: {authinfo_path}")

            t0 = time.perf_counter()

            old_userprofile = os.environ.get("USERPROFILE")
            old_home = os.environ.get("HOME")
            try:
                os.environ["USERPROFILE"] = str(authinfo_path.parent)
                os.environ["HOME"] = str(authinfo_path.parent)

                sas = saspy.SASsession(
                    cfgname=cfgname,
                    cfgfile=str(RUNTIME_CFG_FILE),
                    results="HTML",
                )
            finally:
                if old_userprofile is None:
                    os.environ.pop("USERPROFILE", None)
                else:
                    os.environ["USERPROFILE"] = old_userprofile

                if old_home is None:
                    os.environ.pop("HOME", None)
                else:
                    os.environ["HOME"] = old_home

            self.write_box(
                f"Connected in {time.perf_counter() - t0:.1f} seconds."
            )

            remote_study_root = ""
            if settings["sync_study"]:
                self.set_status("Preparing remote study workspace...")
                remote_work_for_study = self.get_work_path(sas)
                if "%sysfunc" in remote_work_for_study.lower():
                    raise RuntimeError(
                        "SAS WORK path was not resolved: " + remote_work_for_study
                    )
                self.write_box(f"Remote SAS WORK: {remote_work_for_study}")
                remote_study_root, uploaded_count = self.prepare_remote_study(
                    sas, settings["study_root"], remote_work_for_study
                )
                self.write_box(
                    f"Study sync: {uploaded_count} local file(s) uploaded."
                )
                self.write_box(
                    f"Remote study root: {remote_study_root}"
                )

            # Build initialization list in a deterministic order.
            init_files = []

            if settings["use_autoexec"]:
                init_files.append(
                    ("AUTOEXEC", settings["autoexec"])
                )

            if settings["use_global_macros"]:
                global_files = sorted(
                    settings["global_macros"].glob("*.sas"),
                    key=lambda p: p.name.lower()
                )
                for sas_file in global_files:
                    init_files.append(("GLOBAL", sas_file))

            if settings["use_study_macros"]:
                study_files = sorted(
                    settings["study_macros"].glob("*.sas"),
                    key=lambda p: p.name.lower()
                )
                for sas_file in study_files:
                    init_files.append(("STUDY", sas_file))

            total = len(init_files)

            self.write_box("")
            self.write_box("Initialization plan:")
            self.write_box(
                f"  Autoexec: "
                f"{settings['autoexec'] if settings['use_autoexec'] else 'Skipped'}"
            )
            self.write_box(
                f"  Global macros: "
                f"{settings['global_macros'] if settings['use_global_macros'] else 'Skipped'}"
            )
            self.write_box(
                f"  Study macros: "
                f"{settings['study_macros'] if settings['use_study_macros'] else 'Skipped'}"
            )
            self.write_box(f"  Program: {program}")
            self.write_box("")

            # Preserve the proven fast execution model:
            # plain sas.submit(), one source file at a time.
            for index, (group, sas_file) in enumerate(
                init_files, start=1
            ):
                self.set_status(
                    f"Initializing {index}/{total}: {sas_file.name}"
                )
                self.write_box(
                    f"[{index}/{total}] [{group}] {sas_file.name} ..."
                )

                source = sas_file.read_text(
                    encoding="utf-8", errors="replace"
                )
                if settings["sync_study"]:
                    source = self.rewrite_study_paths(
                        source, settings["study_root"], remote_study_root
                    )

                t1 = time.perf_counter()
                result = sas.submit(source)
                init_elapsed = time.perf_counter() - t1

                self.write_box(
                    f"    completed in {init_elapsed:.2f} seconds"
                )

                all_log_parts.append(
                    f"\n/* ===== {group}: {sas_file.name} ===== */\n"
                    + flatten_submit_part(result.get("LOG"))
                )

                # autoexec.sas is intentionally written for the local study
                # structure. On SAS OnDemand, explicitly override the derived
                # study paths and permanent librefs after autoexec executes.
                # This is more reliable than relying only on text rewriting.
                if group == "AUTOEXEC" and settings["sync_study"]:
                    bridge_log = self.assign_remote_libraries(sas, remote_study_root)
                    all_log_parts.append(
                        "\n/* ===== SASPYSTUDIO: REMOTE LIBRARY ASSIGNMENT ===== */\n"
                        + bridge_log
                    )
                    self.write_box("    remote RAW/SDTM/ADAM libraries assigned successfully")

            remote_work = ""
            work_before = set()

            if settings["download_outputs"]:
                self.set_status("Preparing output collection...")
                remote_work = self.get_work_path(sas)
                work_before = self.list_remote_files(
                    sas, remote_work
                )

            programs = settings["programs"]
            execution_items = settings["plan_items"]
            final_lst_parts = []
            program_elapsed = 0.0
            completed_items = 0

            for p_index, item in enumerate(execution_items, start=1):
                current_program = item["path"]
                ptype = item["type"]
                kind = item["kind"]
                label = f"{ptype.upper()} {kind.upper()}" if ptype != "manual" else "PROGRAM"
                self.set_status(f"Running {p_index}/{len(execution_items)}: {current_program.name}...")
                self.write_box(f"[{p_index}/{len(execution_items)}] [{label}] {current_program.name} ...")

                item_work_before = set()
                if settings["use_program_plan"] and settings["download_outputs"]:
                    item_work_before = self.list_remote_files(sas, remote_work)

                source = current_program.read_text(encoding="utf-8", errors="replace")
                if settings["sync_study"]:
                    source = self.rewrite_study_paths(source, settings["study_root"], remote_study_root)

                t2 = time.perf_counter()
                result = sas.submit(source)
                this_elapsed = time.perf_counter() - t2
                program_elapsed += this_elapsed
                completed_items += 1

                main_log = flatten_submit_part(result.get("LOG"))
                this_lst = flatten_submit_part(result.get("LST"))
                all_log_parts.append(f"\n/* ===== {label}: {current_program.name} ===== */\n" + main_log)

                # Program-plan runs are routed automatically by type and validation status.
                if settings["use_program_plan"] and settings["study_root"]:
                    if kind == "validation":
                        log_dir = settings["study_root"] / "logs" / "sas" / "validation" / ptype
                        result_dir = settings["study_root"] / "results" / "sas" / "validation" / ptype
                    else:
                        log_dir = settings["study_root"] / "logs" / "sas" / ptype
                        result_dir = settings["study_root"] / "results" / "sas" / ptype
                    log_dir.mkdir(parents=True, exist_ok=True)
                    result_dir.mkdir(parents=True, exist_ok=True)
                    item_log = log_dir / f"{current_program.stem}_{run_id}.log"
                    item_log.write_text(main_log, encoding="utf-8", errors="replace")
                    if this_lst.strip():
                        item_result = result_dir / f"{current_program.stem}_{run_id}.html"
                        item_result.write_text(this_lst, encoding="utf-8", errors="replace")
                elif this_lst.strip():
                    final_lst_parts.append(f"<!-- ===== {current_program.name} ===== -->\n" + this_lst)

                if settings["use_program_plan"] and settings["download_outputs"]:
                    if kind == "validation":
                        item_output_dir = settings["study_root"] / "output" / "sas" / "validation" / ptype
                    else:
                        item_output_dir = settings["study_root"] / "output" / "sas" / ptype
                    item_downloaded = self.download_work_outputs(
                        sas, remote_work, item_work_before, item_output_dir
                    )
                    if item_downloaded:
                        self.write_box(
                            "    outputs: " + ", ".join(p.name for p in item_downloaded)
                        )

                p_errors, _, _ = check_sas_log(main_log)
                self.write_box(f"    completed in {this_elapsed:.2f} seconds")
                if p_errors:
                    self.write_box(f"    ERROR detected in {current_program.name}; run stopped.")
                    break

            final_lst = "\n".join(final_lst_parts)
            elapsed = program_elapsed

            combined_log = "\n".join(all_log_parts)
            log_errors, log_warnings, log_notes = check_sas_log(
                combined_log
            )

            self.after(
                0,
                lambda e=len(log_errors),
                       w=len(log_warnings),
                       n=len(log_notes):
                    self.log_counts.set(
                        f"Errors: {e}   Warnings: {w}   Important Notes: {n}"
                    )
            )

            if log_errors or log_warnings or log_notes:
                self.write_box("")
                self.write_box("LOG CHECK")
                self.write_box("-" * 60)
                for item in log_errors:
                    self.write_box(item)
                for item in log_warnings:
                    self.write_box(item)
                for item in log_notes:
                    self.write_box(item)
                self.write_box("-" * 60)
            else:
                self.write_box("")
                self.write_box(
                    "LOG CHECK: No errors, warnings or important notes found."
                )

            study_downloaded = []
            if settings["sync_study"]:
                self.set_status("Synchronizing study files...")
                study_downloaded = self.sync_remote_study_back(
                    sas,
                    settings["study_root"],
                    remote_study_root,
                )
                self.write_box(
                    f"Study sync: {len(study_downloaded)} file(s) synchronized back locally."
                )

            downloaded = []
            if settings["download_outputs"] and not settings["use_program_plan"]:
                self.set_status("Downloading WORK outputs...")
                downloaded = self.download_work_outputs(
                    sas,
                    remote_work,
                    work_before,
                    run_output_dir,
                )

            log_file.write_text(
                combined_log,
                encoding="utf-8",
                errors="replace",
            )

            if final_lst.strip() and not settings["use_program_plan"]:
                result_file.write_text(
                    final_lst,
                    encoding="utf-8",
                    errors="replace",
                )

            total_elapsed = time.perf_counter() - started

            self.write_box(
                f"Program completed in {elapsed:.2f} seconds."
            )
            self.write_box(
                f"Total run time: {total_elapsed:.2f} seconds."
            )
            self.write_box(f"LOG: {log_file}")

            if final_lst.strip() and not settings["use_program_plan"]:
                self.write_box(f"RESULT: {result_file}")
            else:
                self.write_box("No HTML result was returned.")

            if settings["download_outputs"] and not settings["use_program_plan"]:
                if downloaded:
                    self.write_box(f"OUTPUT: {run_output_dir}")
                    self.write_box(
                        "Downloaded: "
                        + ", ".join(p.name for p in downloaded)
                    )
                else:
                    self.write_box(
                        "OUTPUT: no new WORK output files detected."
                    )

            self.set_status(f"Completed - {completed_items} program(s)")

        except Exception as exc:
            self.set_status("Run failed")
            self.write_box("")
            self.write_box("ERROR:")
            self.write_box(str(exc))

        finally:
            # End SAS only from this worker thread after submit() returns.
            if sas is not None:
                self.set_status("Closing SAS session...")
                try:
                    sas.endsas()
                except Exception as exc:
                    self.write_box(
                        f"SAS close warning: {exc}"
                    )

            self.running = False
            self.after(
                0,
                lambda: self.runbtn.configure(state="normal")
            )

            if self.status.get() == "Closing SAS session...":
                self.set_status("Ready")


if __name__ == "__main__":
    app = SASPyStudio()
    app.mainloop()
