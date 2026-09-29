"""
===============================================================================
SASPy Studio
===============================================================================

A lightweight desktop application for running local SAS programs against
SAS OnDemand for Academics through SASPy.

Author:
    Manivannan Mathialagan

Purpose:
    SASPy Studio provides a simple local development environment for SAS
    programmers who want to use local SAS source files, reusable macros,
    Git/version control, and VS Code while executing programs on a remote
    SAS OnDemand session.

Main Features:
    - Connect to SAS OnDemand through SASPy.
    - Start a fresh SAS session for each program run.
    - Optionally execute autoexec.sas.
    - Automatically load individual SAS macros from the macros/ directory.
    - Run local SAS programs using the SASPy submit() interface.
    - Display initialization and execution progress.
    - Perform basic SAS log checking for errors, warnings, and important notes.
    - Save SAS logs and HTML results locally.
    - Download supported files generated in the remote SAS WORK directory.
    - Keep SAS credentials outside source control using _authinfo.

Execution Model:
    1. Start a fresh SAS OnDemand session.
    2. Submit autoexec.sas, if enabled.
    3. Submit each macros/*.sas file separately, if enabled.
    4. Execute the selected SAS program.
    5. Check the combined SAS log.
    6. Download supported WORK output files, if enabled.
    7. Save local log/results.
    8. End the SAS session.

Security:
    SAS credentials must never be embedded in this source file.
    Authentication is provided through the local _authinfo file.
    _authinfo must remain excluded from Git/source control.

Repository:
    SASPyStudio

Version:
    1.0.0

Created:
    September 2026

===============================================================================
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


APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "sas_config.json"
RUNTIME_CFG_FILE = APP_DIR / ".saspy_runtime_cfg.py"

AUTOEXEC_FILE = APP_DIR / "autoexec.sas"
MACRO_DIR = APP_DIR / "macros"
LOG_DIR = APP_DIR / "logs"
RESULT_DIR = APP_DIR / "results"
OUTPUT_DIR = APP_DIR / "output"

for folder in (LOG_DIR, RESULT_DIR, OUTPUT_DIR):
    folder.mkdir(parents=True, exist_ok=True)


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
        raise ValueError(
            "sas_config.json is missing: " + ", ".join(missing)
        )

    return cfg


def resolve_authinfo_path(cfg):
    """Resolve the private _authinfo path from sas_config.json."""
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

    # SASPy IOM supports authkey, but not an arbitrary authinfo-path key.
    # The session creation code below temporarily points HOME/USERPROFILE to
    # the directory containing our private _authinfo file.
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
    """Original-style SASPy submit() returns LOG/LST as strings."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value

    # Defensive fallback only. The baseline does not request structured logs.
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
    """Conservative checker for actual SAS log messages.

    SOURCE lines are normally suppressed by autoexec (NOSOURCE/NOSOURCE2).
    This intentionally avoids loglines=True so initialization performance
    remains identical to the fast baseline.
    """
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

        self.title("SASPy Studio - Baseline")
        self.geometry("1060x720")
        self.minsize(900, 620)

        self.file = tk.StringVar()
        self.use_autoexec = tk.BooleanVar(value=True)
        self.use_macros = tk.BooleanVar(value=True)
        self.download_outputs = tk.BooleanVar(value=True)

        self.status = tk.StringVar(value="Ready")
        self.info = tk.StringVar(
            value="Fresh SAS session • fast initialization • local logs/results/outputs"
        )
        self.log_counts = tk.StringVar(value="Errors: 0   Warnings: 0   Important Notes: 0")

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
        style.configure("Card.TFrame", background=PANEL)
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
                self.bind("<Enter>", lambda e: self.draw(self.active_bg) if self.enabled else None)
                self.bind("<Leave>", lambda e: self.draw(self.normal_bg) if self.enabled else None)
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
                # Compatibility with existing self.runbtn.configure(state=...)
                state = kwargs.pop("state", None)
                if state is not None:
                    self.enabled = str(state) not in ("disabled", "DISABLED")
                    self.configure_cursor()
                    self.draw(self.normal_bg if self.enabled else "#B8C0CC")
                if kwargs or cnf:
                    return super().configure(cnf, **kwargs)

            config = configure

            def configure_cursor(self):
                super().configure(cursor="hand2" if self.enabled else "arrow")

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
                        2, 2, 18, 18, fill=GREEN,
                        outline=GREEN, width=1
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
        header.pack(fill="x", padx=20, pady=(16, 8))

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

        card = ttk.LabelFrame(
            self, text="SAS Program", style="Card.TLabelframe"
        )
        card.pack(fill="x", padx=20, pady=(0, 9))

        ttk.Label(
            card, text="Program:", style="Card.TLabel"
        ).grid(row=0, column=0, padx=(12, 6), pady=12, sticky="w")

        ttk.Entry(card, textvariable=self.file).grid(
            row=0, column=1, padx=5, pady=12, sticky="ew"
        )

        browse_wrap = tk.Frame(card, bg=PANEL)
        browse_wrap.grid(row=0, column=2, padx=(5, 12), pady=7)
        RoundedButton(
            browse_wrap, "Browse", self.browse,
            PURPLE, PURPLE_ACTIVE, width=92, height=34
        ).pack()

        card.columnconfigure(1, weight=1)

        options = tk.Frame(card, bg=PANEL)
        options.grid(
            row=1, column=1, columnspan=2,
            sticky="w", pady=(0, 10)
        )

        GreenCheck(
            options, "Run autoexec.sas", self.use_autoexec
        ).pack(side="left", padx=(0, 20))
        GreenCheck(
            options, "Load macros folder", self.use_macros
        ).pack(side="left", padx=(0, 20))
        GreenCheck(
            options, "Download WORK outputs", self.download_outputs
        ).pack(side="left")

        toolbar = tk.Frame(
            self, bg=PANEL,
            highlightbackground=BORDER, highlightthickness=1
        )
        toolbar.pack(fill="x", padx=20, pady=(0, 9))

        left = tk.Frame(toolbar, bg=PANEL)
        left.pack(side="left", padx=10, pady=9)

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
        right.pack(side="right", padx=10, pady=9)

        for text_, folder, label, width in (
            ("Clear Logs", LOG_DIR, "logs", 92),
            ("Clear Results", RESULT_DIR, "results", 104),
            ("Clear Outputs", OUTPUT_DIR, "outputs", 108),
        ):
            RoundedButton(
                right, text_,
                lambda f=folder, l=label: self.clear_folder(f, l),
                SOFT, SOFT_ACTIVE, fg="#374151",
                width=width, height=32, radius=9,
                font=("Segoe UI", 8, "bold")
            ).pack(side="left", padx=(5, 0))

        info_card = ttk.LabelFrame(
            self, text="Run Information", style="Card.TLabelframe"
        )
        info_card.pack(fill="x", padx=20, pady=(0, 9))

        self.status_label = tk.Label(
            info_card, textvariable=self.status,
            bg=PANEL, fg=BLUE,
            font=("Segoe UI", 10, "bold")
        )
        self.status_label.pack(anchor="w", padx=12, pady=(9, 2))

        tk.Label(
            info_card, textvariable=self.info,
            bg=PANEL, fg=MUTED, font=("Segoe UI", 9)
        ).pack(anchor="w", padx=12, pady=(0, 2))

        tk.Label(
            info_card, textvariable=self.log_counts,
            bg=PANEL, fg=TEXT, font=("Segoe UI", 9, "bold")
        ).pack(anchor="w", padx=12, pady=(0, 9))

        output_card = ttk.LabelFrame(
            self, text="Execution / Log Check", style="Card.TLabelframe"
        )
        output_card.pack(
            fill="both", expand=True, padx=20, pady=(0, 16)
        )

        self.box = ScrolledText(
            output_card,
            font=("Consolas", 9),
            bg="#FCFCFD", fg="#263238",
            relief="flat", bd=0
        )
        self.box.pack(fill="both", expand=True, padx=8, pady=8)
        self.box.configure(state="disabled")

    def restart_sas(self):
        # Baseline already starts a new SAS session for every Run SAS.
        if self.running:
            messagebox.showinfo(
                "SAS is running",
                "Wait for the current run to finish. The next Run SAS "
                "automatically starts a fresh SAS session."
            )
            return
        self.status.set("Ready - next run starts a fresh SAS session")
        self.write_box("Restart SAS: next run will start a fresh SAS session.")

    def close_app(self):
        # Never call endsas() from a second thread while submit() is active.
        if self.running:
            messagebox.showwarning(
                "SAS is running",
                "SAS is currently executing. To avoid the socket error seen "
                "earlier, Studio will not force-close the active IOM connection. "
                "Wait for the current submit to return, then close Studio."
            )
            return
        self.destroy()

    @staticmethod
    def get_work_path(sas):
        result = sas.submit(
            "%put NOTE: __SPYWORK__=%sysfunc(pathname(work));"
        )
        log = flatten_submit_part(result.get("LOG"))
        for line in log.splitlines():
            if "__SPYWORK__=" in line:
                return line.split("__SPYWORK__=", 1)[1].strip()
        return ""

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

    def browse(self):
        filename = filedialog.askopenfilename(
            title="Select SAS Program",
            initialdir=str(APP_DIR),
            filetypes=[("SAS programs", "*.sas"), ("All files", "*.*")],
        )
        if filename:
            self.file.set(filename)

    def write_box(self, text):
        def update():
            self.box.configure(state="normal")
            self.box.insert("end", text.rstrip() + "\n")
            self.box.see("end")
            self.box.configure(state="disabled")
        self.after(0, update)

    def set_status(self, text):
        self.after(0, lambda: self.status.set(text))

    def clear_folder(self, folder, label):
        if self.running:
            messagebox.showinfo(
                "SAS is running",
                "Wait for the current SAS run to finish before clearing files."
            )
            return

        if not messagebox.askyesno(
            f"Clear {label.title()}",
            f"Delete all files in the {label} folder?"
        ):
            return

        for item in folder.iterdir():
            if item.is_file():
                item.unlink()

        self.status.set(f"Cleared {label}")

    def start_run(self):
        if self.running:
            return

        program = Path(self.file.get().strip())
        if not self.file.get().strip():
            messagebox.showwarning("SAS Program", "Select a SAS program first.")
            return

        if not program.exists():
            messagebox.showerror(
                "SAS Program",
                f"Program not found:\n{program}"
            )
            return

        self.running = True
        self.log_counts.set("Errors: 0   Warnings: 0   Important Notes: 0")
        self.runbtn.configure(state="disabled")

        self.box.configure(state="normal")
        self.box.delete("1.0", "end")
        self.box.configure(state="disabled")

        threading.Thread(
            target=self.run_worker,
            args=(program,),
            daemon=True,
        ).start()

    def run_worker(self, program):
        sas = None
        started = time.perf_counter()

        all_log_parts = []
        final_lst = ""

        run_id = time.strftime("%Y%m%d_%H%M%S")
        log_file = LOG_DIR / f"{program.stem}_{run_id}.log"
        result_file = RESULT_DIR / f"{program.stem}_{run_id}.html"

        try:
            self.set_status("Connecting to SAS OnDemand...")
            self.write_box("Connecting to SAS OnDemand...")

            cfg = load_sas_config()
            cfgname = build_runtime_config(cfg)
            authinfo_path = resolve_authinfo_path(cfg)

            self.write_box(f"Authinfo: {authinfo_path}")

            t0 = time.perf_counter()

            # SASPy's Windows authkey lookup expects _authinfo in the user's
            # home directory. Temporarily point the home lookup to the folder
            # containing our project-local _authinfo while SASsession starts.
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

            init_files = []

            if self.use_autoexec.get():
                if not AUTOEXEC_FILE.exists():
                    raise FileNotFoundError(
                        f"autoexec.sas not found:\n{AUTOEXEC_FILE}"
                    )
                init_files.append(AUTOEXEC_FILE)

            if self.use_macros.get():
                if not MACRO_DIR.exists():
                    raise FileNotFoundError(
                        f"macros folder not found:\n{MACRO_DIR}"
                    )
                init_files.extend(sorted(MACRO_DIR.glob("*.sas")))

            total = len(init_files)

            # IMPORTANT: plain submit(), one file at a time.
            # No loglines=True and no combined startup source.
            for index, sas_file in enumerate(init_files, start=1):
                self.set_status(
                    f"Initializing {index}/{total}: {sas_file.name}"
                )
                self.write_box(
                    f"[{index}/{total}] {sas_file.name} ..."
                )

                source = sas_file.read_text(
                    encoding="utf-8", errors="replace"
                )

                t1 = time.perf_counter()
                result = sas.submit(source)

                elapsed = time.perf_counter() - t1
                self.write_box(
                    f"    completed in {elapsed:.2f} seconds"
                )

                all_log_parts.append(
                    f"\n/* ===== {sas_file.name} ===== */\n"
                    + flatten_submit_part(result.get("LOG"))
                )

            remote_work = ""
            work_before = set()
            run_output_dir = OUTPUT_DIR / run_id

            if self.download_outputs.get():
                self.set_status("Preparing output collection...")
                remote_work = self.get_work_path(sas)
                work_before = self.list_remote_files(sas, remote_work)

            self.set_status(f"Running {program.name}...")
            self.write_box(f"Running {program.name} ...")

            source = program.read_text(
                encoding="utf-8", errors="replace"
            )

            t2 = time.perf_counter()

            # Same simple submit mechanism for the selected program.
            result = sas.submit(source)

            elapsed = time.perf_counter() - t2

            main_log = flatten_submit_part(result.get("LOG"))
            final_lst = flatten_submit_part(result.get("LST"))

            all_log_parts.append(
                f"\n/* ===== {program.name} ===== */\n" + main_log
            )

            # Check the complete run log without changing the fast submit method.
            combined_log = "\n".join(all_log_parts)
            log_errors, log_warnings, log_notes = check_sas_log(combined_log)

            self.after(
                0,
                lambda e=len(log_errors), w=len(log_warnings), n=len(log_notes):
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
                self.write_box("LOG CHECK: No errors, warnings or important notes found.")

            downloaded = []
            if self.download_outputs.get():
                self.set_status("Downloading WORK outputs...")
                downloaded = self.download_work_outputs(
                    sas, remote_work, work_before, run_output_dir
                )

            log_file.write_text(
                "\n".join(all_log_parts),
                encoding="utf-8",
                errors="replace",
            )

            if final_lst.strip():
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

            if final_lst.strip():
                self.write_box(f"RESULT: {result_file}")
            else:
                self.write_box("No HTML result was returned.")

            if self.download_outputs.get():
                if downloaded:
                    self.write_box(f"OUTPUT: {run_output_dir}")
                    self.write_box(
                        "Downloaded: " + ", ".join(p.name for p in downloaded)
                    )
                else:
                    self.write_box("OUTPUT: no new WORK output files detected.")

            self.set_status(
                f"Completed - {program.name}"
            )

        except Exception as exc:
            self.set_status("Run failed")
            self.write_box("")
            self.write_box("ERROR:")
            self.write_box(str(exc))

        finally:
            # End SAS only from this same worker thread, after submit() returns.
            # This avoids the cross-thread socket error seen previously.
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
