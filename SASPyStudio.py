###-------------------------------------------------------------------------###
### Program Name:   SASPyStudio.py                                          ###
###                                                                         ###                                              
### Application:    SASPy Studio                                            ###                                          
### Version:        1.0                                                     ###
###                                                                         ###
### Purpose:        Provide a lightweight desktop interface for executing   ###
###                 local SAS programs using SAS OnDemand for Academics     ###
###                 through SASPy.                                          ###
###-------------------------------------------------------------------------###
### Programmed By:  Manivannan Mathialagan                                  ###
### Created On:     30Sep2026                                               ###
###-------------------------------------------------------------------------###
### Process:        1. Connect to SAS OnDemand.                             ###
###                 2. Create and synchronize the remote study workspace.   ###
###                 3. Execute the study AUTOEXEC program.                  ###
###                 4. Assign RAW, SDTM and ADAM libraries.                 ###
###                 5. Load optional global and automatic study macros.     ###
###                 6. Execute the selected SAS program/program plan.       ###
###                 7. Display status and log checks in SAS Console.        ###
###                 8. Download generated SAS datasets and outputs.         ###
###                 9. Synchronize outputs to the local study structure.    ###
###                10. End the SAS session.                                 ###
###-------------------------------------------------------------------------###
### Security:       SAS OnDemand credentials are read from the private      ###
###                 _authinfo file configured in sas_config.json.           ###
###                 Authentication files must not be committed to source    ###
###                 control.                                                ###
###-------------------------------------------------------------------------###
### Notes:          Programs Location points to the repository root.        ###
###                 SASPy Studio automatically uses the /sas subfolder for  ###
###                 AUTOEXEC, macros, program_plan.xlsx and SAS programs.   ###
###                                                                         ###
###                 Local SAS programs remain platform independent.         ###
###                 Windows study paths are translated as required when     ###
###                 programs are submitted to SAS OnDemand/Linux.           ###
###-------------------------------------------------------------------------###
### Change History:                                                         ###
### 1.0       30Sep2026    Initial release of SASPy Studio.                 ###
###-------------------------------------------------------------------------###

import json
import os
import threading
import time
from pathlib import Path
import pandas as pd

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



def preferred_ui_font():
    return "Times New Roman"

def preferred_mono_font():
    return "Consolas"

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



class DatasetViewer(tk.Toplevel):
    """Browse synchronized local RAW/SDTM/ADAM SAS7BDAT files."""

    LIBRARIES = ("RAW", "SDTM", "ADAM")

    def __init__(self, master, study_location):
        super().__init__(master)
        self.title("SASPy Studio - Dataset Viewer")
        self.geometry("1180x680")
        self.minsize(900, 520)
        self.configure(bg="#EEF5FC")
        self.ui_font = preferred_ui_font()
        self.mono_font = preferred_mono_font()

        self.study_location = str(study_location or "").strip()
        self.full_df = None
        self.meta_df = None
        self.current_file = ""
        self.var_checks = {}
        self._sort_column = None
        self._sort_ascending = True

        self.card_bg = "#f7fbff"
        self.viewer_bg = "#f3f7fd"

        self.library_var = tk.StringVar(value="SDTM")
        self.dataset_var = tk.StringVar()
        self.filter_var = tk.StringVar()
        self.row_limit_var = tk.StringVar(value="100")
        self.status_var = tk.StringVar(value="Select a library and dataset.")
        self.search_var = tk.StringVar()

        self._build_ui()
        self.refresh_datasets()

    def _library_dir(self):
        # Study Location is the study data/output root used by SASPy Studio.
        return Path(self.study_location) / "data" / "sas" / self.library_var.get().lower()

    def _build_ui(self):
        BG = "#f3f7fd"
        CARD = "#f7fbff"
        TEXT = "#172033"
        MUTED = "#667085"
        BORDER = "#d8e6f3"
        BLUE = "#3b82f6"
        PURPLE = "#2563eb"
        TEAL = "#0F766E"
        GREEN = "#16A34A"
        SLATE = "#9ca3af"

        VIEW_FONT = 12
        VIEW_FONT_BOLD = 12

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Viewer.Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#172033",
            rowheight=30,
            font=(self.ui_font, VIEW_FONT),
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Viewer.Treeview.Heading",
            background="#dceaf8",
            foreground="#173a63",
            font=(self.ui_font, VIEW_FONT_BOLD, "bold"),
            relief="solid",
            borderwidth=1,
            padding=(6, 5),
        )
        style.map(
            "Viewer.Treeview",
            background=[("selected", "#d9ecff")],
            foreground=[("selected", "#102a43")],
        )
        style.map(
            "Viewer.Treeview.Heading",
            background=[("active", "#cfe3f5")],
        )

        top = tk.Frame(self, bg=CARD, bd=1, relief="solid", highlightbackground=BORDER)
        top.pack(fill="x", padx=12, pady=(12, 7))

        tk.Label(top, text="Library", bg=CARD, font=(self.ui_font, VIEW_FONT_BOLD, "bold")).grid(
            row=0, column=0, padx=(10, 5), pady=8, sticky="w"
        )
        lib = ttk.Combobox(
            top, textvariable=self.library_var, values=self.LIBRARIES,
            state="readonly", width=10
        )
        lib.grid(row=0, column=1, padx=(0, 12), pady=8)
        lib.bind("<<ComboboxSelected>>", lambda _e: self.refresh_datasets())

        tk.Label(top, text="Dataset", bg=CARD, font=(self.ui_font, VIEW_FONT_BOLD, "bold")).grid(
            row=0, column=2, padx=(0, 5), pady=8
        )
        self.dataset_combo = ttk.Combobox(
            top, textvariable=self.dataset_var, state="readonly", width=24
        )
        self.dataset_combo.grid(row=0, column=3, padx=(0, 8), pady=8)
        self.dataset_combo.bind("<<ComboboxSelected>>", lambda _e: self.load_dataset())

        tk.Button(top, text="Refresh", command=self.refresh_datasets, bg=PURPLE, fg="white", activebackground="#4a7df2", activeforeground="white", relief="flat", padx=10).grid(
            row=0, column=4, padx=(0, 16), pady=8
        )

        tk.Label(top, text="Rows", bg=CARD, font=(self.ui_font, VIEW_FONT_BOLD, "bold")).grid(
            row=0, column=5, padx=(0, 5), pady=8
        )
        rows = ttk.Combobox(
            top, textvariable=self.row_limit_var,
            values=("50", "100", "500", "1000", "All"),
            state="readonly", width=7
        )
        rows.grid(row=0, column=6, padx=(0, 16), pady=8)
        rows.bind("<<ComboboxSelected>>", lambda _e: self.apply_view())

        tk.Label(top, text="Search", bg=CARD, font=(self.ui_font, VIEW_FONT_BOLD, "bold")).grid(
            row=0, column=7, padx=(0, 5), pady=8
        )
        search = ttk.Entry(top, textvariable=self.search_var, width=24)
        search.grid(row=0, column=8, padx=(0, 8), pady=8, sticky="ew")
        search.bind("<Return>", lambda _e: self.apply_view())
        tk.Button(top, text="Apply", command=self.apply_view, bg=BLUE, fg="white", activebackground="#5c9cff", activeforeground="white", relief="flat", padx=10).grid(
            row=0, column=9, padx=(0, 10), pady=8
        )
        top.grid_columnconfigure(8, weight=1)

        body = tk.PanedWindow(self, orient="horizontal", sashwidth=5, bg="#D6DEE8")
        body.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        # Keep the Variables pane compact so the dataset grid receives
        # most of the available width, including when the viewer is maximized.
        left = tk.Frame(body, bg=CARD, width=250)
        left.pack_propagate(False)
        body.add(left, minsize=210, width=250, stretch="never")

        # Tk can otherwise redistribute pane width during initial layout.
        # Re-assert the preferred sash position once geometry is available.
        self.after_idle(lambda: body.sash_place(0, 250, 0))

        tk.Label(
            left, text="Variables", bg=CARD,
            font=(self.ui_font, VIEW_FONT_BOLD, "bold")
        ).pack(anchor="w", padx=10, pady=(9, 3))

        var_actions = tk.Frame(left, bg=CARD)
        var_actions.pack(fill="x", padx=8, pady=(0, 4))
        tk.Button(var_actions, text="All", width=7, command=self.select_all_vars).pack(side="left")
        tk.Button(var_actions, text="None", width=7, command=self.clear_all_vars).pack(
            side="left", padx=(5, 0)
        )

        canvas = tk.Canvas(left, bg=CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left, orient="vertical", command=canvas.yview)
        self.var_frame = tk.Frame(canvas, bg=CARD)
        self.var_frame.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.var_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
        scrollbar.pack(side="right", fill="y", pady=(0, 6))

        right = tk.Frame(body, bg=CARD)
        body.add(right, stretch="always")

        filterbar = tk.Frame(right, bg=CARD)
        filterbar.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(
            filterbar, text="Filter", bg=CARD,
            font=(self.ui_font, VIEW_FONT_BOLD, "bold")
        ).pack(side="left")
        filter_entry = ttk.Entry(filterbar, textvariable=self.filter_var)
        filter_entry.pack(side="left", fill="x", expand=True, padx=(6, 6))
        filter_entry.bind("<Return>", lambda _e: self.apply_view())
        tk.Button(filterbar, text="Apply Filter", command=self.apply_view, bg=TEAL, fg="white", activebackground="#0D9488", activeforeground="white", relief="flat", padx=10).pack(side="left")
        tk.Button(filterbar, text="Clear", command=self.clear_filters, bg=SLATE, fg="white", activebackground="#b6bcc7", activeforeground="white", relief="flat", padx=10).pack(
            side="left", padx=(5, 0)
        )
        tk.Button(filterbar, text="Metadata", command=self.show_metadata, bg=PURPLE, fg="white", activebackground="#4a7df2", activeforeground="white", relief="flat", padx=10).pack(
            side="left", padx=(10, 0)
        )

        tree_frame = tk.Frame(right, bg=CARD)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 6))

        self.tree = ttk.Treeview(tree_frame, show="headings", style="Viewer.Treeview")
        self.tree.tag_configure("oddrow", background="#ffffff")
        self.tree.tag_configure("evenrow", background="#f5f9fd")
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        tk.Label(
            self, textvariable=self.status_var, anchor="w",
            bg=BG, fg=MUTED, font=(self.ui_font, VIEW_FONT)
        ).pack(fill="x", padx=12, pady=(0, 8))

    def refresh_datasets(self):
        folder = self._library_dir()
        datasets = []
        try:
            if folder.exists():
                datasets = sorted(
                    p.stem.upper()
                    for p in folder.iterdir()
                    if p.is_file() and p.suffix.lower() == ".sas7bdat"
                )
        except Exception as exc:
            messagebox.showerror("Dataset Viewer", f"Unable to scan:\n{folder}\n\n{exc}", parent=self)

        self.dataset_combo["values"] = datasets
        if datasets:
            if self.dataset_var.get() not in datasets:
                self.dataset_var.set(datasets[0])
            self.load_dataset()
        else:
            self.dataset_var.set("")
            self.full_df = None
            self._clear_tree()
            self._build_variable_checks([])
            self.status_var.set(f"No SAS7BDAT datasets found in {folder}")

    def load_dataset(self):
        name = self.dataset_var.get().strip()
        if not name:
            return
        folder = self._library_dir()
        path = folder / f"{name.lower()}.sas7bdat"
        if not path.exists():
            # Preserve actual case if the file was created with upper/mixed case.
            matches = [p for p in folder.glob("*.sas7bdat") if p.stem.upper() == name.upper()]
            if matches:
                path = matches[0]
        try:
            df = pd.read_sas(path, format="sas7bdat", encoding="utf-8")
            # Decode any remaining byte-valued cells safely.
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].map(
                        lambda x: x.decode("utf-8", errors="replace")
                        if isinstance(x, (bytes, bytearray)) else x
                    )
            self.full_df = df
            self.current_file = str(path)
            self._build_variable_checks(list(df.columns))
            self.filter_var.set("")
            self.search_var.set("")
            self._sort_column = None
            self._sort_ascending = True
            self.apply_view()
        except ImportError:
            messagebox.showerror(
                "Dataset Viewer",
                "Pandas SAS support is not available.\n\nInstall pandas and pyreadstat:\n"
                "pip install pandas pyreadstat",
                parent=self
            )
        except Exception as exc:
            messagebox.showerror(
                "Dataset Viewer",
                f"Unable to read SAS dataset:\n{path}\n\n{exc}",
                parent=self
            )

    def _build_variable_checks(self, columns):
        for child in self.var_frame.winfo_children():
            child.destroy()
        self.var_checks = {}
        for col in columns:
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(
                self.var_frame, text=str(col), variable=var,
                command=self.apply_view, bg=self.card_bg, anchor="w",
                activebackground=self.card_bg
            )
            cb.pack(fill="x", anchor="w")
            self.var_checks[str(col)] = var

    def select_all_vars(self):
        for var in self.var_checks.values():
            var.set(True)
        self.apply_view()

    def clear_all_vars(self):
        for var in self.var_checks.values():
            var.set(False)
        self.apply_view()

    def clear_filters(self):
        self.filter_var.set("")
        self.search_var.set("")
        self.apply_view()

    def _selected_columns(self):
        return [name for name, var in self.var_checks.items() if var.get()]

    @staticmethod
    def _format_value(value):
        if pd.isna(value):
            return ""
        if hasattr(value, "strftime"):
            try:
                return value.strftime("%Y-%m-%d")
            except Exception:
                pass
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def _filtered_df(self):
        if self.full_df is None:
            return None

        df = self.full_df
        expression = self.filter_var.get().strip()
        if expression:
            try:
                # Pandas query syntax supports useful expressions such as:
                # AGE > 40, SEX == "F", SITEID == "101".
                df = df.query(expression, engine="python")
            except Exception as exc:
                raise ValueError(
                    "Invalid filter expression.\n\n"
                    'Examples:\nAGE > 40\nSEX == "F"\nSITEID == 101\n\n'
                    f"{exc}"
                )

        search = self.search_var.get().strip().lower()
        if search:
            mask = pd.Series(False, index=df.index)
            for col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(
                    search, na=False, regex=False
                )
            df = df[mask]

        if self._sort_column and self._sort_column in df.columns:
            try:
                df = df.sort_values(
                    self._sort_column,
                    ascending=self._sort_ascending,
                    na_position="last"
                )
            except Exception:
                pass

        return df

    def apply_view(self):
        if self.full_df is None:
            return
        try:
            filtered = self._filtered_df()
        except ValueError as exc:
            messagebox.showwarning("Dataset Filter", str(exc), parent=self)
            return

        selected = self._selected_columns()
        self._clear_tree()
        if not selected:
            self.status_var.set(
                f"{len(filtered):,} of {len(self.full_df):,} observations | 0 variables selected"
            )
            return

        shown = filtered[selected]
        limit_text = self.row_limit_var.get()
        if limit_text != "All":
            try:
                shown = shown.head(int(limit_text))
            except Exception:
                shown = shown.head(100)

        self.tree["columns"] = selected
        for col in selected:
            heading = col
            if self._sort_column == col:
                heading += " ▲" if self._sort_ascending else " ▼"
            self.tree.heading(col, text=heading, command=lambda c=col: self.sort_by(c))
            width = max(85, min(220, max(len(str(col)) * 9, 100)))
            self.tree.column(col, width=width, minwidth=60, stretch=True, anchor="w")

        for row in shown.itertuples(index=False, name=None):
            self.tree.insert("", "end", values=[self._format_value(v) for v in row])

        self.status_var.set(
            f"{self.library_var.get()}.{self.dataset_var.get()}  |  "
            f"Showing {len(shown):,} of {len(filtered):,} filtered / {len(self.full_df):,} total observations  |  "
            f"{len(selected)} of {len(self.full_df.columns)} variables  |  {self.current_file}"
        )

    def sort_by(self, column):
        if self._sort_column == column:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_column = column
            self._sort_ascending = True
        self.apply_view()

    def _clear_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    def show_metadata(self):
        if self.full_df is None:
            return

        win = tk.Toplevel(self)
        win.title(f"Metadata - {self.library_var.get()}.{self.dataset_var.get()}")
        win.geometry("820x480")
        win.configure(bg=self.card_bg)

        columns = ("Order", "Variable", "Type", "Length", "Pandas Type")
        tree = ttk.Treeview(win, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
        tree.column("Order", width=60, anchor="center")
        tree.column("Variable", width=160)
        tree.column("Type", width=90)
        tree.column("Length", width=90, anchor="center")
        tree.column("Pandas Type", width=180)

        yscroll = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        yscroll.pack(side="right", fill="y", padx=(0, 10), pady=10)

        for idx, col in enumerate(self.full_df.columns, start=1):
            series = self.full_df[col]
            dtype = str(series.dtype)
            if pd.api.types.is_numeric_dtype(series):
                sas_type = "Numeric"
                length = "8"
            else:
                sas_type = "Character"
                try:
                    length = str(int(series.dropna().astype(str).map(len).max() or 0))
                except Exception:
                    length = ""
            tree.insert("", "end", values=(idx, col, sas_type, length, dtype))



class SASPyStudio(tk.Tk):
    def __init__(self):
        super().__init__()

        self.ui_font = preferred_ui_font()
        self.mono_font = preferred_mono_font()

        self.title("SASPy Studio")
        self.geometry("1180x700")
        self.minsize(980, 600)

        # Project locations. The finalized repository/study structure is derived from these roots.
        self.study_root = tk.StringVar()
        self.programs_root = tk.StringVar()
        self.global_macros_path = tk.StringVar(value="")
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
        self.use_global_macros = tk.BooleanVar(value=False)
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
        UI_FONT = self.ui_font
        MONO_FONT = self.mono_font

        # Unified typography scale.
        FONT_SECTION = 13
        FONT_LABEL = 12
        FONT_FIELD = 12
        FONT_BUTTON = 12
        FONT_OPTION = 12
        FONT_SMALL = 11
        FONT_CONSOLE = 11

        BG = "#f3f7fd"
        PANEL = "#f7fbff"
        TEXT = "#202124"
        MUTED = "#667085"
        BORDER = "#d8e6f3"
        BLUE = "#3b82f6"
        BLUE_ACTIVE = "#5c9cff"
        PURPLE = "#2563eb"
        PURPLE_ACTIVE = "#4a7df2"
        GREEN = "#16A34A"
        AMBER = "#F59E0B"
        AMBER_ACTIVE = "#D97706"
        RED = "#DC2626"
        RED_ACTIVE = "#B91C1C"
        SOFT = "#e6f2fb"
        SOFT_ACTIVE = "#d9ecff"

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
            font=(UI_FONT, FONT_SECTION, "bold"),
        )
        style.configure(
            "Card.TLabel", background=PANEL, foreground=TEXT,
            font=(UI_FONT, FONT_LABEL)
        )

        class RoundedButton(tk.Canvas):
            def __init__(
                self, parent, text, command, bg, active_bg,
                fg="white", width=150, height=46, radius=10,
                font=(UI_FONT, FONT_BUTTON, "bold")
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
                    self.draw(self.normal_bg if self.enabled else "#d9d9d9")
                if kwargs or cnf:
                    return super().configure(cnf, **kwargs)

            config = configure

        class GreenCheck(tk.Frame):
            """AnnotateCRF-style compact checkbox."""

            def __init__(self, parent, text, variable):
                super().__init__(parent, bg="#f7fbff")
                self.variable = variable
                self.control = tk.Checkbutton(
                    self,
                    text=text,
                    variable=variable,
                    bg="#f7fbff",
                    activebackground="#f7fbff",
                    fg="#27496d",
                    activeforeground="#27496d",
                    selectcolor="#ffffff",
                    font=(UI_FONT, FONT_OPTION),
                    bd=0,
                    relief="flat",
                    highlightthickness=0,
                    padx=2,
                    pady=4,
                    cursor="hand2",
                )
                self.control.pack(side="left")

            def configure(self, cnf=None, **kwargs):
                state = kwargs.pop("state", None)
                if state is not None:
                    self.control.configure(state=state)
                if cnf or kwargs:
                    return super().configure(cnf, **kwargs)

        class ToolTip:
            def __init__(self, widget, text, delay=450):
                self.widget = widget
                self.text = text
                self.delay = delay
                self.tip = None
                self.after_id = None
                widget.bind("<Enter>", self._schedule, add="+")
                widget.bind("<Leave>", self._hide, add="+")
                widget.bind("<ButtonPress>", self._hide, add="+")

            def _schedule(self, _event=None):
                self._cancel()
                self.after_id = self.widget.after(self.delay, self._show)

            def _cancel(self):
                if self.after_id:
                    try:
                        self.widget.after_cancel(self.after_id)
                    except tk.TclError:
                        pass
                    self.after_id = None

            def _show(self):
                if self.tip or not self.text:
                    return
                x = self.widget.winfo_pointerx() + 14
                y = self.widget.winfo_pointery() + 18
                self.tip = tk.Toplevel(self.widget)
                self.tip.wm_overrideredirect(True)
                self.tip.wm_geometry(f"+{x}+{y}")
                tk.Label(
                    self.tip, text=self.text, justify="left",
                    bg="#FFFBEA", fg="#1F2937", relief="solid", bd=1,
                    font=(UI_FONT, FONT_SMALL), padx=7, pady=4
                ).pack()

            def _hide(self, _event=None):
                self._cancel()
                if self.tip:
                    self.tip.destroy()
                    self.tip = None

        def add_tooltip(widget, text):
            ToolTip(widget, text)
            # Composite checkbox controls need hover help on both the box and label.
            for child in widget.winfo_children():
                ToolTip(child, text)

        self._RoundedButton = RoundedButton

        # AnnotateCRF-style rounded title band.
        header = tk.Canvas(
            self,
            height=54,
            bg=BG,
            highlightthickness=0,
            bd=0,
        )
        header.pack(fill="x", padx=12, pady=(10, 8))

        def draw_header_band(_event=None):
            header.delete("all")
            width = max(100, header.winfo_width())
            x1, y1, x2, y2, r = 0, 1, width - 1, 52, 16
            points = [
                x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
                x2, y2-r, x2, y2, x2-r, y2, x1+r, y2,
                x1, y2, x1, y2-r, x1, y1+r, x1, y1
            ]
            header.create_polygon(
                points,
                smooth=True,
                fill="#bfe9f7",
                outline="#bfe9f7",
            )
            cx = width / 2
            header.create_text(
                cx - 48,
                27,
                text="SASPy Studio",
                fill="black",
                font=(UI_FONT, 17, "bold"),
                anchor="e",
            )
            header.create_text(
                cx - 42,
                28,
                text="(SAS OnDemand Programming Studio)",
                fill="black",
                font=(UI_FONT, 12),
                anchor="w",
            )

        header.bind("<Configure>", draw_header_band)

        # ---------- Study setup and execution ----------
        setup_card = ttk.LabelFrame(
            self, text="Study Setup & Execution", style="Card.TLabelframe"
        )
        setup_card.pack(fill="x", padx=20, pady=(0, 5))

        # Responsive two-column setup. The input field itself is clickable,
        # so separate Browse buttons are unnecessary.
        for col in (1, 3):
            setup_card.columnconfigure(col, weight=1, uniform="setup_fields")
        setup_card.columnconfigure(0, weight=0)
        setup_card.columnconfigure(2, weight=0)

        def clickable_field(parent, variable, command, placeholder, icon="", optional=False):
            outer = tk.Frame(
                parent,
                bg="#ffffff",
                highlightbackground="#b8cfe4",
                highlightcolor="#5f9ed1",
                highlightthickness=1,
                bd=0,
                cursor="hand2",
                height=46,
            )
            outer.pack_propagate(False)

            display = tk.Entry(
                outer,
                textvariable=variable,
                relief="flat",
                bd=0,
                bg="#ffffff",
                fg="#12395f",
                readonlybackground="#ffffff",
                font=(UI_FONT, FONT_FIELD, "bold"),
                cursor="hand2",
            )
            display.pack(
                side="left",
                fill="both",
                expand=True,
                padx=(12, 6),
                pady=6,
            )
            display.configure(state="readonly")

            drop = tk.Canvas(
                outer,
                width=30,
                height=44,
                bg="#eef6ff",
                highlightthickness=0,
                bd=0,
                cursor="hand2",
            )
            drop.pack(side="right", fill="y")
            drop.create_line(0, 0, 0, 44, fill="#d4e2ef")
            if optional:
                drop.create_text(
                    15, 22, text="+", fill="#12395f",
                    font=(UI_FONT, FONT_BUTTON, "bold")
                )
            else:
                drop.create_polygon(
                    10, 19, 20, 19, 15, 26,
                    fill="#12395f", outline="#12395f"
                )

            for widget in (outer, display, drop):
                widget.bind("<Button-1>", lambda _e, cmd=command: cmd())

            add_tooltip(outer, placeholder)
            return outer

        # Row 0: primary required roots.
        ttk.Label(
            setup_card, text="Study Location", style="Card.TLabel"
        ).grid(row=0, column=0, padx=(12, 7), pady=(8, 4), sticky="w")

        study_field = clickable_field(
            setup_card,
            self.study_root,
            lambda: self.browse_folder(self.study_root),
            "Select the local study data/output root.",
            "▣",
        )
        study_field.grid(row=0, column=1, padx=(0, 16), pady=(6, 4), sticky="ew")

        ttk.Label(
            setup_card, text="Programs Location", style="Card.TLabel"
        ).grid(row=0, column=2, padx=(0, 7), pady=(8, 4), sticky="w")

        programs_field = clickable_field(
            setup_card,
            self.programs_root,
            lambda: self.browse_folder(self.programs_root),
            "Select the repository root containing the sas folder.",
            "▣",
        )
        programs_field.grid(row=0, column=3, padx=(0, 12), pady=(6, 4), sticky="ew")

        # Row 1: selected program and truly optional global macros.
        ttk.Label(
            setup_card, text="Program", style="Card.TLabel"
        ).grid(row=1, column=0, padx=(12, 7), pady=4, sticky="w")

        program_field = clickable_field(
            setup_card,
            self.program_path,
            self.browse_manual_program,
            "Select one SAS program when Program Plan is not used.",
            "▤",
        )
        program_field.grid(row=1, column=1, padx=(0, 16), pady=4, sticky="ew")

        ttk.Label(
            setup_card, text="Global Macros (Optional)", style="Card.TLabel"
        ).grid(row=1, column=2, padx=(0, 7), pady=4, sticky="w")

        global_field = clickable_field(
            setup_card,
            self.global_macros_path,
            lambda: self.browse_folder(self.global_macros_path),
            "Optional: add a folder of global SAS macros. Leave blank to skip.",
            "＋",
            optional=True,
        )
        global_field.grid(row=1, column=3, padx=(0, 12), pady=4, sticky="ew")

        # ---------- Execution controls ----------
        # Two compact rows keep all actions visible at the normal window size.
        control_wrap = tk.Frame(setup_card, bg=PANEL)
        control_wrap.grid(
            row=2, column=0, columnspan=4,
            padx=10, pady=(5, 6), sticky="ew"
        )

        # Left: plan selection. Right: optional download control.
        option_row = tk.Frame(control_wrap, bg=PANEL)
        option_row.pack(fill="x", pady=(0, 5))

        plan_wrap = tk.Frame(option_row, bg=PANEL)
        plan_wrap.pack(side="left")

        self.plan_check = GreenCheck(plan_wrap, "Use Program Plan", self.use_program_plan)
        self.plan_check.pack(side="left")
        add_tooltip(
            self.plan_check,
            "Run programs from <Programs Location>/sas/program_plan.xlsx instead of the Program field."
        )

        self.plan_display = tk.Label(
            plan_wrap, text="sas/program_plan.xlsx",
            bg=PANEL, fg=MUTED, font=(UI_FONT, FONT_SMALL)
        )
        self.plan_display.pack(side="left", padx=(8, 6))
        add_tooltip(
            self.plan_display,
            "Program plan is automatically derived from <Programs Location>/sas/program_plan.xlsx."
        )

        preview_wrap = tk.Frame(plan_wrap, bg=PANEL)
        preview_wrap.pack(side="left")
        self.previewbtn = RoundedButton(
            preview_wrap, "Preview", self.preview_program_plan,
            "#0F766E", "#0D9488", width=90, height=46, radius=10,
            font=(UI_FONT, FONT_BUTTON, "bold")
        )
        self.previewbtn.pack()
        add_tooltip(
            self.previewbtn,
            "Preview the programs and execution order in sas/program_plan.xlsx."
        )

        self.download_check = GreenCheck(
            option_row, "Download WORK outputs", self.download_outputs
        )
        self.download_check.pack(side="right")
        add_tooltip(
            self.download_check,
            "Download supported files created in remote SAS WORK in addition to permanent study outputs."
        )

        # Action row is deliberately separated so Terminate never gets pushed
        # outside the normal application window.
        action_row = tk.Frame(control_wrap, bg=PANEL)
        action_row.pack(fill="x")

        action_left = tk.Frame(action_row, bg=PANEL)
        action_left.pack(side="left")

        self.runbtn = RoundedButton(
            action_left, "▶  Run SAS", self.start_run,
            "#3b82f6", "#5c9cff", width=145, height=46,
            radius=12, font=(UI_FONT, FONT_BUTTON, "bold")
        )
        self.runbtn.pack(side="left")
        add_tooltip(
            self.runbtn,
            "Run the selected SAS program or the enabled program plan on SAS OnDemand."
        )

        self.restartbtn = RoundedButton(
            action_left, "↻  Restart SAS", self.restart_sas,
            "#F59E0B", "#D97706", fg="#1F2937",
            width=145, height=46, radius=12,
            font=(UI_FONT, FONT_BUTTON, "bold")
        )
        self.restartbtn.pack(side="left", padx=(7, 0))
        add_tooltip(
            self.restartbtn,
            "End the current SAS session so the next run starts with a fresh SAS OnDemand session."
        )

        action_right = tk.Frame(action_row, bg=PANEL)
        action_right.pack(side="right")

        self.viewerbtn = RoundedButton(
            action_right, "Dataset Viewer", self.open_dataset_viewer,
            "#2563eb", "#4a7df2", width=145, height=46,
            radius=12, font=(UI_FONT, FONT_BUTTON, "bold")
        )
        self.viewerbtn.pack(side="left")
        add_tooltip(
            self.viewerbtn,
            "Browse synchronized local RAW, SDTM and ADAM SAS7BDAT datasets without starting SAS."
        )

        self.clearbtn = RoundedButton(
            action_right, "Clear Window", self.clear_execution_window,
            "#9ca3af", "#b6bcc7", fg="white",
            width=135, height=46, radius=12,
            font=(UI_FONT, FONT_BUTTON, "bold")
        )
        self.clearbtn.pack(side="left", padx=(7, 0))
        add_tooltip(
            self.clearbtn,
            "Clear the Run Information and Execution / Log Check display without ending SAS."
        )

        self.closebtn = RoundedButton(
            action_right, "Terminate", self.close_app,
            "#DC2626", "#B91C1C", width=135, height=46,
            radius=12, font=(UI_FONT, FONT_BUTTON, "bold")
        )
        self.closebtn.pack(side="left", padx=(7, 0))
        add_tooltip(
            self.closebtn,
            "Terminate the SAS session and close SASPy Studio."
        )

        # ---------- Execution ----------
        output_card = ttk.LabelFrame(
            self, text="SAS Console", style="Card.TLabelframe"
        )
        output_card.pack(
            fill="both", expand=True, padx=20, pady=(0, 12)
        )

        self.box = ScrolledText(
            output_card,
            height=8,
            font=(MONO_FONT, FONT_CONSOLE),
            bg="#ffffff", fg="#1f2937",
            relief="flat", bd=0
        )
        self.box.pack(fill="both", expand=True, padx=8, pady=6)
        self.box.configure(state="disabled")

        console_footer = tk.Frame(output_card, bg=PANEL)
        console_footer.pack(fill="x", padx=10, pady=(0, 7))

        tk.Label(
            console_footer,
            textvariable=self.status,
            bg=PANEL,
            fg=BLUE,
            font=(UI_FONT, FONT_SMALL, "bold"),
        ).pack(side="left")

        tk.Label(
            console_footer,
            textvariable=self.log_counts,
            bg=PANEL,
            fg=MUTED,
            font=(UI_FONT, FONT_BUTTON, "bold"),
        ).pack(side="right")


    def refresh_derived_paths(self):
        programs_text = self.programs_root.get().strip()
        study_text = self.study_root.get().strip()
        if programs_text:
            repository = Path(programs_text).expanduser()
            programs = repository / "sas"
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
        initial = (Path(programs) / "sas") if programs else APP_DIR
        if current and Path(current).exists():
            initial = Path(current).parent
        filename = filedialog.askopenfilename(
            title="Select SAS Program",
            initialdir=str(initial) if initial.exists() else str(APP_DIR),
            filetypes=[("SAS programs", "*.sas"), ("All files", "*.*")],
        )
        if filename:
            self.program_path.set(filename)

    def open_dataset_viewer(self):
        """Open the local SAS7BDAT dataset browser."""
        study_location = self.study_root.get().strip()
        if not study_location:
            messagebox.showwarning(
                "Dataset Viewer",
                "Select a Study Location first.",
                parent=self
            )
            return
        viewer = DatasetViewer(self, study_location)
        viewer.transient(self)
        viewer.lift()
        viewer.focus_force()

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
            repository_root = self._required_folder(self.programs_root.get(), "Programs location")
            programs_root = self._required_folder(str(repository_root / "sas"), "SAS programs folder")
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


    def cleanup_runtime_config(self):
        """Remove the temporary SASPy runtime configuration file."""
        candidates = {Path(__file__).resolve().parent / "saspy_runtime_cfg.py"}

        for attr in (
            "runtime_cfg_path",
            "runtime_config_path",
            "saspy_runtime_cfg",
            "sas_config_path",
        ):
            value = getattr(self, attr, None)
            if value:
                try:
                    candidates.add(Path(value))
                except TypeError:
                    pass

        for path in candidates:
            try:
                if path.name.lower() == "saspy_runtime_cfg.py" and path.exists():
                    path.unlink()
            except OSError:
                pass

    def close_app(self):
        if self.running:
            messagebox.showwarning(
                "SAS is running",
                "SAS is currently executing. Studio will not force-close the "
                "active IOM connection. Wait for the current submit to return, "
                "then close Studio."
            )
            return
        self.cleanup_runtime_config()
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
    def snapshot_remote_sas_datasets(sas):
        """
        Snapshot permanent RAW/SDTM/ADAM members currently visible to SAS.

        MODATE detects overwritten datasets; NOBS/NVAR are included as
        additional change indicators. The returned keys are (LIBNAME, MEMNAME).
        """
        token = "__SPYDS__="
        sas_code = (
            "proc sql noprint;\n"
            "  create table work._spyds_snapshot as\n"
            "  select upcase(libname) as libname length=8,\n"
            "         upcase(memname) as memname length=32,\n"
            "         modate,\n"
            "         nobs,\n"
            "         nvar\n"
            "  from dictionary.tables\n"
            "  where upcase(libname) in ('RAW','SDTM','ADAM')\n"
            "    and upcase(memtype)='DATA'\n"
            "  order by libname, memname;\n"
            "quit;\n"
            "data _null_;\n"
            "  set work._spyds_snapshot;\n"
            "  length _stamp $40;\n"
            "  _stamp=put(modate,hex16.);\n"
            f'  put "{token}" libname "|" memname "|" _stamp "|" nobs "|" nvar;\n'
            "run;\n"
        )
        result = sas.submit(sas_code)
        log = flatten_submit_part(result.get("LOG"))

        snapshot = {}
        for line in log.splitlines():
            if token not in line:
                continue
            payload = line.split(token, 1)[1].strip()
            parts = [part.strip() for part in payload.split("|")]
            if len(parts) < 5:
                continue
            libname, memname, modate, nobs, nvar = parts[:5]
            snapshot[(libname.upper(), memname.upper())] = (
                modate,
                nobs,
                nvar,
            )
        return snapshot

    @staticmethod
    def changed_remote_sas_datasets(before, after):
        """Return permanent SAS datasets created or changed during this run."""
        changed = []
        for key, state in after.items():
            if key not in before or before.get(key) != state:
                changed.append(key)
        return sorted(changed)

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

    def sync_remote_study_back(
        self,
        sas,
        local_root,
        remote_root,
        changed_datasets=None,
    ):
        """
        Refresh only permanent SAS datasets changed by the current run.

        changed_datasets contains (LIBNAME, MEMNAME) keys such as:
            ("SDTM", "DM")
            ("SDTM", "SUPPDM")

        This avoids re-downloading every RAW/SDTM/ADAM dataset after a
        one-program run.
        """
        downloaded = []
        changed_datasets = set(changed_datasets or [])

        lib_dirs = {
            "RAW": "data/sas/raw",
            "SDTM": "data/sas/sdtm",
            "ADAM": "data/sas/adam",
        }

        for libname, memname in sorted(changed_datasets):
            rel = lib_dirs.get(libname.upper())
            if not rel:
                continue

            remote_dir = remote_root.rstrip("/") + "/" + rel
            local_dir = local_root / Path(rel)
            local_dir.mkdir(parents=True, exist_ok=True)

            # SAS member filenames on the OnDemand Linux host are normally
            # lower-case. Resolve against the actual directory listing so the
            # code remains safe if case differs.
            names = self.list_remote_files(sas, remote_dir)
            target_name = None
            expected = f"{memname}.sas7bdat".lower()
            for name in names:
                if name.lower() == expected:
                    target_name = name
                    break

            if target_name is None:
                self.write_box(
                    f"Study sync warning: changed dataset "
                    f"{libname}.{memname} was not found in {remote_dir}."
                )
                continue

            remote_file = remote_dir.rstrip("/") + "/" + target_name
            local_file = local_dir / target_name.lower()

            try:
                sas.download(
                    str(local_file),
                    remote_file,
                    overwrite=True,
                )
                downloaded.append(local_file)
            except Exception as exc:
                self.write_box(
                    f"Study sync download warning for "
                    f"{libname}.{memname}: {exc}"
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
            repository_root = self._required_folder(self.programs_root.get(), "Programs location")
            programs_root = self._required_folder(str(repository_root / "sas"), "SAS programs folder")
            autoexec = self._required_file(str(programs_root / "autoexec.sas"), "Autoexec")
            global_macros = Path(self.global_macros_path.get().strip()).expanduser() if self.global_macros_path.get().strip() else None
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
                "use_global_macros": global_macros is not None and global_macros.is_dir(),
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

            if settings["use_global_macros"] and settings["global_macros"] is not None:
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

            # Baseline permanent datasets immediately before the selected
            # program(s) execute. This lets us detect only members changed by
            # this run, including overwritten existing datasets.
            permanent_before = {}
            if settings["sync_study"]:
                self.set_status("Snapshotting permanent datasets...")
                permanent_before = self.snapshot_remote_sas_datasets(sas)

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
            changed_datasets = []
            if settings["sync_study"]:
                self.set_status("Checking changed permanent datasets...")
                permanent_after = self.snapshot_remote_sas_datasets(sas)
                changed_datasets = self.changed_remote_sas_datasets(
                    permanent_before,
                    permanent_after,
                )

                if changed_datasets:
                    changed_text = ", ".join(
                        f"{lib}.{mem}" for lib, mem in changed_datasets
                    )
                    self.write_box(
                        f"Changed permanent datasets: {changed_text}"
                    )
                    self.set_status("Refreshing changed datasets locally...")
                    study_downloaded = self.sync_remote_study_back(
                        sas,
                        settings["study_root"],
                        remote_study_root,
                        changed_datasets=changed_datasets,
                    )
                    self.write_box(
                        f"Study sync: {len(study_downloaded)} changed "
                        f"dataset(s) refreshed locally."
                    )
                    if study_downloaded:
                        self.write_box(
                            "Refreshed: "
                            + ", ".join(p.name for p in study_downloaded)
                        )
                else:
                    self.write_box(
                        "Study sync: No permanent SAS datasets changed; "
                        "nothing to refresh locally."
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
    try:
        app.mainloop()
    finally:
        app.cleanup_runtime_config()
