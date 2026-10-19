"""
Activity Report Tool

-----------------------
Changes vs v3:
  - [NEW] Real-time date field validation (red border on invalid format)
  - [NEW] Unsaved-changes guard on window close
  - [NEW] Undo stack for last action (Delete / Save / Import / Reset)
  - [NEW] Live search/filter bar above Treeview
  - [NEW] Dynamic template discovery via glob (no hardcoded filename)
  - [NEW] Auto-backup every 10 saves to backups/ folder
  - [NEW] Stale lock detection via PID (auto-clears orphan locks)
  - [NEW] Help → View Log window
  - [NEW] Dark-mode aware colours (follows system theme on Windows/macOS)
  - [NEW] Keyboard shortcuts: Ctrl+S save, Ctrl+Z undo, Del delete row, F5 refresh
  - [NEW] Monthly statistics panel (person-days per user/task)
  - [NEW] Export to .xlsx with formatting and totals
  - [NEW] Next-week planned activities editor (separate window, saved to next_week_plan.json)
  - [NEW] Weekly email includes next-week plan when available
"""

import json
import csv
import os
import sys
import signal
import socket
import logging
import traceback
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
import zipfile
import shutil
from lxml import etree
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
# Logging: basic setup without file handler (added when IO folder is set)
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s %(levelname)s %(funcName)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)
_log_file_handler: logging.FileHandler | None = None

def setup_log_file() -> None:
    """Add file handler once IO folder is known."""
    global _log_file_handler
    if _log_file_handler:
        log.removeHandler(_log_file_handler)
    try:
        _log_file_handler = logging.FileHandler(log_file(), encoding="utf-8")
        _log_file_handler.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)s %(funcName)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        log.addHandler(_log_file_handler)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Configuration / constants
# ---------------------------------------------------------------------------
APP_TITLE         = "Activity Report Tool"
APP_VERSION       = "1.1.2"   # updated by build.bat on each release
# Release folder (VERSION.txt, EXE_LOCATION.txt, exe) — written by build.bat.
# "~" = the user's home, so a OneDrive folder resolves on every PC.
RELEASE_DIR       = ""
UPDATE_CHECK_MS   = 4 * 60 * 60 * 1000   # re-check for a new release while open
SAVE_BACKUP_EVERY = 10

# ---------------------------------------------------------------------------
# All data files live in the user-selected IO folder.
# app_state.json and activity_report.lock stay next to the exe (always local).
# ---------------------------------------------------------------------------

def _exe_dir() -> Path:
    """Folder containing the executable (or script during dev)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

# app_state.json stays local — per-user preferences (last user, io folder path)
APP_STATE_FILE = _exe_dir() / "app_state.json"

# LOCK_FILE lives in the IO folder — shared across users on the network
def lock_file() -> Path:
    return (_io_root / "activity_report.lock") if _io_root else _exe_dir() / "activity_report.lock"

# These live in the IO folder — set via set_io_dir()
_io_root: Path | None = None


def get_io_root() -> Path | None:
    return _io_root


def set_io_dir(io_dir: str) -> None:
    """Set the working IO folder. All data files will be stored here."""
    global _io_root
    if io_dir:
        p = Path(io_dir)
        if p.name.lower() == "weekly_reports":
            p = p.parent
        _io_root = p
        _io_root.mkdir(parents=True, exist_ok=True)
        # Create default config files if missing
        create_default_config()
        reload_holidays()


def require_io_dir() -> bool:
    """Return True if IO folder is set, otherwise show error dialog."""
    if _io_root:
        return True
    from tkinter import messagebox
    messagebox.showerror(
        "No folder selected",
        "Please set the Working folder first.\n\n"
        "Click the 📁 button at the top of the window.",
    )
    return False


def data_file() -> Path:
    return (_io_root / "activities_current.json") if _io_root else _exe_dir() / "activities_current.json"


def weekly_dir() -> Path:
    d = (_io_root / "weekly_reports") if _io_root else _exe_dir() / "weekly_reports"
    d.mkdir(exist_ok=True)
    return d


def backups_dir() -> Path:
    d = (_io_root / "backups") if _io_root else _exe_dir() / "backups"
    d.mkdir(exist_ok=True)
    return d


def next_week_file() -> Path:
    return (_io_root / "next_week_plan.json") if _io_root else _exe_dir() / "next_week_plan.json"


def log_file() -> Path:
    """Log file goes in logs/ subfolder of io_root."""
    if _io_root:
        logs_dir = _io_root / "logs"
        logs_dir.mkdir(exist_ok=True)
        return logs_dir / "activity_report.log"
    return _exe_dir() / "activity_report.log"


def _release_dirs() -> list[Path]:
    """Where to look for a published release: the folder baked in at build
    time first, then the user's working folder (old behaviour)."""
    dirs: list[Path] = []
    if RELEASE_DIR:
        dirs.append(Path(os.path.expanduser(RELEASE_DIR)))
    if _io_root:
        dirs.append(_io_root)
    return dirs


def release_dir() -> Path | None:
    """First candidate folder that actually contains VERSION.txt."""
    for d in _release_dirs():
        try:
            if (d / "VERSION.txt").is_file():
                return d
        except OSError:
            continue
    return None


def _parse_version(v: str) -> tuple:
    try:
        return tuple(int(x) for x in v.strip().lstrip("vV").split("."))
    except ValueError:
        return (0,)


def check_version() -> tuple[str, bool]:
    """Read the published VERSION.txt. Returns (latest_version, update_available).
    Never writes VERSION.txt — that is the build script's job.
    """
    try:
        d = release_dir()
        if d is None:
            log.warning("check_version: VERSION.txt not found in %s", _release_dirs())
            return APP_VERSION, False
        latest = (d / "VERSION.txt").read_text(encoding="utf-8").strip()
        if not latest:
            log.warning("check_version: VERSION.txt is empty in %s", d)
            return APP_VERSION, False
        update_available = _parse_version(latest) > _parse_version(APP_VERSION)
        log.warning("check_version: release=%s (%s) local=%s update=%s",
                    latest, d, APP_VERSION, update_available)
        return latest, update_available
    except Exception as exc:
        log.warning("check_version failed: %s", exc)
        return APP_VERSION, False


def get_latest_exe_path() -> tuple[Path | None, str]:
    """Locate the published exe. Returns (path, error message).
    EXE_LOCATION.txt may hold an absolute path from the build PC
    (C:\\Users\\<builder>\\...), so only its FILE NAME is trusted and it is
    looked up next to EXE_LOCATION.txt, wherever the folder is synced."""
    d = release_dir()
    if d is None:
        return None, "Release folder not found (VERSION.txt missing)."
    loc = d / "EXE_LOCATION.txt"
    try:
        name = Path(loc.read_text(encoding="utf-8").strip().replace("\\", "/")).name
    except OSError:
        return None, f"EXE_LOCATION.txt not found in:\n{d}"
    exe = d / name
    if not exe.is_file():
        return None, (f"{name} is not (yet) available in:\n{d}\n\n"
                      "If the folder is synced by OneDrive, the new version may still "
                      "be downloading — try again in a few minutes.")
    return exe, ""


def auto_update(on_progress=None) -> tuple[bool, str]:
    """Copy the latest exe from share to local, replacing current exe.
    Returns (success, message).
    on_progress(msg) called with status updates.
    """
    import shutil

    def progress(msg):
        if on_progress:
            on_progress(msg)

    # Find latest exe on share
    progress("Locating new version…")
    src, err = get_latest_exe_path()
    if not src:
        return False, err

    # Only works as compiled exe
    if not getattr(sys, "frozen", False):
        return False, "Auto-update only works when running as compiled exe."

    dst = Path(sys.executable)
    try:
        same = src.resolve() == dst.resolve() or os.path.samefile(src, dst)
    except OSError:
        same = False
    if same:
        return False, ("You are running the exe directly from the shared folder.\n"
                       "Copy it to your PC (e.g. Desktop) and start it from there.")
    progress(f"Copying from share…")

    # Copy new exe to a temp file next to current exe
    tmp = dst.with_name(dst.stem + "_new.exe")
    try:
        shutil.copy2(str(src), str(tmp))
    except Exception as exc:
        return False, f"Failed to copy new exe:\n{exc}"

    # Write launcher bat: waits for this process to exit, then renames
    # Does NOT restart automatically — user must reopen manually
    # (avoids DLL-in-use error with --onefile)
    # Retries for up to ~60 s: a --onefile exe can stay locked for a while
    # after the window closes (temp-dir cleanup, antivirus scan).
    launcher = dst.parent / "_update_launcher.bat"
    bat = (
        '@echo off\r\n'
        'chcp 65001 >nul\r\n'
        'set /a n=0\r\n'
        ':retry\r\n'
        'timeout /t 2 /nobreak >nul\r\n'
        f'move /y "{tmp}" "{dst}" >nul 2>&1\r\n'
        'if not errorlevel 1 goto done\r\n'
        'set /a n+=1\r\n'
        'if %n% lss 30 goto retry\r\n'
        'exit /b 1\r\n'
        ':done\r\n'
        'del "%~f0"\r\n'
    )
    import subprocess
    try:
        launcher.write_bytes(bat.encode("utf-8"))
        subprocess.Popen(
            ["cmd", "/c", str(launcher)],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        return False, (f"Could not prepare the update in:\n{dst.parent}\n\n{exc}\n\n"
                       "Copy the new exe by hand from the shared folder.")
    return True, f"ok_close:{dst}"


def config_dir() -> Path:
    d = (_io_root / "config") if _io_root else _exe_dir() / "config"
    d.mkdir(exist_ok=True)
    return d


def users_file() -> Path:
    return config_dir() / "users.txt"


def tasks_file() -> Path:
    return config_dir() / "tasks.txt"


def holidays_file() -> Path:
    return config_dir() / "holidays.txt"


HOLIDAYS_TEMPLATE = """# EUMETSAT Public Holidays
# Format: YYYY-MM-DD | Description
# Lines starting with # are comments

# 2026
2026-01-01 | New Years Day
2026-02-16 | Rosenmontag
2026-04-03 | Good Friday
2026-04-06 | Easter Monday
2026-05-01 | Labour Day
2026-05-14 | Ascension Day
2026-05-15 | Ascension Day Extension
2026-05-25 | Whit Monday
2026-06-04 | Corpus Christi
2026-06-05 | Corpus Christi Extension
2026-10-05 | Republic Day Portugal
2026-10-26 | National Day Austria
2026-12-24 | Christmas Eve
2026-12-25 | Christmas Day
2026-12-31 | New Year\'s Eve
"""


def read_config_text(path: Path) -> str:
    """Read a hand-edited config file: UTF-8 with or without BOM (Notepad),
    falling back to Windows ANSI. A BOM would otherwise turn the first
    comment line into a user, and ANSI accents (e.g. "Nicolò") crashed
    the tool at startup."""
    data = path.read_bytes()
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def load_holidays() -> dict[date, str]:
    """Load holidays from config/holidays.txt. Returns {date: description}."""
    hf = holidays_file()
    if not hf.exists():
        return {}
    result: dict[date, str] = {}
    for line in read_config_text(hf).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split("|", 1)]
        if not parts:
            continue
        try:
            d    = datetime.strptime(parts[0], "%Y-%m-%d").date()
            desc = parts[1] if len(parts) > 1 else ""
            result[d] = desc
        except Exception:
            pass
    return result


# Global holiday cache — reloaded when IO folder changes
_holidays: dict[date, str] = {}


def reload_holidays() -> None:
    global _holidays
    _holidays = load_holidays()


# ---------------------------------------------------------------------------
# Config file templates
# ---------------------------------------------------------------------------

USERS_TEMPLATE = """# List of users — one per line
# Format: First LASTNAME
# Lines starting with # are comments
Bruno MORABITO
Fabio CIALFI
"""

TASKS_TEMPLATE = """
# Task list — one per line
# Format: task_no | Task description | Project | WP | Initials (space-separated)
# Project must be: EPS-Sterna, EPS-Aeolus, Common, or SLM
# WP: WP6, WP7, or SLM (SLM tasks use SLM as both Project and WP)
# Lines starting with # are comments
#
# ── WP6 / EPS-Sterna ────────────────────────────────────────────────────────
1  | Contribution to the closure of any follow-on actions triggered by the EPS-Sterna Ground Segment PDR | EPS-Sterna | WP6 | BrM FaC
2  | Support refining/maintaining the EPS-Sterna MSRD and GSRD VCDs and harmonisation of the lower level VCDs | EPS-Sterna | WP6 | BrM FaC
3  | Support advancing and completion of EPS-Sterna interface or any other documents during detailed design required for GS or System CDR | EPS-Sterna | WP6 | BrM FaC
4  | Contribute to EPS-Sterna RF Suitcase and NDIU documentation and refine the existing ones as necessary | EPS-Sterna | WP6 | BrM FaC
5  | Contribution to EPS-Sterna RF compatibility verification activities and documentation | EPS-Sterna | WP6 | BrM FaC
6  | Complete RF characterisation/compatibility testing based on Arctic Weather Satellite signal recording and replaying and generate all test documentation needed | EPS-Sterna | WP6 | BrM FaC
7  | Contribute to the preparation of test plans and test specifications for EPS-Sterna System Verification activities involving RF suitcase and NDIU | EPS-Sterna | WP6 | BrM FaC
8  | Contribution to EPS-Sterna IV&V campaigns planning and related documentation | EPS-Sterna | WP6 | BrM FaC
#
# ── WP6 / EPS-Aeolus ────────────────────────────────────────────────────────
9  | Maintenance of EPS-Aeolus System and Ground Segment VCDs exported from DOORS/VERDI | EPS-Aeolus | WP6 | BrM FaC
10 | Maintenance of EPS-Aeolus System and GS IV&V Plan, Space-to-Ground IRD and RF Suitcases Requirements Document | EPS-Aeolus | WP6 | BrM FaC
11 | Support the maintenance and update of the EPS-Aeolus ESA SSRD DOORS module | EPS-Aeolus | WP6 | BrM FaC MaB
12 | Support the maintenance and update of the EPS-Aeolus SRD DOORS module including generation of Trace document | EPS-Aeolus | WP6 | BrM FaC
#
# ── WP6 / Common ────────────────────────────────────────────────────────────
13 | Contribution to design activities linked to system modelling and architecture | Common | WP6 | BrM FaC
14 | Contribution to the requirement traceability and maintenance | Common | WP6 | BrM FaC
15 | Contribute to the update and maintenance of system and/or ground segment technical budgets | Common | WP6 | BrM FaC
16 | Participation in any necessary meeting to progress on the design/definition of EPS-Aeolus and EPS-Sterna ground segment, system and/or IV&V activities | Common | WP6 | BrM FaC
17 | Preparation of presentations summarizing outcome of discussions and analysis on the work being done | Common | WP6 | BrM FaC MaB
18 | Support preparation activities for the future project reviews (e.g. SRR-2 of EPS-Aeolus, GS-CDR, System CDR of EPS-Sterna or similar) | Common | WP6 | BrM FaC
19 | Provide Systems Engineering support to the EPS-Aeolus Ground Segment System Engineer in other areas as required | Common | WP6 | BrM FaC
#
# ── WP7 / EPS-Sterna ────────────────────────────────────────────────────────
1  | Contribution to the development and verification of the EPS-Sterna SATSIM and exploitation of the SATSIM prototype as a pre-available tool for the operations preparation activities | EPS-Sterna | WP7 | MaB
2  | Contribution to the preparation of the EPS-Sterna operations preparation documentation as needed (e.g. Flight Operations Procedures (FOP), populating the operation static data (OSD), the operation baseline documentation (OBD) database) | EPS-Sterna | WP7 | MaB
3  | Contribution to the preparation of the EPS-Sterna operational validation documentation as needed (e.g. pre-SVT, SVT, OSV test documentation (specification, procedure, report) | EPS-Sterna | WP7 | MaB
4  | Contribution to the EPS-Sterna LEOP & SIOV planning as necessary | EPS-Sterna | WP7 | MaB
5  | Preparation of presentations summarizing outcome of discussions and analyses on the work being done | EPS-Sterna | WP7 | MaB
6  | Support/Contribute preparation activities for the EPS-Sterna project reviews from operations preparation point of view (e.g. GS-CDR, System CDR or similar) | EPS-Sterna | WP7 | MaB
#
# ── SLM ─────────────────────────────────────────────────────────────────────
1  | Weekly report identifying the progress and status of the activities carried out | SLM | SLM | BrM FaC MaB
2  | Work package report to be submitted together with any invoice | SLM | SLM | BrM FaC MaB
"""


def create_default_config() -> None:
    """Create default config files if they don't exist."""
    uf = users_file()
    if not uf.exists():
        uf.write_text(USERS_TEMPLATE, encoding="utf-8")
    tf = tasks_file()
    if not tf.exists():
        tf.write_text(TASKS_TEMPLATE, encoding="utf-8")
    hf = holidays_file()
    if not hf.exists():
        hf.write_text(HOLIDAYS_TEMPLATE, encoding="utf-8")


def load_users_from_config() -> list[str]:
    """Load user list from config/users.txt."""
    uf = users_file()
    if not uf.exists():
        return []
    users = []
    for line in read_config_text(uf).splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            users.append(line)
    return users


def _looks_like_initials(text: str) -> bool:
    import re
    tokens = text.split()
    return bool(tokens) and all(re.fullmatch(r"[A-Z][a-z][A-Z]", t) for t in tokens)


def load_tasks_from_config() -> tuple[list[tuple], dict[tuple, str], dict[tuple, list[str]], dict[tuple, str]]:
    """Load tasks from config/tasks.txt.
    Format: task_no | description | Project | [WP] | Initials
    WP is optional (e.g. WP6, WP7, SLM). Task numbers can repeat across WPs.
    Returns:
        tasks:        list of (task_no, task_title, wp)
        project_map:  {(task_no, wp): project}
        initials_map: {(task_no, wp): [initials]}
        wp_map:       {(task_no, wp): wp}
    """
    tf = tasks_file()
    if not tf.exists():
        return [], {}, {}, {}
    tasks        = []
    project_map  = {}
    initials_map = {}
    wp_map       = {}
    for line in read_config_text(tf).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 2:
            continue
        try:
            task_no = int(parts[0])
        except ValueError:
            continue
        title   = parts[1] if len(parts) > 1 else ""
        project = parts[2] if len(parts) > 2 else "Common"
        # "N | title | BrM FaC" (project column missing) — don't take initials as project
        if len(parts) == 3 and _looks_like_initials(project):
            parts.append(project)
            project = "Common"
        # parts[3] is WP if it looks like "WP6", "WP7", "SLM" etc.
        # otherwise it is initials
        wp    = ""
        inits = []
        if len(parts) >= 4:
            p3 = parts[3].strip()
            upper = p3.upper()
            if upper.startswith("WP") or upper in ("SLM", "SLM2"):
                wp    = p3
                inits = parts[4].split() if len(parts) > 4 else []
            else:
                # No WP field — parts[3] are initials
                inits = p3.split()
        # Avoid duplicate project field (e.g. "EPS-Sterna | WP6 | BrM" where project==parts[2])
        # already handled above
        key = (task_no, wp)
        tasks.append((task_no, title, wp))
        project_map[key]  = project
        initials_map[key] = inits
        wp_map[key]       = wp
    # Sort: WP first, then task_no
    tasks.sort(key=lambda x: (x[2], x[0]))
    return tasks, project_map, initials_map, wp_map


# Global config state — loaded once IO folder is set
_project_map:    dict[tuple, str]       = {}
_initials_map:   dict[tuple, list[str]] = {}
_wp_map:         dict[tuple, str]       = {}
_users_initials: dict[str, str]         = {}   # user full name -> initials
_task_titles:    dict[tuple, str]       = {}   # (task_no, wp) -> title


def reload_config() -> tuple[list[str], list[tuple]]:
    """Reload users and tasks from config files. Updates global maps."""
    global _project_map, _initials_map, _wp_map, _users_initials, _task_titles
    create_default_config()
    users = load_users_from_config()
    tasks, _project_map, _initials_map, _wp_map = load_tasks_from_config()
    _task_titles = {(n, wp): t for n, t, wp in tasks}
    # Build user->initials map from users.txt names
    # Format: first 2 chars of first name + first char of last name, capitalised
    _users_initials = {}
    for u in users:
        parts = u.replace("\t", " ").split()
        if len(parts) >= 2:
            _users_initials[u] = (parts[0][:2] + parts[-1][:1]).title()
        else:
            _users_initials[u] = u[:3]
    return users, tasks



ENTRY_TYPES = ["Activity", "Holiday", "Sick Leave"]

TEMPLATE_REPLACEMENTS: list[tuple[str, str]] = [
    ("Monthly activity report for April 2026", "title"),
    ("for the month of 30/04/2026",            "for_month"),
    ("EXO_EUM_EPS-Sterna_MAR_04-2026",         "ref"),
    ("EXO_EUM_EPS-Sterna_MAR_03-2026",         "ref"),
    ("April 2026",                             "month_full"),
    ("Apr 26",                                 "month_short"),
    ("30/04/2026",                             "month_end_date"),
]

# ---------------------------------------------------------------------------
# Dark-mode detection
# ---------------------------------------------------------------------------

def _system_is_dark() -> bool:
    try:
        if sys.platform == "win32":
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            )
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return val == 0
        if sys.platform == "darwin":
            import subprocess
            out = subprocess.check_output(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                stderr=subprocess.DEVNULL,
            ).decode().strip()
            return out.lower() == "dark"
    except Exception:
        pass
    return False

DARK = _system_is_dark()

# ---------------------------------------------------------------------------
# Design system — "Aerospace Professional" theme
# Deep navy background, cyan accents, clean typography
# ---------------------------------------------------------------------------

# Core palette (used regardless of OS dark/light setting)
# We enforce our own theme for consistency across all machines
ACCENT       = "#00b4d8"   # cyan accent
ACCENT_DARK  = "#0077a8"   # darker cyan for hover
DANGER       = "#e63946"   # red for destructive actions
DANGER_DARK  = "#c1121f"

if DARK:
    # Dark: deep space navy
    BG_ROOT      = "#0d1117"
    BG_FRAME     = "#161b22"
    BG_PANEL     = "#1c2333"
    BG_ENTRY     = "#21262d"
    BG_HEADER    = "#0d1117"
    FG_MAIN      = "#e6edf3"
    FG_DIM       = "#7d8590"
    FG_ACCENT    = "#00b4d8"
    BORDER       = "#30363d"
    ROW_ODD        = "#161b22"
    ROW_EVEN       = "#1c2333"
    ROW_HOLIDAY    = "#1a2a1a"
    ROW_OTHER_WEEK = "#2a1f0a"
    ROW_SICK     = "#2a1a1e"
    ROW_SEL      = "#1f3a4f"
    SPLASH_BG    = "#0d1117"
    SPLASH_FG    = "#e6edf3"
    SPLASH_DIM   = "#7d8590"
else:
    # Light: cool slate
    BG_ROOT      = "#f4f6f9"
    BG_FRAME     = "#f4f6f9"
    BG_PANEL     = "#ffffff"
    BG_ENTRY     = "#ffffff"
    BG_HEADER    = "#1c2b3a"
    FG_MAIN      = "#1a2332"
    FG_DIM       = "#5a6a7a"
    FG_ACCENT    = "#0077a8"
    BORDER       = "#dce1e8"
    ROW_ODD        = "#f8fafc"
    ROW_EVEN       = "#ffffff"
    ROW_HOLIDAY    = "#fef9e7"
    ROW_SICK       = "#fdf0f2"
    ROW_OTHER_WEEK = "#fff3cd"
    ROW_SEL      = "#e8f4fd"
    SPLASH_BG    = "#1c2b3a"
    SPLASH_FG    = "#e6edf3"
    SPLASH_DIM   = "#7d9db5"

# ---------------------------------------------------------------------------
# Lock file (with PID-based stale detection)
# ---------------------------------------------------------------------------

LOCK_STALE_SECONDS  = 15 * 60          # lock without heartbeat for this long = orphan
LOCK_HEARTBEAT_MS   = 5 * 60 * 1000    # how often a running instance refreshes its lock
_HOST = socket.gethostname()
_TS_FMT = "%Y-%m-%d %H:%M:%S"


def _pid_is_running(pid: int) -> bool:
    """Return True if the given PID is an active process ON THIS MACHINE."""
    try:
        if sys.platform == "win32":
            import ctypes
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000   # required by GetExitCodeProcess
            ERROR_ACCESS_DENIED = 5
            STILL_ACTIVE = 259
            k32 = ctypes.windll.kernel32
            handle = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if not handle:
                # Access denied means the process exists (e.g. another user's)
                return k32.GetLastError() == ERROR_ACCESS_DENIED
            try:
                code = ctypes.c_ulong(0)
                if not k32.GetExitCodeProcess(handle, ctypes.byref(code)):
                    return True   # cannot tell — assume alive (safer)
                return code.value == STILL_ACTIVE
            finally:
                k32.CloseHandle(handle)
        os.kill(pid, 0)
        return True
    except PermissionError:
        return True
    except (OSError, ProcessLookupError):
        return False
    except Exception:
        return False


def _read_lock() -> dict | None:
    try:
        return json.loads(lock_file().read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    except Exception as exc:
        log.warning("lock file unreadable: %s", exc)
        return {}


def _lock_is_ours(info: dict | None) -> bool:
    return bool(info) and info.get("host") == _HOST and info.get("pid") == os.getpid()


def _lock_is_stale(info: dict) -> bool:
    """Same machine: check the PID. Other machine: a PID means nothing here,
    so rely on the heartbeat timestamp the owner refreshes periodically."""
    pid = info.get("pid")
    if info.get("host") == _HOST and pid:
        return not _pid_is_running(int(pid))
    ts = info.get("heartbeat") or info.get("since")
    try:
        age = (datetime.now() - datetime.strptime(ts, _TS_FMT)).total_seconds()
    except Exception:
        return True   # malformed lock
    return age > LOCK_STALE_SECONDS


def _write_lock(since: str | None = None) -> None:
    now = datetime.now().strftime(_TS_FMT)
    info = {
        "user":      os.environ.get("USERNAME") or os.environ.get("USER") or "unknown",
        "host":      _HOST,
        "pid":       os.getpid(),
        "since":     since or now,
        "heartbeat": now,
    }
    lock_file().write_text(json.dumps(info), encoding="utf-8")


def acquire_lock() -> bool:
    # Can only lock once IO folder is set (lock lives on the share)
    if not _io_root:
        return True  # no IO folder yet — let the user through
    info = _read_lock()
    if info is not None and not _lock_is_ours(info):
        if info and not _lock_is_stale(info):
            return False   # held by a live instance
        log.warning("Stale lock detected (%s) — auto-clearing.", info)
    try:
        _write_lock()
        return True
    except Exception as exc:
        log.warning("acquire_lock write failed: %s", exc)
        return True   # on error, let the user through


def refresh_lock() -> bool:
    """Heartbeat: rewrite our lock timestamp. Returns False if someone else
    took the lock over (e.g. Force Unlock from another PC)."""
    if not _io_root:
        return True
    info = _read_lock()
    if info is not None and info and not _lock_is_ours(info):
        return False
    try:
        _write_lock(since=(info or {}).get("since"))
    except Exception as exc:
        log.warning("refresh_lock failed: %s", exc)
    return True


def read_lock_info() -> str:
    info = _read_lock()
    if not info:
        return "another user"
    return (f"{info.get('user', '?')} on {info.get('host', '?')} "
            f"(since {info.get('since', '?')}, last seen {info.get('heartbeat', '?')})")


def release_lock(force: bool = False) -> None:
    """Remove the lock — only if it is ours, unless force=True (Force Unlock)."""
    try:
        if force or _lock_is_ours(_read_lock()):
            lock_file().unlink(missing_ok=True)
    except Exception as exc:
        log.warning("release_lock failed: %s", exc)

# ---------------------------------------------------------------------------
# App state persistence
# ---------------------------------------------------------------------------

def save_app_state(template_path: Path | None, users: list[str], tasks: list[tuple],
                   last_user: str = "",
                   last_io_dir: str = "") -> None:
    try:
        state = {
            "template_path": str(template_path) if template_path else None,
            "last_user":     last_user,
            "last_io_dir":   last_io_dir,
        }
        APP_STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:
        log.warning("save_app_state failed: %s", exc)


def load_app_state() -> tuple[Path | None, list[str], list[tuple], str, str]:
    if not APP_STATE_FILE.exists():
        return None, [], [], "", ""
    try:
        state = json.loads(APP_STATE_FILE.read_text(encoding="utf-8"))
        raw_path = state.get("template_path")
        template_path = Path(raw_path) if raw_path else None
        if template_path and not template_path.exists():
            log.warning("Saved template no longer exists: %s", template_path)
            template_path = None
        last_user   = state.get("last_user", "")
        last_io_dir = state.get("last_io_dir", state.get("last_export_dir", ""))
        return template_path, [], [], last_user, last_io_dir
    except Exception as exc:
        log.warning("load_app_state failed: %s", exc)
        return None, [], [], "", ""

# ---------------------------------------------------------------------------
# Dynamic template discovery
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Exostaff logo / icon (assets/ next to the script, or bundled in the exe)
# ---------------------------------------------------------------------------

_ASSETS_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "assets"
LOGO_HEADER_FILE = "exostaff_logo_header.png"   # 44 px high, window header
LOGO_SPLASH_FILE = "exostaff_logo_splash.png"   # 60 px high, startup splash
ICON_FILE        = "exostaff_icon.ico"
ICON_PNG_FILE    = "exostaff_icon_64.png"


def _asset_path(name: str) -> Path | None:
    path = _ASSETS_DIR / name
    return path if path.is_file() else None


def _load_image(name: str, master=None):
    """An assets/ PNG as a Tk image, or None when missing/unreadable.
    The caller must keep a reference to it, or Tk discards the image."""
    path = _asset_path(name)
    if path is None:
        return None
    try:
        return tk.PhotoImage(file=str(path), master=master)
    except tk.TclError:
        return None


def _template_label(path) -> str:
    """Return header label for the loaded template."""
    if path is None:
        return "\u26a0  No template loaded"
    return "\u2714  Template loaded"


def _find_default_template() -> Path | None:
    """Find the most recent activity-report .docx near the executable."""
    patterns = [
        "EPS-Sterna and EPS-Aeolus System Engineering Service Template.docx",
    ]
    candidates: list[Path] = []
    search_dirs = [Path(sys.executable).parent, Path(".")]
    for d in search_dirs:
        for pat in patterns:
            candidates.extend(d.glob(pat))
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


DEFAULT_TEMPLATE = _find_default_template()

# ---------------------------------------------------------------------------
# Date / file helpers
# ---------------------------------------------------------------------------

def safe_date(s: str) -> date:
    """Parse DD/MM/YY, DD/MM/YYYY or YYYY-MM-DD (and common Excel variants)."""
    for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d-%m-%Y",
                "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Cannot parse date: {s!r}")


def _try_format_date(s: str) -> str:
    """Try to format a stored ISO date as DD/MM/YY for filter matching. Returns '' on error."""
    try:
        return datetime.strptime(s.strip(), "%Y-%m-%d").strftime("%d/%m/%y")
    except Exception:
        return ""

def month_label(d: date) -> str:
    return d.strftime("%b %y")

def full_month_label(d: date) -> str:
    return d.strftime("%B %Y")

def last_day_of_month(year: int, month: int) -> date:
    if month == 12:
        return date(year, 12, 31)
    return date(year, month + 1, 1) - timedelta(days=1)

def iso_week_monday(year: int, cw: int) -> date:
    return date.fromisocalendar(int(year), int(cw), 1)

def cw_range(year: int, cw: int) -> tuple[date, date]:
    start = iso_week_monday(year, cw)
    return start, start + timedelta(days=4)

def iso_year_week_for_date(d: date) -> tuple[int, int]:
    iso = d.isocalendar()
    return iso.year, iso.week

def current_week_file_for_date(d: date) -> Path:
    y, cw = iso_year_week_for_date(d)
    d = weekly_dir(); d.mkdir(exist_ok=True)
    return d / f"week_{y}_CW{cw:02d}.json"

def _stable_key_fn(a: dict) -> tuple:
    """Identify an entry by date + user + type + WP + task_no.
    Used for merge/delete tracking — ignores editable fields like description/effort.
    WP is part of the key because task numbers can repeat across WPs."""
    return (
        str(a.get("date", "")),
        str(a.get("user", "")),
        str(a.get("entry_type", "")),
        str(a.get("task_wp", "") or ""),
        str(a.get("task_no", "")),
    )


def merge_entries(base: list, overrides: list) -> list:
    """Return base + overrides, where overrides REPLACE base entries with the same
    stable key (so an edited entry never coexists with its old version)."""
    keys = {_stable_key_fn(a) for a in overrides}
    return [a for a in base if _stable_key_fn(a) not in keys] + list(overrides)


def activity_key(a: dict) -> tuple:
    a = normalise_activity(a)
    return (
        str(a.get("date", "")),
        str(a.get("user", "")),
        str(a.get("entry_type", "")),
        str(a.get("task_no", "")),
        str(a.get("task_title", "")),
        format_effort(a.get("effort", 0)),   # "1" == "1.0"
        str(a.get("description", "")).strip(),
    )

# ---------------------------------------------------------------------------
# Data normalisation / persistence
# ---------------------------------------------------------------------------

def normalise_activity(a: dict) -> dict:
    item = dict(a)
    if "description" not in item and "desc" in item:
        item["description"] = item.get("desc", "")
    # null / missing text fields (hand-edited JSON) would crash the table
    # refresh — at startup too, making the tool unusable for everybody
    for f in ("date", "user", "task_title", "description", "task_wp"):
        if item.get(f) is None:
            item[f] = ""
        elif not isinstance(item[f], str):
            item[f] = str(item[f])
    if item.get("entry_type") is None:
        item.pop("entry_type", None)
    if item.get("effort") is None:
        item["effort"] = 0
    if "effort" not in item:
        item["effort"] = 1
    # Same formats everywhere: ISO date and numeric effort (files edited in
    # Excel come back as "01/09/2026" and "0,5")
    try:
        item["date"] = safe_date(str(item.get("date", ""))).isoformat()
    except ValueError:
        pass   # left as is; flagged on import
    if isinstance(item["effort"], str):
        try:
            item["effort"] = float(item["effort"].replace(",", "."))
        except ValueError:
            pass
    elif isinstance(item["effort"], int) and not isinstance(item["effort"], bool):
        item["effort"] = float(item["effort"])   # 1 and 1.0 must compare equal
    if "entry_type" not in item:
        item["entry_type"] = item.get("type", "Activity")
    # Normalise entry_type — case-insensitive match against known types
    et = item["entry_type"]
    if et not in ENTRY_TYPES:
        et_lower = et.lower()
        if "holiday" in et_lower or "vacation" in et_lower:
            item["entry_type"] = "Holiday"
        elif "sick" in et_lower or "illness" in et_lower:
            item["entry_type"] = "Sick Leave"
        else:
            item["entry_type"] = "Activity"   # default fallback
    if item["entry_type"] == "Activity":
        if "task_no" not in item or "task_title" not in item:
            task = str(item.get("task", "")).strip()
            if " - " in task:
                no_s, title = task.split(" - ", 1)
                try:
                    item["task_no"] = int(no_s.strip())
                except Exception:
                    item["task_no"] = 0
                item["task_title"] = title.strip()
            else:
                try:
                    item["task_no"] = int(item.get("task_no", 0) or 0)
                except Exception:
                    item["task_no"] = 0
                item["task_title"] = item.get("task_title", task)
    else:
        item["task_no"]    = 0
        item["task_title"] = item["entry_type"]
        if not item.get("description"):
            item["description"] = item["entry_type"]
    return item

def normalise_data(data: list) -> list:
    return [normalise_activity(a) for a in data]

# Set when another user forced the lock away from us: from then on nothing
# is written to the shared folder (viewing and reports still work).
read_only: bool = False
READ_ONLY_MSG = ("The tool is READ-ONLY: another user has taken over the lock.\n"
                 "Close and reopen it when they are done.")


def atomic_write_json(path: Path, obj) -> None:
    """Write JSON via temp file + rename, so a network drop or crash mid-write
    never leaves a truncated file on the share."""
    if read_only:
        raise PermissionError(READ_ONLY_MSG)
    text = json.dumps(obj, indent=2, ensure_ascii=False)
    tmp  = path.with_name(path.name + ".tmp")
    try:
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, path)
    except OSError as exc:
        # e.g. OneDrive/antivirus holding the target — fall back to direct write
        log.warning("atomic write failed for %s (%s) — writing directly", path, exc)
        path.write_text(text, encoding="utf-8")
        tmp.unlink(missing_ok=True)


def write_json(data: list) -> None:
    atomic_write_json(data_file(), normalise_data(data))

# ---------------------------------------------------------------------------
# Table state: which entries were loaded into the table. Persisted next to
# activities_current.json so that deletions and edits are applied to the
# weekly archive even if the user closes without "Save week".
# ---------------------------------------------------------------------------

def table_state_file() -> Path:
    return data_file().with_name("activities_current.state.json")


def save_table_state(loaded_keys: set) -> None:
    try:
        atomic_write_json(table_state_file(),
                          {"loaded_keys": sorted(list(k) for k in loaded_keys)})
    except Exception as exc:
        log.warning("save_table_state failed: %s", exc)


def load_table_state() -> set:
    try:
        raw = json.loads(table_state_file().read_text(encoding="utf-8"))
        return {tuple(str(x) for x in k) for k in raw.get("loaded_keys", [])}
    except FileNotFoundError:
        return set()
    except Exception as exc:
        log.warning("load_table_state failed: %s", exc)
        return set()


def _week_of_key(key: tuple) -> tuple[int, int] | None:
    try:
        return iso_year_week_for_date(safe_date(key[0]))
    except (ValueError, IndexError):
        return None


def merge_week(existing: list, entries: list, loaded_keys: set | frozenset = frozenset()) -> list:
    """Archive content of one week after saving the table's entries for it:
    - table versions replace archived entries with the same key;
    - archived entries that were loaded into the table but are no longer
      there were deleted (or edited into another key/date) and are dropped;
    - everything else (e.g. other users' entries) is kept."""
    table = {_stable_key_fn(a) for a in entries}
    kept = [a for a in existing
            if _stable_key_fn(a) not in table and _stable_key_fn(a) not in loaded_keys]
    return kept + list(entries)


def write_week_archive(year: int, cw: int, entries: list,
                       loaded_keys: set | frozenset = frozenset()) -> tuple[Path, int]:
    """Merge entries into week_<year>_CW<cw>.json. Returns (file, entry count)."""
    week_file = weekly_dir() / f"week_{year}_CW{cw:02d}.json"
    existing: list = []
    if week_file.exists():
        try:
            existing = normalise_data(json.loads(week_file.read_text(encoding="utf-8")))
        except Exception as exc:
            # Do not overwrite an unreadable archive — that would lose its content
            log.error("Could not read existing weekly archive %s: %s", week_file, exc)
            raise RuntimeError(f"Weekly archive {week_file.name} is unreadable; "
                               f"not overwritten. Check the file.") from exc
    merged = normalise_data(merge_week(existing, normalise_data(entries), loaded_keys))
    merged.sort(key=lambda a: (str(a.get("date", "")), str(a.get("user", ""))))
    atomic_write_json(week_file, merged)
    return week_file, len(merged)


def load_week_entries(year: int, cw: int) -> list:
    """Entries archived for one week ([] if none or unreadable)."""
    wf = weekly_dir() / f"week_{year}_CW{cw:02d}.json"
    if not wf.exists():
        return []
    try:
        return normalise_data(json.loads(wf.read_text(encoding="utf-8")))
    except Exception as exc:
        log.warning("load_week_entries %s: %s", wf, exc)
        return []


def archive_is_up_to_date(data: list, loaded_keys: set) -> bool:
    """True if "Save week" would not change any weekly archive file."""
    groups: dict[tuple, list] = {}
    for a in normalise_data(data):
        try:
            groups.setdefault(iso_year_week_for_date(safe_date(a["date"])), []).append(a)
        except ValueError:
            pass
    weeks = set(groups) | {w for w in map(_week_of_key, loaded_keys) if w}
    for y, cw in weeks:
        existing = load_week_entries(y, cw)
        merged = merge_week(existing, groups.get((y, cw), []), loaded_keys)
        if sorted(map(activity_key, merged)) != sorted(map(activity_key, existing)):
            return False
    return True


def save_weekly_archive(data: list, ref_date: date | None = None,
                        loaded_keys: set | frozenset = frozenset()) -> Path | None:
    if ref_date is None:
        ref_date = date.today()
    if not data:
        return None
    return write_week_archive(*iso_year_week_for_date(ref_date), data, loaded_keys)[0]

def reset_if_new_week(data: list) -> tuple[list, bool, Path | None]:
    """Auto-archive past weeks. Groups by week so each lands in its own file."""
    data = normalise_data(data)
    loaded = load_table_state()
    today = date.today()
    current_yw = iso_year_week_for_date(today)

    # Group all entries by calendar week
    groups: dict[tuple, list] = {}
    undated: list = []
    for a in data:
        try:
            key = iso_year_week_for_date(safe_date(a["date"]))
            groups.setdefault(key, []).append(a)
        except Exception:
            undated.append(a)   # kept in the table, never silently dropped

    # Past weeks with entries in the table, or with deletions still to apply
    past_weeks = {k for k in groups if k != current_yw}
    past_weeks |= {w for w in map(_week_of_key, loaded) if w and w != current_yw}
    if not past_weeks:
        return data, False, None

    # Only auto-archive if current week has NO entries in the file.
    # If current week already has entries, the user is working on mixed weeks
    # intentionally — do not auto-archive silently.
    if groups.get(current_yw):
        return data, False, None

    # Genuine week change: archive past weeks (additions, edits and deletions)
    last_saved = None
    for yw_y, yw_cw in sorted(past_weeks):
        last_saved, _ = write_week_archive(yw_y, yw_cw, groups.get((yw_y, yw_cw), []), loaded)

    # Load current week from archive (may already have entries saved earlier)
    y, cw = current_yw
    week_file = weekly_dir() / f"week_{y}_CW{cw:02d}.json"
    current_from_archive: list = []
    if week_file.exists():
        try:
            current_from_archive = normalise_data(
                json.loads(week_file.read_text(encoding="utf-8"))
            )
        except Exception:
            pass

    current_from_archive += undated
    write_json(current_from_archive)
    save_table_state({_stable_key_fn(a) for a in current_from_archive})
    return current_from_archive, True, last_saved

# Set when the data file could not be read — shown to the user by the App
data_load_warning: str | None = None


def read_json() -> tuple[list, bool, Path | None]:
    global data_load_warning
    data_load_warning = None
    if data_file().exists():
        try:
            data = normalise_data(json.loads(data_file().read_text(encoding="utf-8")))
        except Exception as exc:
            log.error("read_json failed: %s", exc)
            # Keep a copy before the next save overwrites the damaged file
            dest = backups_dir() / f"corrupt_{datetime.now():%Y%m%d_%H%M%S}_{data_file().name}"
            try:
                shutil.copy2(data_file(), dest)
            except Exception as copy_exc:
                log.error("could not back up corrupt data file: %s", copy_exc)
                dest = None
            data_load_warning = (
                f"The current activity file could not be read:\n{exc}\n\n"
                + (f"A copy was saved to:\n{dest}\n\n" if dest else "")
                + "Weekly archives are not affected. Use 'Load CW' to reload a week."
            )
            return [], False, None
        try:
            data, reset_done, saved_file = reset_if_new_week(data)
        except Exception as exc:
            log.error("auto-archive failed: %s", exc)
            data_load_warning = f"Automatic weekly archive failed:\n{exc}"
            return data, False, None
        return data, reset_done, saved_file
    return [], False, None

# Weekly files that could not be read by the last monthly load (shown to the user)
skipped_archives: list[str] = []


def load_monthly_data_from_weekly_reports(year: int, month: int) -> list:
    global skipped_archives
    skipped_archives = []
    wd = weekly_dir(); wd.mkdir(exist_ok=True)
    month_start = date(year, month, 1)
    month_end   = last_day_of_month(year, month)
    monthly_data: list = []
    seen: set = set()
    for path in sorted(wd.glob("week_*_CW*.json")):
        try:
            items = normalise_data(json.loads(path.read_text(encoding="utf-8")))
        except Exception as exc:
            log.warning("Skipping unreadable weekly file %s: %s", path, exc)
            skipped_archives.append(path.name)
            continue
        in_file: set = set()
        for a in items:
            try:
                d = safe_date(a["date"])
            except Exception:
                continue
            if month_start <= d <= month_end:
                # Skip the same entry found again in ANOTHER weekly file, but
                # keep identical entries within one week (e.g. 2 x 0.5 PD "Meetings")
                key = activity_key(a)
                if key in seen and key not in in_file:
                    continue
                monthly_data.append(a)
                in_file.add(key)
        seen |= in_file
    return normalise_data(monthly_data)

# ---------------------------------------------------------------------------
# Auto-backup
# ---------------------------------------------------------------------------

_save_counter = 0

def auto_backup(data: list, force: bool = False) -> None:
    """Create a backup of current data.
    Triggers every SAVE_BACKUP_EVERY saves, on first save of session, or when forced.
    """
    global _save_counter
    _save_counter += 1
    # Always backup on first save of session or when forced
    if not force and _save_counter > 1 and _save_counter % SAVE_BACKUP_EVERY != 0:
        return
    try:
        backups_dir().mkdir(exist_ok=True)
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = backups_dir() / f"auto_backup_{ts}.json"
        dest.write_text(json.dumps(normalise_data(data), indent=2, ensure_ascii=False), encoding="utf-8")
        # keep only last 20 backups
        backups = sorted(backups_dir().glob("auto_backup_*.json"))
        for old in backups[:-20]:
            old.unlink(missing_ok=True)
        log.warning("Auto-backup created: %s", dest)
    except Exception as exc:
        log.warning("auto_backup failed: %s", exc)

# ---------------------------------------------------------------------------
# Build replacements
# ---------------------------------------------------------------------------

def build_replacements(year: int, month: int) -> dict[str, str]:
    month_start    = date(year, month, 1)
    month_end      = last_day_of_month(year, month)
    month_full     = full_month_label(month_start)
    month_short    = month_label(month_start)
    month_end_date = month_end.strftime("%d/%m/%Y")
    month_no       = f"{month:02d}"
    ref            = f"EXO_EUM_EPS-Sterna_MAR_{month_no}-{year}"
    return {
        "title":          f"Monthly activity report for {month_full}",
        "for_month":      f"for the month of {month_end_date}",
        "ref":            ref,
        "month_full":     month_full,
        "month_short":    month_short,
        "month_end_date": month_end_date,
    }

# ---------------------------------------------------------------------------
# Word formatting helpers
# ---------------------------------------------------------------------------

def _set_paragraph_center(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pr = paragraph._p.get_or_add_pPr()
    jc   = p_pr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        p_pr.append(jc)
    jc.set(qn("w:val"), "center")

def center_cell_content(cell) -> None:
    try:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    except Exception:
        pass
    try:
        tc_pr   = cell._tc.get_or_add_tcPr()
        v_align = tc_pr.find(qn("w:vAlign"))
        if v_align is None:
            v_align = OxmlElement("w:vAlign")
            tc_pr.append(v_align)
        v_align.set(qn("w:val"), "center")
    except Exception:
        pass
    for paragraph in cell.paragraphs:
        try:
            _set_paragraph_center(paragraph)
        except Exception:
            pass

def set_cell_text(cell, text: str, align_center: bool = True) -> None:
    """Set cell text, always using Calibri 11pt for consistency."""
    from docx.oxml.ns import qn as _qn
    from docx.oxml   import OxmlElement as _OxmlElement
    text = str(text)

    def _apply_calibri(run):
        run.font.name = "Calibri"
        rPr = run._r.get_or_add_rPr()
        for e in rPr.findall(_qn('w:rFonts')):
            rPr.remove(e)
        rf = _OxmlElement('w:rFonts')
        for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rf.set(_qn(attr), "Calibri")
        rPr.insert(0, rf)
        for e in rPr.findall(_qn('w:sz')):
            rPr.remove(e)
        for e in rPr.findall(_qn('w:szCs')):
            rPr.remove(e)
        sz = _OxmlElement('w:sz'); sz.set(_qn('w:val'), '22'); rPr.append(sz)
        sc = _OxmlElement('w:szCs'); sc.set(_qn('w:val'), '22'); rPr.append(sc)

    if cell.paragraphs:
        para = cell.paragraphs[0]
        if para.runs:
            for run in para.runs[1:]:
                run.text = ""
            para.runs[0].text = text
            _apply_calibri(para.runs[0])
        else:
            run = para.add_run(text)
            _apply_calibri(run)
        for extra_para in cell.paragraphs[1:]:
            p = extra_para._element
            p.getparent().remove(p)
    else:
        run = cell.add_paragraph(text).add_run()
        run.text = text
        _apply_calibri(run)
    if align_center:
        center_cell_content(cell)

def _effort_text(value) -> str:
    """Effort for the edit form: "1" not "1.0", but never rounded."""
    try:
        return f"{float(value):g}"
    except (TypeError, ValueError):
        return str(value)


def format_effort(value) -> str:
    """Format effort with 2 decimal places max, no trailing zeros.
    Uses Decimal for precision to avoid float rounding artifacts.
    Examples: 1.0->1, 0.5->0.5, 1.25->1.25, 3.30->3.3"""
    from decimal import Decimal, ROUND_HALF_UP
    try:
        v = float(value)
    except Exception:
        return "0"
    d = Decimal(str(round(v, 10))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if d == d.to_integral_value():
        return str(int(d))
    return format(d, "f").rstrip("0")

# ---------------------------------------------------------------------------
# Template extraction / iteration
# ---------------------------------------------------------------------------

def iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for section in doc.sections:
        for container in (section.header, section.footer):
            for p in container.paragraphs:
                yield p
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p

def iter_all_tables(doc):
    for table in doc.tables:
        yield table
    for section in doc.sections:
        for container in (section.header, section.footer):
            for table in container.tables:
                yield table

def extract_template_data(template_path) -> tuple[list, list]:
    doc   = Document(template_path)
    users: list[str] = []
    for p in iter_all_paragraphs(doc):
        txt = p.text.strip()
        if txt.startswith("Key person:"):
            name = txt.replace("Key person:", "").strip()
            if name and name not in users:
                users.append(name)
    task_map: dict[int, str] = {}
    for table in doc.tables:
        rows = table.rows
        if len(rows) < 5 or len(table.columns) < 3:
            continue
        if any("Tasks Performed" in c.text for r in rows for c in r.cells):
            for r in rows:
                c0 = r.cells[0].text.strip()
                c1 = " ".join(r.cells[1].text.split())
                if c0.isdigit() and c1:
                    task_map[int(c0)] = c1
    tasks = [(n, task_map[n]) for n in sorted(task_map)]
    return users, tasks

def get_user_tables(doc) -> list[tuple]:
    """Return [(user_name, table), ...].
    Strategy 1: look for 'Key person:' paragraphs before each task table (old template).
    Strategy 2: match task tables positionally to users from config (new template).
    """
    # Find all task tables (have "Tasks Performed" row)
    task_tables = [
        t for t in doc.tables
        if any("Tasks Performed" in c.text for r in t.rows for c in r.cells)
    ]
    if not task_tables:
        return []

    # Strategy 1: Key person paragraphs
    user_names: list[str] = []
    for p in iter_all_paragraphs(doc):
        txt = p.text.strip()
        if txt.startswith("Key person:"):
            name = txt.replace("Key person:", "").strip()
            if name:
                user_names.append(name)

    if user_names and len(user_names) == len(task_tables):
        return list(zip(user_names, task_tables))

    # Strategy 2: positional — use users from config in order
    config_users = load_users_from_config()
    if config_users:
        # Match as many as we have tables
        paired = list(zip(config_users[:len(task_tables)], task_tables))
        return paired

    return []

def replace_text_keep_simple(paragraph, replacements: dict) -> None:
    if not paragraph.text:
        return
    new = paragraph.text
    for old, repl in replacements.items():
        new = new.replace(old, repl)
    if new != paragraph.text:
        paragraph.text = new

# ---------------------------------------------------------------------------
# Word month/date patching
# ---------------------------------------------------------------------------

def _template_month(docx_path: Path) -> tuple[int, int] | None:
    """Month the template was made for, from "Monthly activity report for <Month> <Year>".
    Tags are stripped first, so text split across Word runs is still found."""
    import re
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            text = re.sub(r"<[^>]+>", "", z.read(name).decode("utf-8", "ignore"))
            m = re.search(r"Monthly activity report for ([A-Za-z]+) (\d{4})", text)
            if m:
                try:
                    return int(m.group(2)), datetime.strptime(m.group(1), "%B").month
                except ValueError:
                    continue
    return None


def patch_month_and_date_in_docx_file(docx_path, year: int, month: int) -> None:
    docx_path = Path(docx_path)
    repl_map  = build_replacements(year, month)
    replacements = {old: repl_map[key] for old, key in TEMPLATE_REPLACEMENTS}
    # The template may come from any month (e.g. a previous report reused as
    # template): replace THAT month's texts too, not only the April ones.
    tpl_month = _template_month(docx_path)
    if tpl_month and tpl_month != (year, month):
        old_map = build_replacements(*tpl_month)
        for key in ("title", "for_month", "ref", "month_full", "month_short", "month_end_date"):
            replacements.setdefault(old_map[key], repl_map[key])
    new_ref   = repl_map["ref"]
    new_title = repl_map["title"]
    month_end = last_day_of_month(year, month)
    generation_date_obj = date.today()
    generation_date = generation_date_obj.strftime("%d/%m/%Y")
    generation_iso  = generation_date_obj.strftime("%Y-%m-%dT00:00:00Z")
    month_end_iso   = month_end.strftime("%Y-%m-%dT00:00:00")
    month_end_date  = repl_map["month_end_date"]
    ns = {
        "w":     "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "cp":    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "dc":    "http://purl.org/dc/elements/1.1/",
        "vt":    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        "cust":  "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        "cover": "http://schemas.microsoft.com/office/2006/coverPageProps",
    }

    def replace_text(text: str | None) -> str | None:
        if text is None:
            return text
        out = text
        for old, new in replacements.items():
            out = out.replace(old, new)
        if out.startswith("EXO_EUM_EPS-Sterna_MAR_"):
            out = new_ref
        return out

    def patch_xml(name: str, data: bytes) -> bytes:
        try:
            parser = etree.XMLParser(remove_blank_text=False, recover=True)
            root   = etree.fromstring(data, parser=parser)
        except Exception as exc:
            log.warning("XML parse failed for %s, falling back to text replace: %s", name, exc)
            txt = data.decode("utf-8", errors="ignore")
            for old, new in replacements.items():
                txt = txt.replace(old, new)
            return txt.encode("utf-8")
        for elem in root.iter():
            if isinstance(elem.tag, str) and elem.text:
                elem.text = replace_text(elem.text)
        for elem in root.xpath("//*[local-name()='title' or local-name()='lpstr']"):
            if elem.text and "Monthly activity report" in elem.text:
                elem.text = new_title
        for prop in root.xpath("//cust:property", namespaces=ns):
            prop_name = prop.get("name")
            vals = prop.xpath(".//vt:lpwstr | .//vt:lpstr", namespaces=ns)
            if prop_name == "Reference":
                for v in vals: v.text = new_ref
            elif prop_name == "Date completed":
                for v in vals: v.text = month_end_date
        for elem in root.xpath("//cover:PublishDate", namespaces=ns):
            elem.text = month_end_iso
        for sdt in root.xpath("//w:sdt", namespaces=ns):
            alias = sdt.xpath("./w:sdtPr/w:alias/@w:val", namespaces=ns)
            if alias and alias[0] == "Date de publication":
                for d in sdt.xpath("./w:sdtPr/w:date", namespaces=ns):
                    d.set(qn("w:fullDate"), generation_iso)
                texts = sdt.xpath(".//w:sdtContent//w:t", namespaces=ns)
                if texts:
                    texts[0].text = generation_date
                    for t in texts[1:]: t.text = ""
        for fld in root.xpath("//w:fldSimple[contains(@w:instr, 'Date completed')]", namespaces=ns):
            texts = fld.xpath(".//w:t", namespaces=ns)
            if texts:
                texts[0].text = month_end_date
                for t in texts[1:]: t.text = ""
        for fld in root.xpath("//w:fldSimple[contains(@w:instr, 'Reference')]", namespaces=ns):
            texts = fld.xpath(".//w:t", namespaces=ns)
            if texts:
                texts[0].text = new_ref
                for t in texts[1:]: t.text = ""
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=False)

    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") and (
                item.filename.startswith("word/")
                or item.filename.startswith("docProps/")
                or item.filename.startswith("customXml/")
            ):
                data = patch_xml(item.filename, data)
            zout.writestr(item, data)
    shutil.move(str(tmp), str(docx_path))

# ---------------------------------------------------------------------------
# Report generation helpers
# ---------------------------------------------------------------------------

def initials_for_user(name: str) -> str:
    """Return initials for a user — from config if available, else auto-generated."""
    if name in _users_initials:
        return _users_initials[name]
    # Fallback: auto-generate
    parts = name.replace("\t", " ").split()
    if len(parts) >= 2:
        return (parts[0][:2] + parts[-1][:1]).title()
    return name[:3]

def configured_groups() -> list[str]:
    """Weekly summary sections: the WPs of tasks.txt, or its projects if it has no WPs."""
    wps = {wp for (_, wp) in _project_map if wp}
    return sorted(wps) if wps else sorted(set(_project_map.values()))


def entry_wp(a: dict) -> str:
    """WP of an entry. Entries created before the WP column existed, or
    imported from CSV, have no task_wp: recover it from config/tasks.txt
    (exact title first, then task number if it is unique across WPs)."""
    wp = a.get("task_wp", "") or ""
    if wp or a.get("entry_type", "Activity") != "Activity":
        return wp
    title = _norm_title(a.get("task_title", ""))
    if title:
        by_title = {k[1] for k in _project_map if _norm_title(_task_titles.get(k, "")) == title}
        if len(by_title) == 1:
            return by_title.pop()
    try:
        n = int(a.get("task_no", 0) or 0)
    except Exception:
        return ""
    by_no = {k[1] for k in _project_map if k[0] == n}
    return by_no.pop() if len(by_no) == 1 else ""


def project_for_task(task_no, wp: str = "") -> str:
    try:
        n = int(task_no)
    except Exception:
        return "Common"
    return _project_map.get((n, wp), _project_map.get((n, ""), "Common"))

def _clean_bullet(text: str) -> str:
    text = " ".join(str(text or "").split())
    return text.rstrip(".;") + "." if text else ""

def _weekly_items_for_period(data: list, start: date, end: date) -> dict[str, list]:
    """Group activity descriptions by WP (or project if no WP).
    Returns {wp_or_project: [description, ...]} deduplicated."""
    grouped: dict[str, list] = {}
    for a in normalise_data(data):
        if a.get("entry_type") != "Activity":
            continue
        try:
            d = safe_date(a["date"])
        except Exception:
            continue
        if not (start <= d <= end):
            continue
        # No description: fall back to the task title rather than dropping the entry
        desc = _clean_bullet(a.get("description", "") or a.get("task_title", ""))
        if not desc:
            continue
        wp      = entry_wp(a)
        project = project_for_task(int(a.get("task_no", 0) or 0), wp)
        # Group key: WP if present, else project
        group_key = wp if wp else project
        grouped.setdefault(group_key, []).append(desc)
    out: dict[str, list] = {}
    for key, items in grouped.items():
        seen:   set  = set()
        unique: list = []
        for desc in items:
            k = desc.lower()
            if k not in seen:
                seen.add(k)
                unique.append(desc)
        out[key] = unique
    return out

# ---------------------------------------------------------------------------
# Next-week plan persistence
# ---------------------------------------------------------------------------

def week_after(year: int, cw: int) -> tuple[int, int]:
    return iso_year_week_for_date(iso_week_monday(year, cw) + timedelta(days=7))


def load_next_week_plan(target: tuple[int, int] | None = None) -> dict[str, list[str]]:
    """Returns {group: [bullet, ...]} or empty dict.
    The file records the week it plans for; with `target`, a plan written
    for another week is ignored so an old plan never reappears as "Next week".
    Files from older versions have no week and are returned as they are."""
    if not next_week_file().exists():
        return {}
    try:
        raw = json.loads(next_week_file().read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning("load_next_week_plan failed: %s", exc)
        return {}
    if isinstance(raw, dict) and "plan" in raw and "target_week" in raw:
        if target is not None and tuple(raw["target_week"]) != tuple(target):
            return {}
        return raw["plan"]
    return raw


def save_next_week_plan(plan: dict[str, list[str]], target: tuple[int, int]) -> None:
    try:
        atomic_write_json(next_week_file(), {"target_week": list(target), "plan": plan})
    except Exception as exc:
        log.warning("save_next_week_plan failed: %s", exc)

# ---------------------------------------------------------------------------
# Weekly text generation (with optional next-week plan override)
# ---------------------------------------------------------------------------

def generate_weekly_text(
    data: list,
    start_s: str,
    end_s: str,
    cw=None,
    year=None,
    next_week_plan: dict | None = None,
) -> str:
    start        = safe_date(start_s)
    end          = safe_date(end_s)
    achievements = _weekly_items_for_period(data, start, end)

    # Next-week items: prefer explicit plan, fall back to data
    if next_week_plan:
        next_steps = {k: [_clean_bullet(b) for b in v if b.strip()]
                      for k, v in next_week_plan.items()}
    else:
        next_start = start + timedelta(days=7)
        next_end   = end   + timedelta(days=7)
        next_steps = _weekly_items_for_period(data, next_start, next_end)

    def add_section(lines, key, items, empty_label):
        # Skip empty Common sections
        if not items and "common" in key.lower():
            return
        lines.append(f"{key}:")
        if items:
            for desc in items:
                lines.append(f"o\t{desc}")
        else:
            lines.append(f"\t{empty_label}")

    lines: list[str] = []
    if year and cw:
        lines.append(f"CW{int(cw):02d} / {year}")
        lines.append("")
    lines.append("Current Week")
    lines.append("")
    # Every configured group is listed (an empty WP says "No activities this
    # week" instead of disappearing), plus any group found in the data.
    order = lambda k: (0 if k.upper().startswith("WP") else 1 if k.upper() == "SLM" else 2, k)
    all_keys = sorted(set(configured_groups()) | set(achievements), key=order)
    if not all_keys:
        all_keys = ["EPS-Sterna", "EPS-Aeolus", "Common"]
    for key in all_keys:
        add_section(lines, key, achievements.get(key, []), "No activities this week.")
    lines.append("")
    lines.append("Next week")
    lines.append("")
    next_keys = sorted(set(configured_groups()) | set(next_steps), key=order) or all_keys
    for key in next_keys:
        add_section(lines, key, next_steps.get(key, []), "To be defined.")
    return "\n".join(lines).strip()

# ---------------------------------------------------------------------------
# Monthly Word generation
# ---------------------------------------------------------------------------

def _norm_title(text: str) -> str:
    """Lower-case, collapse whitespace and drop trailing punctuation for title matching."""
    return " ".join(str(text or "").lower().split()).rstrip(" .;:")


def _match_template_row(entry: dict, task_rows: list[tuple], table_wp: str) -> int | None:
    """Return the row index of the template task row an entry belongs to.
    task_rows: [(row_idx, task_no, normalised_title), ...]
    Title match wins over number match, because task numbers in tasks.txt
    (which may restart per WP) do not have to follow the template numbering.
    """
    title = _norm_title(entry.get("task_title", ""))
    n     = int(entry.get("task_no", 0) or 0)
    wp    = entry.get("task_wp", "")
    if title:
        exact = [r for r in task_rows if r[2] == title]
        cont  = [r for r in task_rows if r[2] and (title in r[2] or r[2] in title)]
        for cands in (exact, cont):
            if len(cands) == 1:
                return cands[0][0]
            same_no = [r for r in cands if r[1] == n]
            if same_no:
                return same_no[0][0]
        # Near-identical wording (e.g. template adds "(e.g. ...)")
        import difflib
        scored = sorted(((difflib.SequenceMatcher(None, title, r[2]).ratio(), r[0])
                         for r in task_rows if r[2]), reverse=True)
        if scored and scored[0][0] >= 0.85 and (len(scored) == 1 or scored[1][0] < scored[0][0] - 0.05):
            return scored[0][1]
    # Number fallback only when the numbering scheme can be trusted
    if not wp or not table_wp or wp == table_wp:
        for r in task_rows:
            if r[1] == n:
                return r[0]
    return None


def generate_monthly_docx(template_path, data: list, year: int, month: int, out_path: str) -> list[str]:
    """Fill the Word template. Returns warnings for entries that could not be placed."""
    # Update month/date text on a copy of the TEMPLATE first: patching the
    # finished document would also rewrite users' descriptions
    # (e.g. "the April 2026 review" -> "the September 2026 review").
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        patched = Path(td) / "template.docx"
        shutil.copy2(template_path, patched)
        patch_month_and_date_in_docx_file(patched, year, month)
        doc = Document(str(patched))
    month_start = date(year, month, 1)
    month_end   = last_day_of_month(year, month)
    user_tables = get_user_tables(doc)

    # Verify names in template match users in config
    template_users = [u for u, _ in user_tables]
    config_users   = list(_users_initials.keys()) if _users_initials else []
    if config_users:
        missing_in_template = [u for u in config_users if u not in template_users]
        missing_in_config   = [u for u in template_users if u not in config_users]
        warnings = []
        if missing_in_template:
            warnings.append(f"In config but NOT in template: {', '.join(missing_in_template)}")
        if missing_in_config:
            warnings.append(f"In template but NOT in config: {', '.join(missing_in_config)}")
        if warnings:
            log.warning("User mismatch between config and template: %s", "; ".join(warnings))
            # Non-blocking — continue with what the template has

    activities_by_user_task: dict = defaultdict(list)
    leave_by_user: dict = defaultdict(float)
    unmapped: list[str] = []
    # Match entry users to template names ignoring case and spacing
    # ("Fabio Cialfi" in users.txt vs "Fabio CIALFI" in the template)
    _norm_user = lambda u: " ".join(str(u).split()).casefold()
    tpl_name = {_norm_user(u): u for u in template_users}
    canon = lambda u: tpl_name.get(_norm_user(u), u)
    for a in normalise_data(data):
        a = dict(a, user=canon(a.get("user", "")))
        try:
            d = safe_date(a["date"])
        except Exception:
            continue
        if not (month_start <= d <= month_end):
            continue
        et = a.get("entry_type", "Activity")
        if et == "Activity":
            task_no = int(a.get("task_no", 0) or 0)
            task_wp = a.get("task_wp", "")
            if task_no > 0:
                activities_by_user_task[(a.get("user", ""), task_no, task_wp)].append(a)
            else:
                unmapped.append(f"{a.get('user', '')}: {a.get('date', '')} activity without task "
                                f"({format_effort(a.get('effort', 0))} PD)")
        elif et in ("Holiday", "Sick Leave"):
            # FIX C: only count genuine leave entries (not malformed Activities)
            leave_by_user[a.get("user", "")] += float(a.get("effort", 0) or 0)
        # Other entry_types silently ignored
    # Users with activity but no table in the template: their effort would vanish
    for u in sorted({u for (u, _, _) in activities_by_user_task} - set(template_users)):
        pd = sum(float(a.get("effort", 0) or 0)
                 for (uu, _, _), acts in activities_by_user_task.items() if uu == u for a in acts)
        unmapped.append(f"{u or '(no user)'}: {format_effort(pd)} PD — user not in template")
    for u in sorted(set(leave_by_user) - set(template_users)):
        if leave_by_user[u]:
            unmapped.append(f"{u or '(no user)'}: {format_effort(leave_by_user[u])} days of leave "
                            "— user not in template")
    # Pre-compute total effort per user (summing across all tasks and WPs)
    effort_by_user: dict[str, float] = {}
    for (u, tn, wp), acts in activities_by_user_task.items():
        effort_by_user[u] = effort_by_user.get(u, 0.0) + sum(
            float(a.get("effort", 0) or 0) for a in acts
        )

    user_efforts: dict[str, float] = {}
    for user, table in user_tables:
        total_effort = 0.0
        # Detect which WP this table belongs to from header row
        table_wp = ""
        for r in table.rows[:3]:
            for c in r.cells:
                txt = c.text.strip()
                if txt.upper().startswith("WP-") or txt.upper().startswith("WP"):
                    table_wp = txt.replace("-", "")  # "WP-06" -> "WP06", "WP6"
                    # Normalise: "WP06" -> "WP6"
                    try:
                        table_wp = "WP" + str(int(table_wp[2:]))
                    except Exception:
                        pass
                    break
            if table_wp:
                break

        rows = list(table.rows)
        task_rows = [
            (idx, int(r.cells[0].text.strip()), _norm_title(r.cells[1].text))
            for idx, r in enumerate(rows) if r.cells[0].text.strip().isdigit()
        ]
        # Assign every entry of this user to exactly ONE template row, so that
        # effort can never be counted twice (or silently dropped).
        acts_by_row: dict[int, list] = defaultdict(list)
        for (u, n, wp), entries in activities_by_user_task.items():
            if u != user:
                continue
            row_idx = _match_template_row(entries[0], task_rows, table_wp)
            if row_idx is None:
                pd = sum(float(a.get("effort", 0) or 0) for a in entries)
                unmapped.append(f"{user}: {wp + ' ' if wp else ''}task {n} "
                                f"({format_effort(pd)} PD) — not found in template")
                continue
            acts_by_row[row_idx].extend(entries)

        for idx, r in enumerate(rows):
            c0 = r.cells[0].text.strip()
            if c0.isdigit():
                acts    = acts_by_row.get(idx, [])
                effort  = sum(float(a.get("effort", 0) or 0) for a in acts)
                total_effort += effort
                set_cell_text(r.cells[2], format_effort(effort))
                if idx + 1 < len(rows):
                    desc_cell = rows[idx + 1].cells[1]
                    if effort > 0 and acts:
                        # Deduplicate descriptions preserving order
                        seen_desc: set = set()
                        unique_descs: list = []
                        for a in acts:
                            d = (a.get("description") or "").strip()
                            if d and d.lower() not in seen_desc:
                                seen_desc.add(d.lower())
                                unique_descs.append(d)
                        text = "\n".join(f"• {d}" for d in unique_descs) if unique_descs else "No activity in the reporting period"
                    else:
                        text = "No activity in the reporting period"
                    set_cell_text(desc_cell, text)
        user_efforts[user] = total_effort
        for r in table.rows:
            label = r.cells[1].text.strip()
            if label == "MONTH":
                set_cell_text(r.cells[2], month_label(month_start))
            if "EFFORTS ALLOCATED" in label:
                set_cell_text(r.cells[2], format_effort(total_effort))
    for table in doc.tables:
        if not table.rows:
            continue
        first_text = " ".join(cell.text for cell in table.rows[0].cells)
        def _match_user_by_initials(code: str) -> str | None:
            """Find user whose initials match code (case-insensitive)."""
            code_lower = code.lower()
            # Try exact match first
            for u in user_efforts:
                if initials_for_user(u) == code:
                    return u
            # Case-insensitive fallback
            for u in user_efforts:
                if initials_for_user(u).lower() == code_lower:
                    return u
            # Try matching against full name fragments
            for u in user_efforts:
                parts = u.split()
                if len(parts) >= 2:
                    auto = (parts[0][:2] + parts[-1][:1]).title()
                    if auto.lower() == code_lower:
                        return u
            return None

        if "Total effort" in first_text:
            total = 0.0
            for row in table.rows[1:]:
                code = row.cells[0].text.strip()
                if not code or code.upper() == "TOTAL":
                    continue
                matched = _match_user_by_initials(code)
                if matched:
                    val = user_efforts[matched]; total += val
                    set_cell_text(row.cells[2], format_effort(val))
                else:
                    log.warning("No user match for initials code: %r — available: %s",
                                code, [initials_for_user(u) for u in user_efforts])
            for row in table.rows:
                if row.cells[0].text.strip().upper() == "TOTAL":
                    set_cell_text(row.cells[2], format_effort(total))

        if "Total absence" in first_text:
            total_abs = 0.0
            for row in table.rows[1:]:
                code = row.cells[0].text.strip()
                if not code or code.upper() == "TOTAL":
                    continue
                matched = _match_user_by_initials(code)
                if matched:
                    val = leave_by_user.get(matched, 0.0); total_abs += val
                    set_cell_text(row.cells[2], format_effort(val))
            for row in table.rows:
                if row.cells[0].text.strip().upper() == "TOTAL":
                    set_cell_text(row.cells[2], format_effort(total_abs))
    for table in iter_all_tables(doc):
        for row in table.rows:
            for cell in row.cells:
                center_cell_content(cell)
    doc.save(out_path)
    for msg in unmapped:
        log.warning("monthly: %s", msg)
    return unmapped

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def export_to_xlsx(data: list, out_path: str) -> None:
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl is not installed. Run: pip install openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Activities"

    header_fill = PatternFill("solid", fgColor="1E3A5F")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill("solid", fgColor="F0F4FA")

    headers = ["Date", "User", "Type", "Task No", "Task Title", "Effort (PD)", "Description"]
    col_widths = [13, 22, 14, 9, 40, 12, 60]
    ws.row_dimensions[1].height = 18
    for col_idx, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font   = header_font
        cell.fill   = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    data_norm = normalise_data(data)
    for row_idx, a in enumerate(data_norm, 2):
        fill = alt_fill if row_idx % 2 == 0 else PatternFill()
        values = [
            a.get("date", ""),
            a.get("user", ""),
            a.get("entry_type", ""),
            a.get("task_no", ""),
            a.get("task_title", ""),
            a.get("effort", ""),
            a.get("description", ""),
        ]
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border
            cell.fill   = fill
            cell.font   = Font(name="Calibri", size=10)
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=(col_idx == 7),
                horizontal="center" if col_idx in (1, 3, 4, 6) else "left",
            )

    # Totals row
    total_row = len(data_norm) + 2
    ws.cell(row=total_row, column=5, value="TOTAL").font = Font(bold=True, name="Calibri", size=10)
    total_effort = sum(float(a.get("effort", 0) or 0) for a in data_norm)
    tc = ws.cell(row=total_row, column=6, value=total_effort)
    tc.font   = Font(bold=True, name="Calibri", size=10)
    tc.fill   = PatternFill("solid", fgColor="D6E4F0")
    tc.border = border
    tc.alignment = Alignment(horizontal="center")

    # Per-user summary sheet
    ws2 = wb.create_sheet("Summary by User")
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 14
    ws2.cell(row=1, column=1, value="User").font         = header_font
    ws2.cell(row=1, column=1).fill                       = header_fill
    ws2.cell(row=1, column=2, value="Total PD").font     = header_font
    ws2.cell(row=1, column=2).fill                       = header_fill
    by_user: dict[str, float] = defaultdict(float)
    for a in data_norm:
        if a.get("entry_type") == "Activity":
            by_user[a.get("user", "")] += float(a.get("effort", 0) or 0)
    for r, (user, pd) in enumerate(sorted(by_user.items()), 2):
        ws2.cell(row=r, column=1, value=user)
        ws2.cell(row=r, column=2, value=pd).alignment = Alignment(horizontal="center")

    wb.save(out_path)

# ---------------------------------------------------------------------------
# Import helpers
# ---------------------------------------------------------------------------

def import_activity_table(path: str | Path) -> tuple[list, list[str]]:
    """Returns (valid entries, descriptions of rejected rows)."""
    path = Path(path)
    if path.suffix.lower() == ".json":
        return _validate_imported(normalise_data(json.loads(path.read_text(encoding="utf-8"))))
    rows: list = []
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        sample = f.read(4096); f.seek(0)
        delimiter = ";" if sample.count(";") >= sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delimiter)
        for row in reader:
            if not row:
                continue
            item = {
                "date":        row.get("date", ""),
                "user":        row.get("user", ""),
                "entry_type":  row.get("entry_type", row.get("type", "Activity")),
                "task_no":     row.get("task_no", 0) or 0,
                "task_title":  row.get("task_title", ""),
                "task_wp":     (row.get("task_wp") or "").strip(),
                "effort":      row.get("effort", 0) or 0,
                "description": row.get("description", row.get("desc", "")),
            }
            try:
                item["task_no"] = int(float(str(item["task_no"]).replace(",", ".")))
            except Exception:
                item["task_no"] = 0
            try:
                item["effort"] = float(str(item["effort"]).replace(",", "."))
            except Exception:
                item["effort"] = 0
            rows.append(item)
    return _validate_imported(normalise_data(rows))


def _validate_imported(rows: list) -> tuple[list, list[str]]:
    """Reject rows the rest of the tool cannot handle (they would crash
    totals or be dropped silently at the next week change)."""
    valid, rejected = [], []
    for i, a in enumerate(rows, 1):
        problems = []
        try:
            safe_date(str(a.get("date", "")))
        except ValueError:
            problems.append(f"date {a.get('date', '')!r}")
        try:
            if float(a.get("effort", 0)) <= 0:
                problems.append("effort must be > 0")
        except (TypeError, ValueError):
            problems.append(f"effort {a.get('effort')!r}")
        if not str(a.get("user", "")).strip():
            problems.append("no user")
        if problems:
            rejected.append(f"row {i}: " + ", ".join(problems))
        else:
            valid.append(a)
    return valid, rejected

# ---------------------------------------------------------------------------
# Date picker popup
# ---------------------------------------------------------------------------

class DatePicker(tk.Toplevel):
    """Minimal calendar popup — returns selected date via callback."""

    def __init__(self, parent, initial: date, callback):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self._cb = callback
        self._year  = initial.year
        self._month = initial.month
        self._sel   = initial
        self.configure(bg=BG_PANEL)
        self._build()
        # Position near parent
        self.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty() + parent.winfo_height() + 2
        self.geometry(f"+{px}+{py}")
        self.bind("<FocusOut>", lambda e: self.destroy())
        self.focus_set()

    def _build(self):
        # Nav bar
        nav = tk.Frame(self, bg=BG_HEADER)
        nav.pack(fill="x")
        tk.Button(nav, text="◀", font=("Segoe UI", 9), bg=BG_HEADER, fg=ACCENT,
                  relief="flat", bd=0, cursor="hand2",
                  command=self._prev).pack(side="left", padx=6, pady=4)
        self._title = tk.Label(nav, text="", font=("Segoe UI", 9, "bold"),
                               bg=BG_HEADER, fg="#ffffff")
        self._title.pack(side="left", expand=True)
        tk.Button(nav, text="▶", font=("Segoe UI", 9), bg=BG_HEADER, fg=ACCENT,
                  relief="flat", bd=0, cursor="hand2",
                  command=self._next).pack(side="right", padx=6, pady=4)
        tk.Frame(self, bg=ACCENT, height=2).pack(fill="x")

        # Day-of-week headers (col 0 = CW label, cols 1-7 = days)
        hdr = tk.Frame(self, bg=BG_PANEL)
        hdr.pack(fill="x", padx=4, pady=(4, 0))
        tk.Label(hdr, text="CW", width=3, font=("Segoe UI", 7, "bold"),
                 bg=BG_PANEL, fg=FG_DIM).grid(row=0, column=0)
        for i, d in enumerate(["Mo","Tu","We","Th","Fr","Sa","Su"]):
            fg = ACCENT if i >= 5 else FG_DIM
            tk.Label(hdr, text=d, width=3, font=("Segoe UI", 8, "bold"),
                     bg=BG_PANEL, fg=fg).grid(row=0, column=i+1)

        self._grid = tk.Frame(self, bg=BG_PANEL)
        self._grid.pack(padx=4, pady=(2, 4))
        self._render()

    def _render(self):
        for w in self._grid.winfo_children():
            w.destroy()
        import calendar
        self._title.config(text=date(self._year, self._month, 1).strftime("%B %Y"))
        cal = calendar.monthcalendar(self._year, self._month)
        for r, week in enumerate(cal):
            # CW label — use Monday (first non-zero day) to determine CW
            first_day = next((day for day in week if day != 0), None)
            if first_day:
                week_date = date(self._year, self._month, first_day)
                iso_w     = week_date.isocalendar()
                cw_label  = tk.Label(self._grid, text=f"{iso_w.week}",
                                     width=3, font=("Segoe UI", 7),
                                     bg=BG_PANEL, fg=FG_DIM)
                cw_label.grid(row=r, column=0, padx=(0, 2), pady=1)
            else:
                tk.Label(self._grid, text="", width=3, bg=BG_PANEL).grid(row=r, column=0)

            for c, day in enumerate(week):
                if day == 0:
                    tk.Label(self._grid, text="", width=3, bg=BG_PANEL).grid(row=r, column=c+1)
                    continue
                d = date(self._year, self._month, day)
                is_sel     = (d == self._sel)
                is_today   = (d == date.today())
                is_weekend = (c >= 5)
                is_holiday = d in _holidays
                holiday_name = _holidays.get(d, "")

                if is_sel:
                    bg, fg = ACCENT, "#ffffff"
                elif is_holiday:
                    bg, fg = ROW_HOLIDAY, "#8b6914"
                elif is_today:
                    bg, fg = ROW_ODD, ACCENT
                elif is_weekend:
                    bg, fg = BG_PANEL, FG_DIM
                else:
                    bg, fg = BG_PANEL, FG_MAIN

                btn = tk.Label(self._grid, text=str(day), width=3,
                               font=("Segoe UI", 9, "bold" if is_sel or is_today or is_holiday else "normal"),
                               bg=bg, fg=fg, cursor="hand2", relief="flat")
                btn.grid(row=r, column=c+1, padx=1, pady=1)
                btn.bind("<Button-1>", lambda e, dd=d: self._pick(dd))
                btn.bind("<Enter>",    lambda e, b=btn, s=is_sel: b.config(
                    bg=ACCENT_DARK if not s else ACCENT))
                btn.bind("<Leave>",    lambda e, b=btn, ob=bg: b.config(bg=ob))
                if is_holiday and holiday_name:
                    add_tooltip(btn, f"🏖  {holiday_name}")

    def _prev(self):
        if self._month == 1: self._year -= 1; self._month = 12
        else: self._month -= 1
        self._render()

    def _next(self):
        if self._month == 12: self._year += 1; self._month = 1
        else: self._month += 1
        self._render()

    def _pick(self, d: date):
        self._sel = d
        self._cb(d)
        self.destroy()


# ---------------------------------------------------------------------------
# Tooltip helper
# ---------------------------------------------------------------------------

class _Tooltip:
    def __init__(self, widget, text: str):
        self._widget = widget; self._text = text; self._tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        if self._tip:
            return
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._tip, text=self._text,
            background="#fffde7", relief="solid", borderwidth=1,
            font=("Segoe UI", 8), wraplength=260, justify="left",
        ).pack()

    def _hide(self, _=None):
        if self._tip:
            self._tip.destroy(); self._tip = None

def add_tooltip(widget, text: str) -> None:
    _Tooltip(widget, text)

# ---------------------------------------------------------------------------
# Splash screen
# ---------------------------------------------------------------------------

class SplashScreen:
    def __init__(self, root: tk.Tk):
        self._win = tk.Toplevel(root)
        self._win.overrideredirect(True)
        self._win.attributes("-topmost", True)
        w, h = 480, 260
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        self._win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self._win.configure(bg=SPLASH_BG)

        # Top accent line
        tk.Frame(self._win, bg=ACCENT, height=3).pack(fill="x")

        # Logo area
        logo = _load_image(LOGO_SPLASH_FILE, master=self._win)
        if logo is not None:
            lbl = tk.Label(self._win, image=logo, bg=SPLASH_BG, bd=0)
            lbl.image = logo
            lbl.pack(pady=(22, 6))
        logo_frame = tk.Frame(self._win, bg=SPLASH_BG)
        logo_frame.pack(pady=(0 if logo is not None else 24, 0))
        if logo is None:
            tk.Label(logo_frame, text="◈", font=("Segoe UI", 22),
                     bg=SPLASH_BG, fg=ACCENT).pack(side="left", padx=(0, 10))
        tk.Label(logo_frame, text=APP_TITLE,
                 font=("Segoe UI Semibold", 16, "bold"),
                 bg=SPLASH_BG, fg=SPLASH_FG).pack(side="left")

        tk.Label(self._win, text="EPS-Sterna · EPS-Aeolus · EUMETSAT",
                 font=("Segoe UI", 8), bg=SPLASH_BG, fg=SPLASH_DIM).pack(pady=(4, 0))

        # Divider
        tk.Frame(self._win, bg=BORDER if not DARK else "#30363d", height=1).pack(
            fill="x", padx=40, pady=(16, 12))

        self._sv = tk.StringVar(value="Starting up…")
        tk.Label(self._win, textvariable=self._sv,
                 font=("Segoe UI", 9), bg=SPLASH_BG, fg=SPLASH_DIM).pack()

        pb = ttk.Progressbar(self._win, mode="indeterminate", length=400, style="TProgressbar")
        pb.pack(pady=(10, 0), padx=40)
        pb.start(10)
        self._pb = pb

        # Bottom accent line
        tk.Frame(self._win, bg=ACCENT, height=3).pack(fill="x", side="bottom")
        self._win.update()

    def set_status(self, msg: str) -> None:
        if getattr(self, "_closed", False):
            return
        self._sv.set(msg); self._win.update_idletasks()

    def close(self) -> None:
        # Safe to call twice: startup closes the splash early when the lock
        # dialog is shown, then again at the end (crashed after Force Unlock)
        if getattr(self, "_closed", False):
            return
        self._closed = True
        self._pb.stop(); self._win.destroy()

# ---------------------------------------------------------------------------
# Lock error dialog
# ---------------------------------------------------------------------------

def show_lock_error(root: tk.Tk, owner: str) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Application already in use")
    dialog.resizable(False, False)
    dialog.attributes("-topmost", True)
    w, h = 500, 250
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    dialog.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
    tk.Label(dialog, text="⚠", font=("Segoe UI", 28), fg="#e65100").pack(pady=(18, 0))
    tk.Label(dialog, text="Activity Report Tool is currently in use by:",
             font=("Segoe UI", 9)).pack()
    tk.Label(dialog, text=owner, font=("Segoe UI", 9, "bold"), fg="#c62828").pack(pady=(2, 8))
    tk.Label(
        dialog,
        text="Please wait until they close the application.\n"
             "If you are sure nobody else is using it, click Force Unlock.",
        font=("Segoe UI", 9), justify="center", wraplength=440,
    ).pack()
    result = {"force": False}
    bf = tk.Frame(dialog); bf.pack(pady=(16, 0))
    ttk.Button(bf, text="Exit",            command=lambda: (result.update(force=False), dialog.destroy()), width=14).grid(row=0, column=0, padx=8)
    ttk.Button(bf, text="🔓  Force Unlock", command=lambda: (result.update(force=True),  dialog.destroy()), style="Danger.TButton", width=18).grid(row=0, column=1, padx=8)
    dialog.protocol("WM_DELETE_WINDOW", lambda: (result.update(force=False), dialog.destroy()))
    dialog.grab_set()
    root.wait_window(dialog)
    if result["force"]:
        release_lock(force=True)
        if not acquire_lock():
            messagebox.showerror("Lock error", "Could not acquire lock even after force unlock. Exiting.")
            root.destroy(); sys.exit(1)
    else:
        root.destroy(); sys.exit(0)

# ---------------------------------------------------------------------------
# Shared Toplevel helper
# ---------------------------------------------------------------------------

class _FittedToplevel(tk.Toplevel):
    """Toplevel that centres itself within the usable screen area."""

    def _fit_window(self, want_w: int, want_h: int) -> None:
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        # Leave room for the taskbar (subtract 80 px as a safe margin)
        avail_h = sh - 80
        w = min(want_w, sw - 40)
        h = min(want_h, avail_h)
        x = (sw - w) // 2
        y = max(0, (avail_h - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")


# ---------------------------------------------------------------------------
# Next-week plan editor window
# ---------------------------------------------------------------------------

class NextWeekEditor(_FittedToplevel):
    DEFAULT_GROUPS = ["EPS-Sterna", "EPS-Aeolus", "Common"]

    def __init__(self, parent, tasks: list[tuple], target: tuple[int, int]):
        super().__init__(parent)
        self.title("Next Week — Planned Activities")
        self.resizable(True, True)
        self._fit_window(780, 560)
        self.tasks   = tasks
        self._target = target   # (year, CW) being planned
        self._plan   = load_next_week_plan(target)
        # Same grouping as the weekly summary: by WP when tasks.txt has WPs
        wps = sorted({wp for _, _, wp in tasks if wp},
                     key=lambda w: (0 if w.upper().startswith("WP") else 1, w))
        groups = wps or self.DEFAULT_GROUPS
        self.PROJECTS = groups + [k for k in self._plan if k not in groups]
        self._texts: dict[str, tk.Text] = {}
        self._build()
        self.grab_set()

    def _build(self):
        # Header
        tk.Label(
            self, text="Plan activities for next week",
            font=("Segoe UI", 11, "bold"), fg=FG_MAIN, bg=BG_ROOT,
        ).pack(pady=(14, 2))

        y, cw    = self._target
        next_mon = iso_week_monday(y, cw)
        next_fri = next_mon + timedelta(days=4)
        tk.Label(
            self,
            text=f"CW{cw:02d} / {y}  —  {next_mon.strftime('%d %b')} → {next_fri.strftime('%d %b %Y')}",
            font=("Segoe UI", 9), fg=FG_DIM, bg=BG_ROOT,
        ).pack()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=14, pady=10)

        for project in self.PROJECTS:
            frame = ttk.Frame(nb)
            nb.add(frame, text=project)

            tk.Label(frame, text="One bullet per line. Lines starting with # are ignored.",
                     font=("Segoe UI", 8), fg=FG_DIM).pack(anchor="w", padx=6, pady=(6, 2))

            txt = scrolledtext.ScrolledText(frame, font=("Segoe UI", 10), wrap="word", height=14)
            txt.pack(fill="both", expand=True, padx=6, pady=(0, 6))

            existing = self._plan.get(project, [])
            txt.insert("1.0", "\n".join(existing))
            self._texts[project] = txt

        # Quick-add from task list
        qa_frame = ttk.LabelFrame(self, text="Quick-add from task list")
        qa_frame.pack(fill="x", padx=14, pady=(0, 6))

        self._qa_project = tk.StringVar(value=self.PROJECTS[0])
        self._qa_task    = tk.StringVar()
        self._qa_desc    = tk.StringVar()

        ttk.Label(qa_frame, text="Project").grid(row=0, column=0, padx=6, pady=4)
        ttk.Combobox(qa_frame, textvariable=self._qa_project, values=self.PROJECTS,
                     width=14, state="readonly").grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(qa_frame, text="Task").grid(row=0, column=2, padx=6)
        task_vals = [f"{(wp+' - ') if wp else ''}{n} - {t}" for n, t, wp in self.tasks]
        ttk.Combobox(qa_frame, textvariable=self._qa_task, values=task_vals,
                     width=40, state="readonly").grid(row=0, column=3, padx=4)

        ttk.Label(qa_frame, text="Description").grid(row=1, column=0, padx=6)
        ttk.Entry(qa_frame, textvariable=self._qa_desc, width=60).grid(
            row=1, column=1, columnspan=3, sticky="we", padx=4, pady=4)

        ttk.Button(qa_frame, text="➕  Add bullet", command=self._quick_add).grid(
            row=1, column=4, padx=8)

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=(4, 12))
        ttk.Button(btn_frame, text="💾  Save plan",   command=self._save,  width=16).grid(row=0, column=0, padx=8)
        ttk.Button(btn_frame, text="🗑  Clear all",   command=self._clear, width=14).grid(row=0, column=1, padx=8)
        ttk.Button(btn_frame, text="✖  Close",        command=self.destroy, width=12).grid(row=0, column=2, padx=8)

    def _quick_add(self):
        desc    = self._qa_desc.get().strip()
        project = self._qa_project.get()
        if not desc:
            return
        txt = self._texts.get(project)
        if txt:
            current = txt.get("1.0", "end").strip()
            bullet  = _clean_bullet(desc)
            txt.delete("1.0", "end")
            txt.insert("1.0", (current + "\n" + bullet).strip())
        self._qa_desc.set("")

    def _collect(self) -> dict[str, list[str]]:
        plan = {}
        for project, txt in self._texts.items():
            lines = [
                _clean_bullet(l.strip())
                for l in txt.get("1.0", "end").splitlines()
                if l.strip() and not l.strip().startswith("#")
            ]
            if lines:
                plan[project] = lines
        return plan

    def _save(self):
        if read_only:
            messagebox.showwarning("Read-only", READ_ONLY_MSG, parent=self)
            return
        plan = self._collect()
        save_next_week_plan(plan, self._target)
        messagebox.showinfo("Saved", f"Plan for CW{self._target[1]:02d}/{self._target[0]} saved.",
                            parent=self)

    def _clear(self):
        if messagebox.askyesno("Clear", "Clear all planned activities?", parent=self):
            for txt in self._texts.values():
                txt.delete("1.0", "end")

# ---------------------------------------------------------------------------
# Monthly statistics window
# ---------------------------------------------------------------------------

class StatsWindow(_FittedToplevel):
    def __init__(self, parent, data: list, year: int, month: int):
        super().__init__(parent)
        self.title(f"Statistics — {full_month_label(date(year, month, 1))}")
        self.resizable(True, True)
        self._fit_window(740, 540)
        self._build(data, year, month)

    def _build(self, data, year, month):
        month_start = date(year, month, 1)
        month_end   = last_day_of_month(year, month)

        def _valid_activity(a):
            if a.get("entry_type") != "Activity":
                return False
            try:
                if not (month_start <= safe_date(a["date"]) <= month_end):
                    return False
            except Exception:
                return False
            return int(a.get("task_no", 0) or 0) > 0

        filtered = [a for a in normalise_data(data) if _valid_activity(a)]

        def _in_month(a):
            try:
                return month_start <= safe_date(a.get("date", "")) <= month_end
            except ValueError:
                return False

        leave_entries = [
            a for a in normalise_data(data)
            if a.get("entry_type") in ("Holiday", "Sick Leave") and _in_month(a)
        ]

        # ── Working days summary panel ────────────────────────────────────────
        # ── Working days calculation ─────────────────────────────────────────
        total_days     = (month_end - month_start).days + 1
        weekend_days   = sum(1 for i in range(total_days)
                             if (month_start + timedelta(days=i)).weekday() >= 5)
        month_holidays = {d for d in _holidays
                          if month_start <= d <= month_end
                          and d.weekday() < 5}
        working_days   = total_days - weekend_days - len(month_holidays)

        # Number of active users in this month — including users who were
        # only on leave, otherwise their leave lowers "expected" below zero
        active_users = sorted({a.get("user", "") for a in filtered + leave_entries if a.get("user")})
        n_users      = max(len(active_users), 1)

        # Total capacity = working_days × n_users
        capacity     = working_days * n_users

        # Leave days (per user, summed across all users)
        leave_effort = sum(float(a.get("effort", 0) or 0) for a in leave_entries)

        # Expected effort = capacity - leave
        expected     = capacity - leave_effort

        # Actual effort logged (Activity entries only)
        total_effort = sum(float(a.get("effort", 0) or 0) for a in filtered)

        # Delta: positive = more than expected, negative = less
        delta        = total_effort - expected
        delta_str    = f"+{delta:g}" if delta > 0 else f"{delta:g}"
        delta_color  = "#2e7d32" if abs(delta) < 0.5 else "#c62828"
        delta_tip    = ("✓ Effort matches capacity" if abs(delta) < 0.5
                        else f"{'Excess' if delta > 0 else 'Missing'} {abs(delta):g} PD — check entries")

        # ── Summary panel ─────────────────────────────────────────────────────
        panel = tk.Frame(self, bg=BG_PANEL, relief="flat", bd=0)
        panel.pack(fill="x", padx=12, pady=(10, 0))

        def _cell(parent, label, value, value_color=FG_MAIN, col=0, tip=""):
            f = tk.Frame(parent, bg=BG_PANEL)
            f.grid(row=0, column=col, padx=(0, 20), sticky="w")
            tk.Label(f, text=label, font=("Segoe UI", 8),
                     bg=BG_PANEL, fg=FG_DIM).pack(anchor="w")
            lbl = tk.Label(f, text=value, font=("Segoe UI Semibold", 13, "bold"),
                           bg=BG_PANEL, fg=value_color)
            lbl.pack(anchor="w")
            if tip:
                add_tooltip(lbl, tip)

        _cell(panel, f"Working days",
              f"{working_days}",
              tip=f"{total_days} days − {weekend_days} weekends − {len(month_holidays)} holidays",
              col=0)
        _cell(panel, f"Users ({n_users})",
              f"× {n_users}",
              tip=", ".join(active_users),
              col=1)
        _cell(panel, "Capacity (PD)",
              f"{capacity:g}",
              tip=f"{working_days} × {n_users} users",
              col=2)
        _cell(panel, "Leave (PD)",
              f"{leave_effort:g}",
              tip="Holiday + Sick Leave days logged",
              col=3)
        _cell(panel, "Effort logged (PD)",
              f"{total_effort:g}",
              col=4)
        _cell(panel, "Delta",
              delta_str, delta_color,
              tip=delta_tip,
              col=5)

        # Holidays list
        if month_holidays:
            hol_names = ", ".join(
                _holidays.get(d, str(d)) for d in sorted(month_holidays)
            )
            tk.Label(panel, text=f"Holidays: {hol_names}",
                     font=("Segoe UI", 8), bg=BG_PANEL, fg=FG_DIM,
                     wraplength=600, justify="left"
                     ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(4, 0))

        # Divider
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(8, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=6)

        # ── By User ──────────────────────────────────────────────────────────
        f1 = ttk.Frame(nb); nb.add(f1, text="By User")
        by_user: dict[str, float] = defaultdict(float)
        for a in filtered:
            by_user[a["user"]] += float(a.get("effort", 0) or 0)
        self._make_table(f1, ["User", "Person-Days"],
                         [(u, f"{v:g}") for u, v in sorted(by_user.items())],
                         totals=("TOTAL", f"{sum(by_user.values()):g}"))

        # ── By Task ──────────────────────────────────────────────────────────
        f2 = ttk.Frame(nb); nb.add(f2, text="By Task")
        by_task: dict[tuple, float] = defaultdict(float)
        for a in filtered:
            wp    = entry_wp(a)
            key   = (wp, int(a.get("task_no", 0) or 0), a.get("task_title", ""))
            by_task[key] += float(a.get("effort", 0) or 0)
        self._make_table(
            f2, ["WP", "Task No", "Task Title", "Person-Days"],
            [(wp, str(no), title, f"{v:g}") for (wp, no, title), v in sorted(by_task.items())],
            totals=("", "", "TOTAL", f"{sum(by_task.values()):g}"),
        )

        # ── By Project ───────────────────────────────────────────────────────
        f3 = ttk.Frame(nb); nb.add(f3, text="By Project")
        by_proj: dict[str, float] = defaultdict(float)
        for a in filtered:
            by_proj[project_for_task(
                a.get("task_no", 0),
                entry_wp(a)
            )] += float(a.get("effort", 0) or 0)
        self._make_table(f3, ["Project", "Person-Days"],
                         [(p, f"{v:g}") for p, v in sorted(by_proj.items())],
                         totals=("TOTAL", f"{sum(by_proj.values()):g}"))

        # ── By WP ────────────────────────────────────────────────────────────
        f5 = ttk.Frame(nb); nb.add(f5, text="By WP")
        by_wp: dict[str, float] = defaultdict(float)
        for a in filtered:
            wp = entry_wp(a) or project_for_task(a.get("task_no", 0), "")
            by_wp[wp] += float(a.get("effort", 0) or 0)
        self._make_table(f5, ["WP", "Person-Days"],
                         [(w, f"{v:g}") for w, v in sorted(by_wp.items())],
                         totals=("TOTAL", f"{sum(by_wp.values()):g}"))

        # ── Absences (Holiday / Sick Leave) ──────────────────────────────────
        if leave_entries:
            f4 = ttk.Frame(nb); nb.add(f4, text="Absences")
            by_leave: dict[tuple, float] = defaultdict(float)
            for a in leave_entries:
                key = (a.get("user", ""), a.get("entry_type", ""))
                by_leave[key] += float(a.get("effort", 0) or 0)
            self._make_table(
                f4, ["User", "Type", "Days"],
                [(u, t, f"{v:g}") for (u, t), v in sorted(by_leave.items())],
                totals=("", "TOTAL", f"{sum(by_leave.values()):g}"),
            )

    @staticmethod
    def _make_table(parent, headers, rows, totals=None):
        tree = ttk.Treeview(parent, columns=headers, show="headings")
        for h in headers:
            tree.heading(h, text=h)
            numeric_cols = ("Person-Days", "Days", "PD")
            wide_cols    = ("User", "Task Title", "Project", "Type")
            tree.column(h,
                        width=180 if h in wide_cols else 80,
                        anchor="center" if h in numeric_cols else "w")
        for i, row in enumerate(rows):
            tree.insert("", "end", values=row, tags=("odd" if i % 2 else "even",))
        if totals:
            tree.insert("", "end", values=totals, tags=("total",))
        tree.tag_configure("odd",   background=ROW_ODD)
        tree.tag_configure("even",  background=ROW_EVEN)
        tree.tag_configure("total", background="#dbeafe", font=("Segoe UI", 9, "bold"))
        sb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscroll=sb.set)
        tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        sb.pack(side="right", fill="y", pady=10)

# ---------------------------------------------------------------------------
# Generating dialog — animated spinner shown during Word generation
# ---------------------------------------------------------------------------

class GeneratingDialog(tk.Toplevel):
    """Modal dialog with animated arc spinner shown while generating the Word report."""

    _STEPS   = ["Collecting data…", "Opening template…",
                "Filling user tables…", "Computing totals…",
                "Updating dates…", "Saving document…"]
    _ARC_GAP = 60        # degrees left empty in the arc
    _SPEED   = 18        # ms per animation frame

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=BG_PANEL)

        w, h = 340, 240
        sw, sh = parent.winfo_screenwidth(), parent.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{max(0,(sh-80-h)//2)}")

        # Top accent
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # Title
        tk.Label(self, text="Generating Word report…",
                 font=("Segoe UI Semibold", 11, "bold"),
                 bg=BG_PANEL, fg=FG_MAIN).pack(pady=(20, 0))

        # Canvas for spinner
        self._canvas = tk.Canvas(self, width=90, height=90,
                                 bg=BG_PANEL, highlightthickness=0)
        self._canvas.pack(pady=(10, 0))

        # Status label
        self._status_var = tk.StringVar(value=self._STEPS[0])
        tk.Label(self, textvariable=self._status_var,
                 font=("Segoe UI", 9), bg=BG_PANEL, fg=FG_DIM).pack(pady=(8, 0))

        # Dots progress indicator
        self._dots_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._dots_var,
                 font=("Segoe UI", 8), bg=BG_PANEL, fg=FG_DIM).pack()

        # Bottom accent
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x", side="bottom")

        self._angle      = 0
        self._step_idx   = 0
        self._dot_count  = 0
        self._arc_id     = None
        self._running    = True

        self._draw_spinner()
        self._animate()
        self.update()

    def _draw_spinner(self):
        cx, cy, r = 45, 45, 32
        # Background ring
        self._canvas.create_oval(cx-r, cy-r, cx+r, cy+r,
                                  outline=BORDER, width=5, fill="")
        # Colored arc — redrawn each frame
        self._arc = self._canvas.create_arc(
            cx-r, cy-r, cx+r, cy+r,
            start=self._angle, extent=360-self._ARC_GAP,
            style="arc", outline=ACCENT, width=5
        )
        # Inner dot
        dr = 6
        self._dot = self._canvas.create_oval(cx-dr, cy-dr, cx+dr, cy+dr,
                                              fill=ACCENT, outline="")

    def _animate(self):
        if not self._running:
            return
        self._angle = (self._angle + 8) % 360
        self._canvas.itemconfig(self._arc, start=self._angle)

        # Pulse inner dot
        pulse = 0.7 + 0.3 * abs((self._angle % 60) / 30 - 1)
        r = int(6 * pulse)
        cx, cy = 45, 45
        self._canvas.coords(self._dot, cx-r, cy-r, cx+r, cy+r)

        # Rotate status text every ~40 frames
        self._dot_count += 1
        if self._dot_count % 40 == 0:
            self._step_idx = (self._step_idx + 1) % len(self._STEPS)
            self._status_var.set(self._STEPS[self._step_idx])

        # Animate dots
        ndots = (self._dot_count // 10) % 4
        self._dots_var.set("●" * ndots + "○" * (3 - ndots))

        self._arc_id = self.after(self._SPEED, self._animate)

    def close(self):
        self._running = False
        if self._arc_id:
            try:
                self.after_cancel(self._arc_id)
            except Exception:
                pass
        self.destroy()


# ---------------------------------------------------------------------------
# Main App
# ---------------------------------------------------------------------------

class App:

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.minsize(1100, 600)
        root.withdraw()

        icon_path = _asset_path(ICON_FILE)
        if icon_path and sys.platform == "win32":
            try:
                root.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass
        icon_image = _load_image(ICON_PNG_FILE, master=root)
        if icon_image is not None:
            self._icon_image = icon_image
            root.iconphoto(True, icon_image)

        # ── Fit window inside usable screen area (taskbar-aware) ─────────────
        # Geometry is set AFTER widgets are built — see _fit_to_screen() call below

        # ── Lock ────────────────────────────────────────────────────────────
        # Apply Danger style early so show_lock_error can use it
        style = ttk.Style()
        try: style.theme_use("clam")
        except Exception: pass
        style.configure("Danger.TButton", foreground="white", background="#c62828")
        style.map("Danger.TButton", background=[("active", "#b71c1c")])

        # Lock is acquired AFTER io_dir is set (see below)
        _running_as_exe = getattr(sys, "frozen", False)

        # ── Splash ──────────────────────────────────────────────────────────
        splash = SplashScreen(root)

        # ── Full style — Aerospace Professional theme ───────────────────────
        splash.set_status("Loading interface…")

        FONT_UI    = ("Segoe UI", 9)
        FONT_BOLD  = ("Segoe UI", 9, "bold")
        FONT_SMALL = ("Segoe UI", 8)
        FONT_MONO  = ("Consolas", 9)

        root.configure(bg=BG_ROOT)

        style.configure("TFrame",        background=BG_FRAME)
        style.configure("TLabel",        background=BG_FRAME, font=FONT_UI, foreground=FG_MAIN)
        style.configure("TEntry",        font=FONT_UI, fieldbackground=BG_ENTRY,
                         foreground=FG_MAIN, insertcolor=FG_MAIN)
        style.configure("TCombobox",     font=FONT_UI, fieldbackground=BG_ENTRY,
                         foreground=FG_MAIN, selectbackground=ACCENT)
        style.map("TCombobox",           fieldbackground=[("readonly", BG_ENTRY)])

        # Primary button — cyan accent
        style.configure("Accent.TButton",
                         font=FONT_BOLD, padding=(10, 5),
                         background=ACCENT, foreground="#ffffff", relief="flat")
        style.map("Accent.TButton",
                  background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK)],
                  relief=[("pressed", "flat")])

        # Standard button
        style.configure("TButton",
                         font=FONT_UI, padding=(8, 4),
                         background=BG_PANEL, foreground=FG_MAIN, relief="flat",
                         borderwidth=1)
        style.map("TButton",
                  background=[("active", BG_ENTRY), ("pressed", BORDER)],
                  relief=[("pressed", "flat")])

        # Danger button
        style.configure("Danger.TButton",
                         font=FONT_BOLD, padding=(10, 5),
                         foreground="white", background=DANGER, relief="flat")
        style.map("Danger.TButton",
                  background=[("active", DANGER_DARK), ("pressed", DANGER_DARK)])

        # Treeview
        style.configure("Treeview",
                         rowheight=24, font=FONT_UI,
                         fieldbackground=ROW_EVEN, background=ROW_EVEN,
                         foreground=FG_MAIN, borderwidth=0)
        style.configure("Treeview.Heading",
                         font=FONT_BOLD, background=BG_HEADER,
                         foreground=FG_ACCENT if DARK else "#ffffff",
                         relief="flat", padding=(6, 6))
        style.map("Treeview",
                  background=[("selected", ROW_SEL)],
                  foreground=[("selected", FG_MAIN)])
        style.map("Treeview.Heading",
                  background=[("active", ACCENT_DARK)])

        # Notebook
        style.configure("TNotebook",       background=BG_FRAME, borderwidth=0)
        style.configure("TNotebook.Tab",   font=FONT_UI, padding=(12, 5),
                         background=BG_PANEL, foreground=FG_DIM)
        style.map("TNotebook.Tab",
                  background=[("selected", BG_FRAME)],
                  foreground=[("selected", FG_ACCENT)])

        # LabelFrame
        style.configure("TLabelframe",       background=BG_FRAME, bordercolor=BORDER)
        style.configure("TLabelframe.Label", background=BG_FRAME, font=FONT_BOLD,
                         foreground=FG_ACCENT)

        # Progressbar (splash)
        style.configure("TProgressbar", troughcolor=BG_PANEL, background=ACCENT,
                         thickness=4)

        # Scrollbar
        style.configure("TScrollbar", background=BG_PANEL, troughcolor=BG_FRAME,
                         borderwidth=0, arrowsize=12)
        style.map("TScrollbar", background=[("active", ACCENT)])

        # Error state for date entry (configured once — BUG H fix)
        style.configure("Error.TEntry", fieldbackground="#ffcccc")

        # ── Init data attrs — data loaded later after io_dir is set ────────────
        self.data:             list       = []
        self.edit_index:       int | None = None
        self._undo_stack:      list       = []
        self._save_count:      int        = 0
        self._loaded_week_keys: set       = set()
        self._table_dirty:     bool       = False
        self._working_week:         tuple | None = None
        self._pending_duplicate_idx: int | None  = None
        self._new_entry_keys: set = set()   # activity_key of entries added after load
        was_reset  = False
        saved_file = None

        # ── App state ───────────────────────────────────────────────────────
        splash.set_status("Restoring app state…")
        saved_tpl, saved_users, saved_tasks, saved_last_user, saved_io_dir = load_app_state()
        self.template_path:   Path | None = None
        self.users:           list[str]   = []
        self.tasks:           list[tuple] = []
        self.last_user:       str         = saved_last_user
        self.last_io_dir: str             = saved_io_dir

        # ── Auto-detect defaults on first launch ─────────────────────────────
        if not self.last_io_dir:
            splash.set_status("Looking for default working folder...")
            _default = Path.home() / "OneDrive - eumetsat.int" / "Exostaff Report Generator"
            if _default.exists():
                self.last_io_dir = str(_default)
                log.warning("Auto-detected working folder: %s", self.last_io_dir)

        if not saved_tpl:
            _tpl_dirs = []
            if self.last_io_dir:
                _tpl_dirs += [Path(self.last_io_dir) / "config", Path(self.last_io_dir)]
            _tpl_dirs.append(_exe_dir())
            for _td in _tpl_dirs:
                # Prefer files named like the template; skip Word lock files (~$...)
                _found = sorted(
                    (f for f in _td.glob("*.docx") if not f.name.startswith("~$")),
                    key=lambda f: ("template" not in f.name.lower(), f.name),
                )
                if _found:
                    saved_tpl = _found[0]
                    log.warning("Auto-detected template: %s", saved_tpl)
                    break

        self._pending_folder_warning = None
        if self.last_io_dir:
            # Validate saved io_dir exists before using it
            _saved_path = Path(self.last_io_dir)
            if not _saved_path.exists():
                log.warning("Saved io_dir does not exist: %s", self.last_io_dir)
                self._pending_folder_warning = self.last_io_dir
                self.last_io_dir = ""
            else:
                set_io_dir(self.last_io_dir)
                setup_log_file()
                splash.set_status("Loading configuration…")
                self.users, self.tasks = reload_config()
                # Acquire lock now that IO folder is known
                if _running_as_exe and not acquire_lock():
                    splash.close()
                    show_lock_error(root, read_lock_info())
                # Read data from share (AFTER io_dir is set)
                splash.set_status("Reading activity data…")
                self.data, was_reset, saved_file = read_json()
                # Entries loaded in the previous session (incl. deleted ones
                # whose deletion still has to reach the archive)
                self._loaded_week_keys = (load_table_state()
                                          | {_stable_key_fn(a) for a in self.data})
                self._table_dirty_after_load()
                # Set working week
                if self.data:
                    try:
                        today_yw = iso_year_week_for_date(date.today())
                        current_entries = [a for a in self.data
                                           if iso_year_week_for_date(safe_date(a["date"])) == today_yw]
                        if current_entries:
                            self._working_week = today_yw
                        else:
                            ref = max(safe_date(a["date"]) for a in self.data if a.get("date"))
                            self._working_week = iso_year_week_for_date(ref)
                    except Exception:
                        pass

        # Version check — runs regardless of io_dir block
        if self.last_io_dir:
            splash.set_status("Checking version…")
            _latest, _update_available = check_version()
            self._pending_version_notice = (_latest, APP_VERSION) if _update_available else None
        else:
            self._pending_version_notice = None

        # Template path — restored from state (used only for monthly Word generation)
        if saved_tpl and saved_tpl.exists():
            self.template_path = saved_tpl
        elif DEFAULT_TEMPLATE and DEFAULT_TEMPLATE.exists():
            self.template_path = DEFAULT_TEMPLATE

        root.protocol("WM_DELETE_WINDOW", self.on_close)

        splash.set_status("Building interface…")
        self._build_menu()
        self.build_ui()
        if self._table_dirty:
            self._set_table_dirty()   # highlight Save week now that it exists
        self.refresh()
        self._bind_shortcuts()

        splash.close()
        self._fit_to_screen()
        root.deiconify()
        root.lift()
        root.focus_force()

        if _running_as_exe:
            self.root.after(LOCK_HEARTBEAT_MS, self._lock_heartbeat)
        self.root.after(UPDATE_CHECK_MS, self._periodic_update_check)

        if data_load_warning:
            messagebox.showwarning("Data file problem", data_load_warning)

        # Show warning if saved working folder was not found
        if getattr(self, "_pending_folder_warning", None):
            missing = self._pending_folder_warning
            messagebox.showwarning(
                "Working folder not found",
                f"The saved working folder was not found:\n\n{missing}\n\n"
                f"This can happen if a network drive is disconnected or "
                f"OneDrive has changed the path.\n\n"
                f"Please set the working folder again with the 📁 button.",
            )

        # Show version update notice if needed
        if getattr(self, "_pending_version_notice", None):
            latest, current = self._pending_version_notice
            self._show_update_notice(latest, current)

        if was_reset:
            n = len(self.data)
            iso = date.today().isocalendar()
            msg = (f"New week detected — previous week archived automatically.\n\n"
                   f"CW{iso.week:02d}/{iso.year}: {n} entries loaded from archive.")
            if saved_file:
                msg += f"\n\nArchived to: {saved_file.name}"
            messagebox.showinfo("New week", msg)

    # -----------------------------------------------------------------------
    # Menu bar
    # -----------------------------------------------------------------------

    def _build_menu(self):
        menubar = tk.Menu(self.root)

        # File
        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="Export CSV/JSON…",      command=self.export_table)
        m_file.add_command(label="Export Excel (.xlsx)…",  command=self.export_xlsx)
        m_file.add_command(label="Import CSV/JSON…",       command=self.import_table)
        m_file.add_separator()
        m_file.add_command(label="Set working folder…", command=self.choose_io_dir)
        m_file.add_separator()
        m_file.add_command(label="Reset…",                 command=self.reset_data)
        m_file.add_separator()
        m_file.add_command(label="Exit",                   command=self.on_close)
        menubar.add_cascade(label="File", menu=m_file)

        # Edit
        m_edit = tk.Menu(menubar, tearoff=0)
        m_edit.add_command(label="Next-week plan…",  command=self.open_next_week_editor)
        m_edit.add_separator()
        m_edit.add_command(label="Edit users…",       command=self.open_users_config)
        m_edit.add_command(label="Edit tasks…",       command=self.open_tasks_config)
        m_edit.add_command(label="Edit holidays…",    command=self.open_holidays_config)
        m_edit.add_command(label="Reload config",     command=self.reload_config_from_files)
        menubar.add_cascade(label="Edit", menu=m_edit)

        # View
        m_view = tk.Menu(menubar, tearoff=0)
        m_view.add_command(label="Monthly statistics…", command=self.show_stats)
        m_view.add_command(label="Refresh\tF5",         command=self.refresh)
        menubar.add_cascade(label="View", menu=m_view)

        # Help
        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="User Guide…",          command=self.show_user_guide)
        m_help.add_command(label="Writing Guidelines…",  command=self.show_writing_guidelines)
        m_help.add_separator()
        m_help.add_command(label="View log…",             command=self.view_log)
        m_help.add_command(label="About",                 command=self._about)
        menubar.add_cascade(label="Help", menu=m_help)

        self.root.config(menu=menubar)

    # -----------------------------------------------------------------------
    # Keyboard shortcuts
    # -----------------------------------------------------------------------

    def _bind_shortcuts(self):
        self.root.bind("<Control-s>", lambda e: self.save_activity())
        # Not while typing in a field: users expect Ctrl+Z to undo text there,
        # not to silently revert the last table change on the share
        self.root.bind("<Control-z>", lambda e: None if isinstance(e.widget, (tk.Entry, tk.Text))
                       else self.undo())
        self.root.bind("<F5>",        lambda e: self.refresh())
        self.tree.bind("<Delete>",    lambda e: self.delete_selected())

    # -----------------------------------------------------------------------
    # Close / undo
    # -----------------------------------------------------------------------

    def on_close(self) -> None:
        if self._form_dirty():
            if not messagebox.askyesno(
                "Unsaved changes",
                "You have unsaved changes in the form.\nExit anyway?",
            ):
                return
        self._shutdown()

    def _shutdown(self) -> None:
        """Persist state, release the shared lock and close."""
        save_app_state(
            self.template_path, self.users, self.tasks,
            self.user_var.get() if hasattr(self, "user_var") else "",
            self.last_io_dir,
        )
        release_lock()
        self.root.destroy()

    def _lock_heartbeat(self) -> None:
        """Keep our lock fresh so other PCs don't treat it as orphaned."""
        if not refresh_lock():
            global read_only
            log.warning("Lock taken over by %s", read_lock_info())
            read_only = True   # never overwrite the other user's work
            self.root.title(f"{APP_TITLE}  —  READ-ONLY")
            messagebox.showwarning(
                "Lock taken over",
                f"Another user has forced the unlock:\n{read_lock_info()}\n\n"
                "The tool is now READ-ONLY so that it cannot overwrite their work: "
                "you can still view data and generate reports.\n\n"
                "Changes made since the last save are kept in this window only. "
                "Close and reopen the tool when they are done.",
            )
            return   # stop heartbeat — lock is no longer ours
        self.root.after(LOCK_HEARTBEAT_MS, self._lock_heartbeat)

    def _writable(self) -> bool:
        """False (with a message) when the tool is read-only."""
        if read_only:
            messagebox.showwarning("Read-only", READ_ONLY_MSG)
            return False
        return True

    def _persist(self) -> None:
        """Write the table and the list of loaded entries to the share."""
        if read_only:
            return   # viewing (e.g. Load week) still works, in memory only
        write_json(self.data)
        save_table_state(self._loaded_week_keys)

    def _push_undo(self, description: str) -> None:
        import copy
        self._undo_stack.append((
            description,
            copy.deepcopy(self.data),
            self._working_week,
            copy.deepcopy(self._loaded_week_keys),
        ))
        if len(self._undo_stack) > 20:
            self._undo_stack.pop(0)

    def undo(self) -> None:
        if not self._writable():
            return
        if not self._undo_stack:
            self._status("Nothing to undo.")
            return
        entry = self._undo_stack.pop()
        desc, snapshot = entry[0], entry[1]
        self.data = snapshot
        self.edit_index = None            # indices may no longer match
        self._pending_duplicate_idx = None
        self._mark_clean()
        # Restore week state if snapshot includes it
        if len(entry) >= 4:
            self._working_week     = entry[2]
            self._loaded_week_keys = entry[3]
        self._persist()
        self._clear_table_dirty()
        self.refresh()
        self._status(f"Undone: {desc}")

    # -----------------------------------------------------------------------
    # Private helpers
    # -----------------------------------------------------------------------

    def _fit_to_screen(self) -> None:
        """Size and centre the window within the usable desktop area (taskbar-aware)."""
        root = self.root
        root.update_idletasks()   # make sure all widgets are measured

        # Get usable area via Windows API, fall back to screen minus 80px
        if sys.platform == "win32":
            try:
                import ctypes
                class RECT(ctypes.Structure):
                    _fields_ = [("left",   ctypes.c_long),
                                ("top",    ctypes.c_long),
                                ("right",  ctypes.c_long),
                                ("bottom", ctypes.c_long)]
                rect = RECT()
                ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
                wa_x, wa_y = rect.left,  rect.top
                wa_w, wa_h = rect.right  - rect.left, rect.bottom - rect.top
            except Exception:
                wa_x, wa_y = 0, 0
                wa_w = root.winfo_screenwidth()
                wa_h = root.winfo_screenheight() - 80
        else:
            wa_x, wa_y = 0, 0
            wa_w = root.winfo_screenwidth()
            wa_h = root.winfo_screenheight() - 80

        # Desired size: natural size of widgets, capped to 96% of usable area
        desired_w = root.winfo_reqwidth()
        desired_h = root.winfo_reqheight()
        win_w = min(max(desired_w, 1100), int(wa_w * 0.96))
        win_h = min(max(desired_h,  600), int(wa_h * 0.96))
        win_x = wa_x + (wa_w - win_w) // 2
        win_y = wa_y + max(0, (wa_h - win_h) // 2)
        root.geometry(f"{win_w}x{win_h}+{win_x}+{win_y}")

    def _table_dirty_after_load(self) -> None:
        """A table restored from the previous session may hold changes that
        never reached the archive: treat it as unsaved, so that loading
        another week asks first instead of discarding them."""
        try:
            up_to_date = archive_is_up_to_date(self.data, self._loaded_week_keys)
        except Exception as exc:
            log.warning("archive check failed: %s", exc)
            up_to_date = False
        if up_to_date:
            self._clear_table_dirty()
        else:
            self._set_table_dirty()

    def _set_table_dirty(self) -> None:
        """Mark table as having unsaved changes — highlight Save week button."""
        self._table_dirty = True
        if hasattr(self, "_btn_save_cw"):
            self._btn_save_cw.configure(style="Accent.TButton")

    def _clear_table_dirty(self) -> None:
        """Clear dirty flag after successful save."""
        self._table_dirty = False
        if hasattr(self, "_btn_save_cw"):
            self._btn_save_cw.configure(style="TButton")

    def _collect_ref_date(self) -> date:
        dates = []
        for a in self.data:
            try: dates.append(safe_date(a["date"]))
            except Exception: pass
        return max(dates) if dates else date.today()

    def _status(self, msg: str) -> None:
        self._status_var.set(msg)
        self.root.after(6000, lambda: self._status_var.set(""))

    def _form_snapshot(self) -> tuple | None:
        if not hasattr(self, "desc"):
            return None
        return (self.date_var.get(), self.user_var.get(), self.entry_type_var.get(),
                self.effort_var.get(), self.wp_var.get(), self.task_var.get(),
                self.desc.get("1.0", "end").strip())

    def _mark_clean(self) -> None:
        """Remember the form as it is now; it is 'dirty' once it differs."""
        self._clean_form = self._form_snapshot()

    def _form_dirty(self) -> bool:
        """Editing an entry: any changed field counts. New entry: only typed
        description text counts (picking user/date/task alone is not work to lose)."""
        snap = getattr(self, "_clean_form", None)
        now  = self._form_snapshot()
        if snap is None or now is None:
            return False
        if self.edit_index is None:
            return now[-1] != snap[-1]
        return now != snap

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------

    def build_ui(self) -> None:
        # ── Header banner ────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=BG_HEADER, height=52)
        header.pack(fill="x")
        header.pack_propagate(False)

        logo = _load_image(LOGO_HEADER_FILE, master=header)
        if logo is not None:
            lbl = tk.Label(header, image=logo, bg=BG_HEADER, bd=0)
            lbl.image = logo
            lbl.pack(side="left", padx=(12, 12), pady=4)
            tk.Frame(header, bg="#3a4a5a", width=1).pack(side="left", fill="y", pady=12, padx=(0, 12))
        else:
            tk.Label(header, text="◈", font=("Segoe UI", 14),
                     bg=BG_HEADER, fg=ACCENT).pack(side="left", padx=(14, 6), pady=8)
        tk.Label(header, text=APP_TITLE,
                 font=("Segoe UI Semibold", 11, "bold"),
                 bg=BG_HEADER, fg="#ffffff").pack(side="left", pady=8)
        tk.Label(header, text="EPS-Sterna · EPS-Aeolus · EUMETSAT",
                 font=("Segoe UI", 8),
                 bg=BG_HEADER, fg="#7d9db5").pack(side="left", padx=(10, 0), pady=8)

        self._template_label_var = tk.StringVar(
            value=_template_label(self.template_path)
        )
        tk.Label(header, textvariable=self._template_label_var,
                 font=("Segoe UI", 8), bg=BG_HEADER, fg=ACCENT).pack(side="right", padx=14)

        # Version label — right side of header
        tk.Label(header, text=f"v{APP_VERSION}",
                 font=("Segoe UI", 8), bg=BG_HEADER, fg="#4a6a82").pack(side="right", padx=(0, 4))

        # Thin accent line under header
        tk.Frame(self.root, bg=ACCENT, height=2).pack(fill="x")

        # ── Form panel ───────────────────────────────────────────────────────
        top_card = tk.Frame(self.root, bg=BG_PANEL)
        top_card.pack(fill="x")
        top = ttk.Frame(top_card, padding=(12, 8, 12, 8))
        top.pack(fill="x", expand=True)
        top.columnconfigure(0, weight=1)

        # Row 0: Template | IO folder — both left-aligned
        r0 = ttk.Frame(top)
        r0.grid(row=0, column=0, columnspan=8, sticky="w", pady=(0, 4))

        btn_tmpl = ttk.Button(r0, text="📂  Select Word template", command=self.choose_template)
        btn_tmpl.pack(side="left")
        add_tooltip(btn_tmpl, "Load a .docx template to extract users and tasks.")

        ttk.Separator(r0, orient="vertical").pack(side="left", fill="y", padx=(14, 14))

        btn_io_dir = ttk.Button(r0, text="📁  Working folder", command=self.choose_io_dir)
        btn_io_dir.pack(side="left")
        add_tooltip(btn_io_dir, "Set the main working folder on the network share\n(weekly_reports, config and all data files live here).")

        self._io_dir_var = tk.StringVar(
            value=self.last_io_dir if self.last_io_dir else "⚠  No folder selected"
        )
        ttk.Label(r0, textvariable=self._io_dir_var, foreground=FG_DIM,
                  font=("Segoe UI", 8)).pack(side="left", padx=(8, 0))

        ttk.Separator(r0, orient="vertical").pack(side="left", fill="y", padx=(14, 14))

        btn_reload_cfg = ttk.Button(r0, text="🔄  Reload config",
                                    command=self.reload_config_from_files)
        btn_reload_cfg.pack(side="left")
        add_tooltip(btn_reload_cfg,
                    "Reload users and tasks from config/users.txt and config/tasks.txt.")

        # (template label is shown in the header bar above)

        # Row 1: inline label+field pairs in sub-frames so labels stay glued to their fields
        self.date_var = tk.StringVar(value=date.today().strftime("%d/%m/%y"))
        self.date_var.trace_add("write", self._validate_date)

        def _pair(parent, label, widget_fn, pad_right=20):
            f = ttk.Frame(parent)
            ttk.Label(f, text=label).pack(side="left", padx=(0, 4))
            w = widget_fn(f)
            w.pack(side="left")
            f.pack(side="left", padx=(0, pad_right))
            return w

        row1 = ttk.Frame(top)
        row1.grid(row=1, column=0, columnspan=8, sticky="w", pady=(6, 0))

        self._date_entry = _pair(row1, "Date",
            lambda f: ttk.Entry(f, textvariable=self.date_var, width=13))
        add_tooltip(self._date_entry, "Format: DD/MM/YY  |  Click 📅 to pick")

        # Calendar button next to date entry
        def _open_picker():
            try:
                d = safe_date(self.date_var.get())
            except Exception:
                d = date.today()
            DatePicker(self._date_entry, d,
                       lambda picked: self.date_var.set(picked.strftime("%d/%m/%y")))
        cal_btn = tk.Label(row1, text="📅", font=("Segoe UI", 11),
                           bg=BG_FRAME, cursor="hand2", fg=ACCENT)
        cal_btn.pack(side="left", padx=(2, 16))
        cal_btn.bind("<Button-1>", lambda e: _open_picker())

        self.user_var = tk.StringVar()
        self.user_cb  = _pair(row1, "User",
            lambda f: ttk.Combobox(f, textvariable=self.user_var,
                                   values=self.users, width=26, state="readonly"))
        if self.users:
            self.user_var.set(self.last_user if self.last_user in self.users else self.users[0])
        # Populate combos (may already have values from config load)
        if hasattr(self, "wp_cb") and self.tasks:
            self._populate_wp_cb()

        self.entry_type_var = tk.StringVar(value="Activity")
        self.entry_type_cb  = _pair(row1, "Type",
            lambda f: ttk.Combobox(f, textvariable=self.entry_type_var,
                                   values=ENTRY_TYPES, width=12, state="readonly"))
        self.entry_type_cb.bind("<<ComboboxSelected>>", lambda e: self.update_task_state())

        self.effort_var = tk.StringVar(value="1")
        effort_entry    = _pair(row1, "Effort (PD)",
            lambda f: ttk.Entry(f, textvariable=self.effort_var, width=7), pad_right=0)
        add_tooltip(effort_entry, "Person-days, e.g. 0.5 or 1. Must be > 0.")

        # Task row — full width
        task_row = ttk.Frame(top)
        task_row.grid(row=2, column=0, columnspan=8, sticky="we", pady=(4, 0))
        top.columnconfigure(0, weight=1)
        task_row.columnconfigure(3, weight=1)

        # WP selector
        ttk.Label(task_row, text="WP").grid(row=0, column=0, sticky="w", padx=(0, 4))
        self.wp_var = tk.StringVar()
        self.wp_cb  = ttk.Combobox(task_row, textvariable=self.wp_var,
                                    state="readonly", width=8)
        self.wp_cb.grid(row=0, column=1, padx=(0, 12))
        add_tooltip(self.wp_cb, "Select Work Package to filter the task list.")

        # Task selector
        ttk.Label(task_row, text="Task").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.task_var = tk.StringVar()
        self.task_cb  = ttk.Combobox(task_row, textvariable=self.task_var,
                                      state="readonly")
        self.task_cb.grid(row=0, column=3, sticky="we")

        # Populate WP list and bind
        self._populate_wp_cb()
        self.wp_cb.bind("<<ComboboxSelected>>", lambda e: self._on_wp_changed())

        # Description row — full width
        desc_row = ttk.Frame(top)
        desc_row.grid(row=3, column=0, columnspan=8, sticky="we", pady=(4, 0))
        desc_row.columnconfigure(1, weight=1)
        ttk.Label(desc_row, text="Description").grid(row=0, column=0, sticky="nw", padx=(0, 6))
        self.desc = tk.Text(desc_row, height=4, font=("Segoe UI", 9), wrap="word",
                            bg=BG_ENTRY, fg=FG_MAIN)
        self.desc.grid(row=0, column=1, sticky="we")

        btn_row = ttk.Frame(top)
        btn_row.grid(row=4, column=0, columnspan=8, sticky="w", pady=(8, 0))

        btn_save = ttk.Button(btn_row, text="💾  Save entry (Ctrl+S)",
                              command=self.save_activity, style="Accent.TButton")
        btn_save.pack(side="left", padx=(0, 6))

        add_tooltip(btn_save, "Add or update the selected entry.")

        btn_del = ttk.Button(btn_row, text="🗑  Delete (Del)", command=self.delete_selected)
        btn_del.pack(side="left", padx=(0, 20))
        add_tooltip(btn_del, "Permanently delete the selected entry.")

        btn_nw = ttk.Button(btn_row, text="📅  Next-week plan",
                            command=self.open_next_week_editor)
        btn_nw.pack(side="left")
        add_tooltip(btn_nw, "Edit planned activities for next week.")


        # ── Divider ─────────────────────────────────────────────────────────
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=0)

        # ── Search bar ──────────────────────────────────────────────────────
        sf_outer = tk.Frame(self.root, bg=BG_FRAME)
        sf_outer.pack(fill="x")
        sf = ttk.Frame(sf_outer, padding=(12, 5, 12, 4))
        sf.pack(fill="x")
        ttk.Label(sf, text="🔍  Filter:").pack(side="left")
        self._filter_var = tk.StringVar()
        self._filter_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self._filter_var, width=40).pack(side="left", padx=(4, 0))
        ttk.Button(sf, text="✖", width=3,
                   command=lambda: self._filter_var.set("")).pack(side="left", padx=2)
        ttk.Label(sf, text="(date, user, task, description, task no.)").pack(side="left", padx=(6, 0))

        self._counter_var = tk.StringVar(value="")
        ttk.Label(sf, textvariable=self._counter_var,
                  font=("Segoe UI", 8), foreground=FG_DIM).pack(side="right", padx=(0, 8))

        # ── Treeview ────────────────────────────────────────────────────────
        # Pack bottom FIRST so it always stays visible above the taskbar
        # then mid expands to fill remaining space
        mid = tk.Frame(self.root, bg=BG_FRAME)

        cols = ("date", "user", "type", "task", "effort", "description")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=16)
        for col, text, width, anchor in [
            ("date",        "Date",        90,  "center"),
            ("user",        "User",       160,  "w"),
            ("type",        "Type",        90,  "center"),
            ("task",        "Task",       340,  "w"),
            ("effort",      "PD",          50,  "center"),
            ("description", "Description", 410, "w"),
        ]:
            self.tree.heading(col, text=text, command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=width, anchor=anchor)
        self.tree.tag_configure("odd",        background=ROW_ODD)
        self.tree.tag_configure("even",       background=ROW_EVEN)
        self.tree.tag_configure("holiday",    background=ROW_HOLIDAY)
        self.tree.tag_configure("sick",       background=ROW_SICK)
        self.tree.tag_configure("other_week", background=ROW_OTHER_WEEK,
                                foreground="#8b6000", font=("Segoe UI", 9, "italic"))
        self.tree.tag_configure("pending",    background="#dbeafe",
                                foreground="#1e40af", font=("Segoe UI", 9, "bold"))
        self.tree.pack(side="left", fill="both", expand=True, padx=(0,0))
        self.tree.bind("<Double-1>", self.load_selected_for_edit)

        # Right-click context menu
        self._ctx_menu = tk.Menu(self.root, tearoff=0)
        self._ctx_menu.add_command(label="✏  Edit (double-click)", command=self.load_selected_for_edit)
        self._ctx_menu.add_command(label="⧉  Duplicate",           command=self.duplicate_selected)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="🗑  Delete",              command=self.delete_selected)
        self.tree.bind("<Button-3>", self._show_context_menu)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")
        self._sort_col: str = "date"
        self._sort_asc: bool = True
        self._display_index_map: dict[str, int] = {}

        # ── Bottom bar — packed BEFORE mid so it is always visible ─────────
        tk.Frame(self.root, bg=BORDER, height=1).pack(side="bottom", fill="x")
        bottom_outer = tk.Frame(self.root, bg=BG_PANEL)
        bottom_outer.pack(side="bottom", fill="x")
        bottom = ttk.Frame(bottom_outer, padding=(12, 6, 12, 6))
        bottom.pack(fill="x")

        # Now pack mid (treeview) to fill remaining space
        mid.pack(fill="both", expand=True)

        today = date.today()
        iso   = today.isocalendar()

        # ── Row 0: Weekly controls ───────────────────────────────────────────
        # Week selector — compact inline frame
        week_frame = ttk.Frame(bottom)
        week_frame.grid(row=0, column=0, sticky="w", padx=(0, 8))

        ttk.Label(week_frame, text="Year").pack(side="left")
        self.week_year = tk.StringVar(value=str(iso.year))
        ttk.Entry(week_frame, textvariable=self.week_year, width=5).pack(side="left", padx=(3, 8))

        ttk.Label(week_frame, text="CW").pack(side="left")

        btn_cw_prev = tk.Label(week_frame, text="◀", font=("Segoe UI", 10),
                               fg=ACCENT, bg=BG_PANEL, cursor="hand2")
        btn_cw_prev.pack(side="left", padx=(4, 1))
        btn_cw_prev.bind("<Button-1>", lambda e: self._step_week(-1))
        add_tooltip(btn_cw_prev, "Previous week")

        self.cw_var = tk.StringVar(value=str(iso.week))
        ttk.Entry(week_frame, textvariable=self.cw_var, width=3).pack(side="left", padx=1)

        btn_cw_next = tk.Label(week_frame, text="▶", font=("Segoe UI", 10),
                               fg=ACCENT, bg=BG_PANEL, cursor="hand2")
        btn_cw_next.pack(side="left", padx=(1, 0))
        btn_cw_next.bind("<Button-1>", lambda e: self._step_week(+1))
        add_tooltip(btn_cw_next, "Next week")

        btn_load_cw = ttk.Button(bottom, text="📂  Load week", command=self.load_week_into_table)
        btn_load_cw.grid(row=0, column=1, padx=(0, 4))
        add_tooltip(btn_load_cw, "Load a specific CW from the archive into the table for editing.")

        btn_load_cur = ttk.Button(bottom, text="⟳  Current week", command=self.load_current_week)
        btn_load_cur.grid(row=0, column=2, padx=(0, 4))
        add_tooltip(btn_load_cur, "Return to the current week's entries.")

        self._btn_save_cw = ttk.Button(bottom, text="💾  Save week",
                                       command=self.save_current_week_button)
        self._btn_save_cw.grid(row=0, column=3, padx=(0, 4))
        add_tooltip(self._btn_save_cw, "Save the current table back to the weekly archive.")

        btn_weekly = ttk.Button(bottom, text="📋  Generate weekly summary", command=self.weekly)
        btn_weekly.grid(row=0, column=4, padx=(0, 4))
        add_tooltip(btn_weekly, "Generate weekly summary text. Uses next-week plan if saved.")

        # ── Row 1: Monthly controls + Export/Import + Exit ───────────────────
        ttk.Label(bottom, text="Month YYYY-MM").grid(row=1, column=0, columnspan=1, sticky="w", pady=(4, 0))
        self.month_var = tk.StringVar(value=date.today().strftime("%Y-%m"))
        ttk.Entry(bottom, textvariable=self.month_var, width=10).grid(row=1, column=1, columnspan=1, sticky="w", padx=(2, 6), pady=(4, 0))

        btn_monthly = ttk.Button(bottom, text="📄  Generate monthly Word", command=self.monthly)
        btn_monthly.grid(row=1, column=2, padx=(0, 4), pady=(4, 0))
        add_tooltip(btn_monthly, "Aggregate all weekly archives for the selected month and fill the Word template.")

        btn_open_rpt = ttk.Button(bottom, text="📂  Open report…", command=self.open_monthly_report)
        btn_open_rpt.grid(row=1, column=3, padx=(0, 4), pady=(4, 0))
        add_tooltip(btn_open_rpt, "Open a previously generated monthly Word report.")

        btn_stats = ttk.Button(bottom, text="📊  Statistics", command=self.show_stats)
        btn_stats.grid(row=1, column=4, padx=(0, 4), pady=(4, 0))
        add_tooltip(btn_stats, "Show monthly statistics by user, task and project.")

        btn_exp = ttk.Button(bottom, text="⬆  Export", command=self._export_menu)
        btn_exp.grid(row=1, column=5, padx=(0, 4), pady=(4, 0))
        add_tooltip(btn_exp, "Export current table to CSV, JSON or Excel.")

        btn_imp = ttk.Button(bottom, text="⬇  Import", command=self.import_table)
        btn_imp.grid(row=1, column=6, padx=(0, 4), pady=(4, 0))
        add_tooltip(btn_imp, "Import entries from CSV or JSON.")

        ttk.Separator(bottom, orient="vertical").grid(row=0, column=7, rowspan=2, sticky="ns", padx=(12, 8))

        btn_exit = ttk.Button(bottom, text="✖  Exit", command=self.on_close, style="Danger.TButton")
        btn_exit.grid(row=0, column=7, rowspan=2, padx=(0, 4), sticky="ns")
        add_tooltip(btn_exit, "Close the application.")

        # ── Status bar — packed last at very bottom ──────────────────────────
        self._status_var = tk.StringVar()
        status_bar = tk.Frame(self.root, bg=BG_HEADER, height=22)
        status_bar.pack(side="bottom", fill="x")
        status_bar.pack_propagate(False)
        tk.Label(status_bar, textvariable=self._status_var,
                 anchor="w", fg=ACCENT, bg=BG_HEADER,
                 font=("Segoe UI", 8), padx=12).pack(side="left", fill="y")
        # Version tag right side
        tk.Label(status_bar, text=APP_TITLE,
                 anchor="e", fg="#7d9db5", bg=BG_HEADER,
                 font=("Segoe UI", 7), padx=12).pack(side="right", fill="y")

        self.update_task_state()
        self._mark_clean()   # baseline for the unsaved-changes check

    # -----------------------------------------------------------------------
    # Real-time date validation
    # -----------------------------------------------------------------------

    def _validate_date(self, *_) -> None:
        val = self.date_var.get()
        ok = False
        for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                datetime.strptime(val, fmt)
                ok = True
                break
            except ValueError:
                pass
        # Style was configured once at startup — just switch (BUG H fix)
        self._date_entry.configure(style="TEntry" if ok else "Error.TEntry")

    # -----------------------------------------------------------------------
    # Treeview helpers
    # -----------------------------------------------------------------------

    def _sort_by(self, col: str) -> None:
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self.refresh()

    def refresh(self) -> None:
        self.data    = normalise_data(self.data)
        filt         = self._filter_var.get().lower() if hasattr(self, "_filter_var") else ""

        # Build display list with original data indices preserved
        if filt:
            indexed = [
                (orig_idx, a) for orig_idx, a in enumerate(self.data)
                if filt in str(a.get("date", "")).lower()
                or filt in str(a.get("user", "")).lower()
                or filt in str(a.get("task_title", "")).lower()
                or filt in str(a.get("task_no", "")).lower()
                or filt in str(a.get("description", "")).lower()
                or filt in str(a.get("entry_type", "")).lower()
                # Also search date in European format
                or (lambda d: filt in d)(
                    _try_format_date(a.get("date", ""))
                )
            ]
        else:
            indexed = list(enumerate(self.data))

        if self._sort_col:
            key_map = {
                "date": "date", "user": "user", "type": "entry_type",
                "task": "task_title", "effort": "effort", "description": "description",
            }
            field = key_map.get(self._sort_col, self._sort_col)
            def _sort_key(x):
                v = x[1].get(field, "")
                if field == "effort":
                    try:
                        return (0, float(v), "")
                    except (TypeError, ValueError):
                        pass
                return (1, 0.0, str(v))
            try:
                indexed.sort(key=_sort_key, reverse=not self._sort_asc)
            except Exception:
                pass

        for i in self.tree.get_children():
            self.tree.delete(i)

        # Store mapping: tree iid -> original data index
        self._display_index_map: dict[str, int] = {}

        total    = len(self.data)
        shown    = len(indexed)
        shown_pd = sum(float(a.get("effort", 0) or 0)
                       for _, a in indexed
                       if a.get("entry_type") == "Activity")
        # Show selected CW from selector, not today's CW
        try:
            sel_year = int(self.week_year.get())
            sel_cw   = int(self.cw_var.get())
            cw_str   = f"CW{sel_cw:02d} / {sel_year}"
        except Exception:
            iso    = date.today().isocalendar()
            cw_str = f"CW{iso.week:02d} / {iso.year}"
        if shown == total:
            counter = f"{shown} entries  |  {shown_pd:g} PD"
        else:
            counter = f"{shown} of {total} entries  |  {shown_pd:g} PD (filtered)"
        counter += f"  |  {cw_str}"
        # Always update counter (use a dedicated label, not status bar)
        if hasattr(self, "_counter_var"):
            self._counter_var.set(counter)

        for tree_idx, (orig_idx, a) in enumerate(indexed):
            entry_type   = a.get("entry_type", "Activity")
            if entry_type == "Activity":
                wp_d = a.get("task_wp", "")
                task_display = (
                    f"{wp_d+' - ' if wp_d else ''}"
                    f"{a.get('task_no', '')} - {a.get('task_title', '')}"
                )
            else:
                task_display = a.get("task_title", "")
            if entry_type == "Holiday":
                tag = "holiday"
            elif entry_type == "Sick Leave":
                tag = "sick"
            else:
                # Check if entry belongs to a different week than working week
                is_other_week = False
                try:
                    # Only highlight entries explicitly added after load
                    # AND whose date belongs to a different week
                    if activity_key(a) in self._new_entry_keys and self._working_week:
                        entry_yw  = iso_year_week_for_date(safe_date(a["date"]))
                        is_other_week = (entry_yw != self._working_week)
                except Exception:
                    pass
                if is_other_week:
                    tag = "other_week"
                elif orig_idx == self._pending_duplicate_idx:
                    tag = "pending"
                else:
                    tag = "odd" if tree_idx % 2 else "even"

            iid = str(tree_idx)
            self._display_index_map[iid] = orig_idx

            # Prefix date with ⚠ if entry is from a different week
            date_display = a.get("date", "")
            if tag == "other_week":
                try:
                    yw = iso_year_week_for_date(safe_date(date_display))
                    date_display = f"⚠ {date_display} (CW{yw[1]:02d})"
                except Exception:
                    date_display = f"⚠ {date_display}"

            self.tree.insert("", "end", iid=iid,
                values=(
                    date_display, a.get("user", ""), entry_type,
                    task_display, format_effort(a.get("effort", 0)), a.get("description", "")[:120],
                ),
                tags=(tag,),
            )

    def selected_index(self) -> int | None:
        """Return the index in self.data of the selected row (filter-aware)."""
        sel = self.tree.selection()
        if not sel:
            return None
        iid = sel[0]
        # Use display->data mapping if available (handles filtered view)
        if hasattr(self, "_display_index_map"):
            return self._display_index_map.get(iid)
        return int(iid)

    # -----------------------------------------------------------------------
    # Form helpers
    # -----------------------------------------------------------------------

    def update_task_state(self) -> None:
        if getattr(self, "entry_type_var", None) and self.entry_type_var.get() == "Activity":
            self.task_cb.configure(state="readonly")
            if hasattr(self, "wp_cb"):
                self.wp_cb.configure(state="readonly")
        else:
            self.task_cb.configure(state="disabled")
            if hasattr(self, "wp_cb"):
                self.wp_cb.configure(state="disabled")

    def _get_wps(self) -> list[str]:
        """Return sorted list of unique WPs from self.tasks.
        Handles both (n, t, wp) and legacy (n, t) tuples."""
        seen = set()
        for entry in self.tasks:
            wp = entry[2] if len(entry) > 2 else ""
            if wp:
                seen.add(wp)
        # Sort: WP* first (WP6, WP7...), then others (SLM)
        wps = sorted(seen, key=lambda w: (0 if w.upper().startswith("WP") else 1, w))
        return wps

    def _populate_wp_cb(self) -> None:
        """Populate WP combobox and set default."""
        wps = self._get_wps()
        self.wp_cb["values"] = wps
        if wps:
            # Default: first non-SLM WP, or first WP
            default = next((w for w in wps if w.upper() != "SLM"), wps[0])
            self.wp_var.set(default)
        else:
            # tasks.txt without WP column — show every task
            self.wp_var.set("")
        self._on_wp_changed(init=True)

    def _on_wp_changed(self, init: bool = False) -> None:
        """Filter task list to selected WP."""
        wp = self.wp_var.get()
        filtered = []
        for entry in self.tasks:
            n, t = entry[0], entry[1]
            w = entry[2] if len(entry) > 2 else ""
            if w == wp:
                filtered.append((n, t, w))
        task_values = [f"{n} - {t}" for n, t, w in filtered]
        self.task_cb["values"] = task_values
        if task_values:
            if init or self.task_var.get() not in task_values:
                self.task_var.set(task_values[0])
        else:
            self.task_var.set("")

    def _update_wp_and_task_from_entry(self, a: dict) -> None:
        """Restore WP and task dropdowns from a loaded entry. The entry's task
        must survive even if its title changed in tasks.txt since it was saved —
        otherwise Save would silently move it to whatever task is selected."""
        if a.get("entry_type") != "Activity":
            return
        wp = entry_wp(a)
        if hasattr(self, "wp_var") and wp != self.wp_var.get():
            self.wp_var.set(wp)
            self._on_wp_changed()
        values = list(self.task_cb["values"])
        task_str = f"{a.get('task_no')} - {a.get('task_title')}"
        if task_str not in values:
            # Title changed in tasks.txt: same task number in this WP
            same_no = [v for v in values if v.split(" - ", 1)[0].strip() == str(a.get("task_no"))]
            if same_no:
                task_str = same_no[0]
            # else: keep the entry's own task (not in the list any more)
        self.task_var.set(task_str)

    def parse_task(self) -> tuple[int, str, str]:
        """Parse task_var and wp_var. Returns (task_no, task_title, task_wp)."""
        import re
        s  = self.task_var.get().strip()
        wp = self.wp_var.get().strip() if hasattr(self, "wp_var") else ""
        # task_var now shows "N - Title" (WP is separate)
        m = re.match(r"^(\d+)\s*-\s*(.*)$", s)
        if m:
            return int(m.group(1)), m.group(2).strip(), wp
        raise ValueError(f"Unrecognised task format: {s!r}")

    # -----------------------------------------------------------------------
    # Template
    # -----------------------------------------------------------------------

    def load_template(self, path, silent: bool = False) -> None:
        """Set the template path only — users and tasks always come from config files."""
        try:
            self.template_path = Path(path)
            # Do NOT load users/tasks from template — use config/users.txt and tasks.txt
            if not silent:
                messagebox.showinfo(
                    "Template loaded",
                    f"Template set: {Path(path).name}\n\n"
                    f"Users and tasks are loaded from config files, not from the template.",
                )
        except Exception as e:
            log.error("load_template failed: %s", e)
            messagebox.showerror("Template error", str(e))

    def choose_template(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
        if p:
            self.load_template(p)
            self._template_label_var.set(_template_label(Path(p)))
            save_app_state(
                self.template_path, self.users, self.tasks,
                self.user_var.get() if hasattr(self, "user_var") else "",
                self.last_io_dir,
            )
            self._status(f"Template: {Path(p).name}")

    # -----------------------------------------------------------------------
    # CRUD
    # -----------------------------------------------------------------------

    def save_activity(self) -> None:
        if not self._writable():
            return
        if not require_io_dir():
            return
        try:
            entry_date = safe_date(self.date_var.get())
        except Exception:
            messagebox.showerror("Input error", "Date must be DD/MM/YY (or DD/MM/YYYY).")
            return
        try:
            effort = float(self.effort_var.get().replace(",", "."))
            if effort <= 0:
                raise ValueError("effort must be > 0")
        except Exception as e:
            messagebox.showerror("Input error", f"Invalid effort value.\n{e}")
            return
        # Warn if daily effort exceeds 1.0 PD
        if effort > 1.0:
            if not messagebox.askyesno(
                "Effort warning",
                f"Effort is {effort} PD — this exceeds one full working day.\n\nSave anyway?",
            ):
                return
        # Warn if total effort for this date+user would exceed 1.0 PD,
        # leave included (when editing, the entry's old value is excluded)
        try:
            same_day = [
                a for i, a in enumerate(self.data)
                if i != self.edit_index
                and a.get("date") == entry_date.isoformat()
                and a.get("user") == self.user_var.get()
            ]
            day_total = sum(float(a.get("effort", 0) or 0) for a in same_day) + effort
            if day_total > 1.0:
                if not messagebox.askyesno(
                    "Daily total warning",
                    f"Total effort for {entry_date.isoformat()} ({self.user_var.get()}) "
                    f"would be {day_total:g} PD.\n\nSave anyway?",
                ):
                    return
        except Exception:
            pass
        # No week restriction — entries from any week can coexist in the table.
        # Auto-archiving happens automatically on startup when a new week begins.
        if not self.user_var.get().strip():
            messagebox.showerror("Input error", "Select a user (see config/users.txt).")
            return
        entry_type = self.entry_type_var.get()
        if entry_type == "Activity":
            try:
                task_no, task_title, task_wp = self.parse_task()
            except Exception as e:
                messagebox.showerror("Input error", f"Check task.\n{e}")
                return
        else:
            task_no, task_title, task_wp = 0, entry_type, ""
        item = {
            "date":        entry_date.isoformat(),
            "user":        self.user_var.get(),
            "entry_type":  entry_type,
            "task_no":     task_no,
            "task_title":  task_title,
            "task_wp":     task_wp,
            "effort":      effort,
            "description": self.desc.get("1.0", "end").strip(),
        }
        # Check date against working week BEFORE modifying self.data
        entry_yw = iso_year_week_for_date(entry_date)
        if self._working_week and entry_yw != self._working_week:
            y_w, cw_w = self._working_week
            y_e, cw_e = entry_yw
            if not messagebox.askyesno(
                "Different week",
                f"This entry belongs to CW{cw_e:02d}/{y_e}\n"
                f"but you are currently working on CW{cw_w:02d}/{y_w}.\n\n"
                f"It will be saved to CW{cw_e:02d}/{y_e} on next Save week.\n\nContinue?",
            ):
                return

        # Now safe to modify self.data
        if self.edit_index is None:
            # Brand new entry — always append
            self._push_undo("add entry")
            self.data.append(item)
            self._new_entry_keys.add(activity_key(item))
            self._status("Entry added.")
        else:
            # Editing existing entry
            old_item = self.data[self.edit_index]
            old_key  = activity_key(old_item)
            is_different_week = (
                self._working_week is not None and
                entry_yw != self._working_week
            )
            if is_different_week:
                self._push_undo("edit entry (moved week)")
                self.data.pop(self.edit_index)
                self.edit_index = None
                self.data.append(item)
                self._status("Entry updated (moved to different week).")
            else:
                self._push_undo("edit entry")
                self.data[self.edit_index] = item
                self._status("Entry updated.")
            # Update _new_entry_keys: if old key was tracked, replace with new key
            if old_key in self._new_entry_keys:
                self._new_entry_keys.discard(old_key)
                self._new_entry_keys.add(activity_key(item))

        # Auto-set working week from first entry if not already set
        if not self._working_week:
            self._working_week = entry_yw
        self._persist()
        auto_backup(self.data)
        self._mark_clean()
        self._set_table_dirty()
        self._pending_duplicate_idx = None  # clear pending highlight
        self.cancel_edit(clear_desc=True)
        self.refresh()
        self.root.update_idletasks()  # force immediate redraw

    def load_selected_for_edit(self, event=None) -> None:
        idx = self.selected_index()
        if idx is None:
            return
        if self._form_dirty() and idx != self.edit_index:
            if not messagebox.askyesno(
                "Discard changes",
                "The form has changes that are not saved.\nDiscard them and load the selected entry?",
            ):
                return
        a = normalise_activity(self.data[idx])
        self.edit_index = idx
        try:
            d = safe_date(a.get("date", date.today().isoformat()))
            self.date_var.set(d.strftime("%d/%m/%y"))
        except Exception:
            self.date_var.set(date.today().strftime("%d/%m/%y"))
        self.user_var.set(a.get("user", ""))
        self.entry_type_var.set(a.get("entry_type", "Activity"))
        if a.get("entry_type") == "Activity":
            self._update_wp_and_task_from_entry(a)
        self.effort_var.set(_effort_text(a.get("effort", 1)))   # no rounding
        self.desc.delete("1.0", "end")
        self.desc.insert("1.0", a.get("description", ""))
        self._mark_clean()
        self.update_task_state()

    def cancel_edit(self, clear_desc: bool = False) -> None:
        if not clear_desc and self._form_dirty() and self.edit_index is not None:
            if not messagebox.askyesno(
                "Cancel edit",
                "You have unsaved changes in the form.\nDiscard them?",
            ):
                return
        self.edit_index = None   # always reset
        if clear_desc:
            self.desc.delete("1.0", "end")
            # Reset date to today or working week Monday
            if self._working_week:
                today_yw = iso_year_week_for_date(date.today())
                if self._working_week == today_yw:
                    self.date_var.set(date.today().strftime("%d/%m/%y"))
                else:
                    self.date_var.set(
                        iso_week_monday(*self._working_week).strftime("%d/%m/%y")
                    )
            else:
                self.date_var.set(date.today().strftime("%d/%m/%y"))
            # Reset task and effort to defaults
            self.effort_var.set("1")
            self.entry_type_var.set("Activity")
            if hasattr(self, "wp_cb"):
                self._populate_wp_cb()   # resets WP and task to first valid option
        self._mark_clean()   # after the reset, so default values don't count as edits

    def duplicate_selected(self) -> None:
        """Copy the selected entry into the form as a NEW entry. Nothing is
        added until Save entry, so an abandoned copy can't be counted twice."""
        idx = self.selected_index()
        if idx is None:
            return
        if self._form_dirty() and not messagebox.askyesno(
            "Discard changes",
            "The form has changes that are not saved.\nDiscard them and copy the selected entry?",
        ):
            return
        item = normalise_activity(self.data[idx])
        self.edit_index = None   # Save entry will ADD
        # Date: today if it is in the working week, else that week's Monday
        target = date.today()
        if self._working_week and iso_year_week_for_date(target) != self._working_week:
            target = iso_week_monday(*self._working_week)
        self.date_var.set(target.strftime("%d/%m/%y"))
        self.user_var.set(item.get("user", ""))
        self.entry_type_var.set(item.get("entry_type", "Activity"))
        if item.get("entry_type") == "Activity":
            self._update_wp_and_task_from_entry(item)   # restores WP too
        self.effort_var.set(_effort_text(item.get("effort", 1)))   # no rounding
        self.desc.delete("1.0", "end")
        self.desc.insert("1.0", item.get("description", ""))
        self.update_task_state()
        self._mark_clean()
        self._status("Copy loaded in the form — adjust date/description and click Save entry.")

    def _show_context_menu(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def delete_selected(self) -> None:
        idx = self.selected_index()
        if idx is None or not self._writable():
            return
        if messagebox.askyesno("Delete", "Delete selected entry?"):
            self._push_undo("delete entry")
            self.data.pop(idx)
            # Keep the entry being edited pointing at the right row
            if self.edit_index is not None:
                if self.edit_index == idx:
                    self.edit_index = None
                    self._mark_clean()
                elif self.edit_index > idx:
                    self.edit_index -= 1
            if self._pending_duplicate_idx is not None and self._pending_duplicate_idx >= idx:
                self._pending_duplicate_idx = None
            self._persist()
            self._set_table_dirty()
            self.refresh()
            self._status("Entry deleted.")

    # -----------------------------------------------------------------------
    # Archive / reset
    # -----------------------------------------------------------------------

    def reset_data(self) -> None:
        if not require_io_dir() or not self._writable():
            return
        if not messagebox.askyesno(
            "Reset",
            "The table will be cleared.\n\n"
            "Entries already saved to the weekly archive are NOT deleted and still "
            "count in the monthly report — to remove them, load the week and delete "
            "them there.\n\nContinue?",
        ):
            return
        try:
            self._push_undo("reset")
            self.data = []
            self.edit_index = None
            self._working_week = None
            self._loaded_week_keys = set()
            self._clear_table_dirty()
            self._persist()
            self.refresh()
            self.cancel_edit(clear_desc=True)   # resets the form and its baseline
            self._status("Table cleared.")
        except Exception as e:
            log.error("reset_data: %s", e)
            messagebox.showerror("Reset error", str(e))

    def _step_week(self, delta: int) -> None:
        """Advance or retreat the CW selector by delta weeks (+1 or -1)."""
        try:
            year = int(self.week_year.get())
            cw   = int(self.cw_var.get())
        except ValueError:
            return
        # Compute new date by stepping 7 days
        current_monday = iso_week_monday(year, cw)
        new_monday     = current_monday + timedelta(days=7 * delta)
        new_iso        = new_monday.isocalendar()
        self.week_year.set(str(new_iso.year))
        self.cw_var.set(str(new_iso.week))

    def load_week_into_table(self) -> None:
        """Load a specific CW from the weekly archive into the table for editing."""
        if not require_io_dir():
            return
        try:
            year = int(self.week_year.get())
            cw   = int(self.cw_var.get())
            iso_week_monday(year, cw)   # e.g. CW54 does not exist
        except Exception:
            messagebox.showerror("Input error", "Please set a valid Year and CW first.")
            return

        # Checked first: both branches below replace the table
        if self._table_dirty:
            if not messagebox.askyesno(
                "Unsaved changes",
                f"You have unsaved changes.\n\nLoading CW{cw:02d}/{year} will discard them.\n\nContinue?",
            ):
                return

        week_file = weekly_dir() / f"week_{year}_CW{cw:02d}.json"
        if not week_file.exists():
            if not messagebox.askyesno(
                "Week not found",
                f"No archive found for CW{cw:02d}/{year}.\n\n"
                f"Create a new empty week and start entering data?",
            ):
                return
            # Create empty week and set as working week
            self._push_undo("create new week")
            self.edit_index = None
            self._mark_clean()
            self.data = []
            self._working_week = (year, cw)
            self._loaded_week_keys = set()
            self._persist()
            self._clear_table_dirty()
            self.refresh()
            self.week_year.set(str(year))
            self.cw_var.set(str(cw))
            mon = iso_week_monday(year, cw)
            self.date_var.set(mon.strftime("%d/%m/%y"))
            self._status(f"New week CW{cw:02d}/{year} ready — "
                         f"{mon.strftime('%d %b')} → {(mon + timedelta(days=4)).strftime('%d %b %Y')}")
            return

        try:
            loaded = normalise_data(json.loads(week_file.read_text(encoding="utf-8")))
            self._push_undo("load week")
            self.edit_index = None          # BUG C fix
            self._mark_clean()
            self.data = loaded
            # Track which entries were present when week was loaded
            # so deletions are respected on Save week
            self._loaded_week_keys = {_stable_key_fn(a) for a in loaded}
            self._new_entry_keys  = set()   # reset new entries tracker
            self._working_week = (year, cw)
            self._persist()
            self._clear_table_dirty()   # clean after load — dirty only on actual edits
            self.refresh()
            self._status(f"Loaded CW{cw:02d}/{year} — {len(loaded)} entries. Edit and click 💾 Save week.")
        except Exception as e:
            log.error("load_week_into_table: %s", e)
            messagebox.showerror("Load error", str(e))

    def load_current_week(self) -> None:
        """Load current week — same logic as load_week_into_table but for today."""
        if not require_io_dir():
            return
        if self._table_dirty:
            if not messagebox.askyesno(
                "Switch to current week",
                "You have unsaved changes.\n\nSwitch to current week anyway?",
            ):
                return
        try:
            iso_today  = date.today().isocalendar()
            current_yw = (iso_today.year, iso_today.week)
            y, cw      = current_yw

            # Source 1: weekly archive
            week_file   = weekly_dir() / f"week_{y}_CW{cw:02d}.json"
            from_archive: list = []
            if week_file.exists():
                try:
                    from_archive = normalise_data(
                        json.loads(week_file.read_text(encoding="utf-8"))
                    )
                except Exception:
                    pass

            # Source 2: activities_current.json — unsaved in-progress entries
            from_current: list = []
            f = data_file()
            if f.exists():
                try:
                    all_current = normalise_data(
                        json.loads(f.read_text(encoding="utf-8"))
                    )
                    for a in all_current:
                        try:   # one unreadable date must not drop the whole week
                            if iso_year_week_for_date(safe_date(a["date"])) == current_yw:
                                from_current.append(a)
                        except ValueError:
                            pass
                except Exception as exc:
                    log.warning("load_current_week: %s", exc)

            # Merge: in-progress entries replace their archived versions
            merged = normalise_data(merge_entries(from_archive, from_current))
            merged.sort(key=lambda a: str(a.get("date", "")))

            self._push_undo("load current week")
            self.edit_index = None
            self._mark_clean()
            self.data = merged
            self._working_week     = current_yw
            self._loaded_week_keys = {_stable_key_fn(a) for a in self.data}
            self._new_entry_keys   = set()   # reset new entries tracker

            self._persist()
            self.refresh()
            self.week_year.set(str(y))
            self.cw_var.set(str(cw))
            self._clear_table_dirty()

            if self.data:
                self._status(f"Current week loaded — CW{cw:02d}/{y}, {len(self.data)} entries")
            else:
                self._status(f"Ready for CW{cw:02d}/{y} — no entries yet")
        except Exception as e:
            log.error("load_current_week: %s", e)
            messagebox.showerror("Load error", str(e))

    def save_current_week_button(self) -> None:
        if not self._writable():
            return
        try:
            # Group table entries by calendar week
            groups: dict[tuple, list] = {}
            for a in normalise_data(self.data):
                try:
                    key = iso_year_week_for_date(safe_date(a["date"]))
                except Exception:
                    continue
                groups.setdefault(key, []).append(a)
            # Weeks where entries were deleted must be saved too, even if the
            # table has nothing left for them
            weeks = set(groups) | {w for w in map(_week_of_key, self._loaded_week_keys) if w}

            if not weeks:
                messagebox.showinfo("No data", "No entries to save.")
                return

            # Warn if entries span multiple weeks
            if len(groups) > 1:
                weeks_str = ", ".join(f"CW{cw:02d}/{y}" for y, cw in sorted(groups))
                if not messagebox.askyesno(
                    "Multiple weeks detected",
                    f"The table contains entries from {len(groups)} different weeks:\n"
                    f"{weeks_str}\n\n"
                    f"Each group will be merged into its own archive file.\n\nContinue?",
                ):
                    return

            saved = []
            for y, cw in sorted(weeks):
                # Table entries win, loaded-then-deleted entries are removed,
                # other users' entries are kept
                _, count = write_week_archive(y, cw, groups.get((y, cw), []),
                                              self._loaded_week_keys)
                saved.append(f"CW{cw:02d}/{y} ({count} entries)")

            auto_backup(self.data, force=True)   # always backup on Save week
            self._clear_table_dirty()
            self._status("Saved: " + "  |  ".join(saved))

            # Reload only the working week's entries into the table
            # (entries from other weeks have been routed to their own archive files)
            if self._working_week:
                target_yw = self._working_week
                filtered_data = []
                for a in normalise_data(self.data):
                    try:
                        if iso_year_week_for_date(safe_date(a["date"])) == target_yw:
                            filtered_data.append(a)
                    except Exception:
                        filtered_data.append(a)   # unreadable date: keep, never drop
                self.data = filtered_data
                self._new_entry_keys   = set()   # after save, nothing is "new" anymore
                self.edit_index = None
                self._mark_clean()
            # The archive now matches the table
            self._loaded_week_keys = {_stable_key_fn(a) for a in self.data}
            self._persist()
            self.refresh()
        except Exception as e:
            log.error("save_current_week_button: %s", e)
            messagebox.showerror("Save error", str(e))

    # -----------------------------------------------------------------------
    # Export / import
    # -----------------------------------------------------------------------

    def _export_menu(self) -> None:
        self.export_table()

    def choose_io_dir(self) -> None:
        """Let the user pick the default import/export folder."""
        d = filedialog.askdirectory(
            title="Select main working folder (network share)",
            initialdir=self.last_io_dir or None,
        )
        if not d:
            return
        # If user accidentally selected the weekly_reports subfolder, step up
        p = Path(d)
        if p.name.lower() == "weekly_reports":
            d = str(p.parent)

        release_lock()   # lock of the previous folder (no-op if not ours)
        set_io_dir(d)
        self.last_io_dir = str(get_io_root())
        self._io_dir_var.set(self.last_io_dir)
        setup_log_file()

        # Load config and holidays
        self.users, self.tasks = reload_config()
        reload_holidays()
        self._update_combos()

        # Acquire lock. show_lock_error exits on "Exit"; after "Force Unlock"
        # we must go on and load THIS folder's data (returning here kept the
        # previous folder's table, which the next save wrote into this folder)
        if getattr(sys, "frozen", False) and not acquire_lock():
            show_lock_error(self.root, read_lock_info())

        # Load data from share (same as startup)
        self.edit_index = None
        self._working_week = None   # belongs to the previous folder
        self._undo_stack.clear()   # snapshots belong to the previous folder
        self.data, was_reset, saved_file = read_json()
        if data_load_warning:
            messagebox.showwarning("Data file problem", data_load_warning)
        self._loaded_week_keys = load_table_state() | {_stable_key_fn(a) for a in self.data}
        self._table_dirty_after_load()
        if self.data:
            try:
                ref = max(safe_date(a["date"]) for a in self.data if a.get("date"))
                self._working_week = iso_year_week_for_date(ref)
            except Exception:
                pass
        self.refresh()

        # Check for new week
        if was_reset:
            msg = "Previous week archived automatically."
            if saved_file:
                msg += f"\n\nArchived to:\n{saved_file}"
            messagebox.showinfo("Week auto-archived", msg)

        # Save state
        save_app_state(
            self.template_path, self.users, self.tasks,
            self.user_var.get() if hasattr(self, "user_var") else "",
            self.last_io_dir,
        )

        # Version check — now that io_dir is known
        _latest, _update_available = check_version()
        if _update_available:
            self.root.after(300, lambda l=_latest: self._show_update_notice(l, APP_VERSION))

        self._status(f"Working folder set: {self.last_io_dir}")

    def export_table(self) -> None:
        if not self.data:
            messagebox.showinfo("No data", "No entries to export."); return
        default = f"activity_entries_{date.today().strftime('%Y%m%d')}.csv"
        types   = [("CSV files", "*.csv"), ("JSON files", "*.json")]
        if HAS_OPENPYXL:
            types.append(("Excel files", "*.xlsx"))
        out = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile=default, filetypes=types,
            initialdir=self.last_io_dir or None,
        )
        if not out:
            return
        try:
            if out.lower().endswith(".json"):
                Path(out).write_text(json.dumps(normalise_data(self.data), indent=2, ensure_ascii=False), encoding="utf-8")
            elif out.lower().endswith(".xlsx"):
                export_to_xlsx(self.data, out)
            else:
                fields = ["date", "user", "entry_type", "task_wp", "task_no", "task_title", "effort", "description"]
                with open(out, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore", delimiter=";")
                    writer.writeheader()
                    for row in normalise_data(self.data):
                        writer.writerow(row)
            self._status(f"Exported → {Path(out).name}")
        except Exception as e:
            log.error("export_table: %s", e)
            messagebox.showerror("Export error", str(e))

    def export_xlsx(self) -> None:
        if not HAS_OPENPYXL:
            messagebox.showerror("Missing library", "openpyxl is not installed.\nRun: pip install openpyxl")
            return
        if not self.data:
            messagebox.showinfo("No data", "No entries to export."); return
        default = f"activity_entries_{date.today().strftime('%Y%m%d')}.xlsx"
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default,
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=self.last_io_dir or None,
        )
        if not out:
            return
        try:
            export_to_xlsx(self.data, out)
            self._status(f"Exported → {Path(out).name}")
        except Exception as e:
            log.error("export_xlsx: %s", e)
            messagebox.showerror("Export error", str(e))

    def import_table(self) -> None:
        if not self._writable():
            return
        path = filedialog.askopenfilename(
            title="Import from another user",
            filetypes=[("Activity table", "*.csv *.json"), ("CSV files", "*.csv"), ("JSON files", "*.json")],
            initialdir=self.last_io_dir or None,
        )
        if not path:
            return
        try:
            imported, rejected = import_activity_table(path)
            for item in imported:   # files from older versions carry no WP
                if not item.get("task_wp"):
                    item["task_wp"] = entry_wp(item)
            existing_keys = {activity_key(a) for a in self.data}
            added = skipped = 0
            self._push_undo("import")
            for item in imported:
                key = activity_key(item)
                if key in existing_keys:
                    skipped += 1; continue
                self.data.append(item); existing_keys.add(key); added += 1
            self.last_io_dir = str(Path(path).parent)
            # Imported entries now count as loaded. Union, not reassignment:
            # rebuilding from self.data would forget entries deleted earlier
            # in this session, and Save week would bring them back.
            self._loaded_week_keys |= {_stable_key_fn(a) for a in imported}
            self._persist()
            self._set_table_dirty()
            self.refresh()
            self._status(f"Import done: {added} added, {skipped} skipped, {len(rejected)} rejected.")
            msg = (f"File: {Path(path).name}\n\nAdded: {added}\nSkipped (duplicates): {skipped}")
            if rejected:
                msg += (f"\nRejected (invalid): {len(rejected)}\n\n"
                        + "\n".join(rejected[:15])
                        + ("\n…" if len(rejected) > 15 else ""))
                messagebox.showwarning("Import completed with errors", msg)
            else:
                messagebox.showinfo("Import completed", msg)
        except Exception as e:
            log.error("import_table: %s", e)
            messagebox.showerror("Import error", str(e))

    # -----------------------------------------------------------------------
    # Next-week editor
    # -----------------------------------------------------------------------

    def open_next_week_editor(self) -> None:
        # Plan the week after the CW selected for the weekly summary
        try:
            base = (int(self.week_year.get()), int(self.cw_var.get()))
            iso_week_monday(*base)
        except (ValueError, TypeError):
            base = iso_year_week_for_date(date.today())
        NextWeekEditor(self.root, self.tasks, week_after(*base))

    # -----------------------------------------------------------------------
    # Statistics
    # -----------------------------------------------------------------------

    def show_stats(self) -> None:
        try:
            y, m = map(int, self.month_var.get().split("-"))
        except Exception:
            messagebox.showerror("Input error", "Month must be YYYY-MM."); return
        all_data = merge_week(load_monthly_data_from_weekly_reports(y, m),
                              self._table_entries_in_month(y, m), self._loaded_week_keys)
        if skipped_archives:
            messagebox.showwarning("Unreadable weekly files",
                                   "These weekly files could not be read and are NOT included:\n\n"
                                   + "\n".join(skipped_archives))
        StatsWindow(self.root, all_data, y, m)

    def _table_entries_in_month(self, y: int, m: int) -> list:
        out = []
        for item in normalise_data(self.data):
            try:
                d = safe_date(item["date"])
            except Exception:
                continue
            if date(y, m, 1) <= d <= last_day_of_month(y, m):
                out.append(item)
        return out

    # -----------------------------------------------------------------------
    # Report generation
    # -----------------------------------------------------------------------

    def weekly(self) -> None:
        if not require_io_dir():
            return
        try:
            year = int(self.week_year.get())
            cw   = int(self.cw_var.get())
            start, end = cw_range(year, cw)
            # Archive the table first (each week to its own file) ...
            groups: dict[tuple, list] = {}
            for a in normalise_data(self.data):
                try:
                    key = iso_year_week_for_date(safe_date(a["date"]))
                    groups.setdefault(key, []).append(a)
                except Exception:
                    pass
            saved_file = None
            for (gy, gcw), gentries in (groups.items() if not read_only else []):
                saved_file, _ = write_week_archive(gy, gcw, gentries, self._loaded_week_keys)
            # Archived entries count as loaded, so later deletions reach the archive
            self._loaded_week_keys |= {_stable_key_fn(a) for a in self.data}
            self._persist()
            # ... then summarise the selected week from the archive, so it also
            # works for a week that is not loaded in the table
            nxt = week_after(year, cw)
            week_data = merge_entries(load_week_entries(year, cw) + load_week_entries(*nxt),
                                      normalise_data(self.data))
            plan = load_next_week_plan(nxt)
            txt  = generate_weekly_text(
                week_data, start.isoformat(), end.isoformat(),
                cw=cw, year=year, next_week_plan=plan or None,
            )
        except Exception as e:
            log.error("weekly: %s", e)
            messagebox.showerror("Generation error", str(e)); return

        win = tk.Toplevel(self.root)
        win.title(f"Weekly summary – CW{cw:02d}/{year}")
        win.geometry(f"820x{min(440, self.root.winfo_screenheight()-120)}")
        t = tk.Text(win, font=("Consolas", 10), wrap="word")
        t.pack(fill="both", expand=True, padx=10, pady=10)
        t.insert("1.0", txt)

        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append(t.get("1.0", "end"))
            self._status("Weekly summary copied to clipboard.")

        ttk.Button(win, text="📋  Copy to clipboard", command=copy_all).pack(pady=(0, 8))
        if saved_file:
            self._status(f"Weekly archive saved → {saved_file.name}")

    def open_monthly_report(self) -> None:
        """Open file dialog to select and open a previously generated monthly report."""
        import subprocess, os
        # Default to OneDrive monthly_reports folder, fallback to last_io_dir
        _candidates = [
            Path.home() / "OneDrive - eumetsat.int" / "Exostaff Report Generator" / "monthly_reports",
            Path.home() / "OneDrive" / "Exostaff Report Generator" / "monthly_reports",
        ]
        if self.last_io_dir:
            _candidates.append(Path(self.last_io_dir) / "monthly_reports")
        default_dir = next((str(c) for c in _candidates if c.exists()), None)
        if not default_dir and self.last_io_dir:
            _mr = Path(self.last_io_dir) / "monthly_reports"
            _mr.mkdir(exist_ok=True)
            default_dir = str(_mr)
        path = filedialog.askopenfilename(
            title="Open monthly report",
            initialdir=default_dir,
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            try:
                if sys.platform == "win32":
                    os.startfile(path)
                else:
                    subprocess.Popen(["open", path])
                self._status(f"Opened: {Path(path).name}")
            except Exception as e:
                messagebox.showerror("Cannot open file", str(e))

    def monthly(self) -> None:
        if not require_io_dir():
            return
        if not self.template_path:
            messagebox.showerror("Missing template", "Select the Word template first."); return
        try:
            y, m = map(int, self.month_var.get().split("-"))
        except Exception:
            messagebox.showerror("Input error", "Month must be YYYY-MM."); return
        default = f"monthly_activity_report_{y}_{m:02d}.docx"
        # Default save dir: OneDrive monthly_reports or io_dir/monthly_reports
        _candidates_s = [
            Path.home() / "OneDrive - eumetsat.int" / "Exostaff Report Generator" / "monthly_reports",
            Path.home() / "OneDrive" / "Exostaff Report Generator" / "monthly_reports",
        ]
        if self.last_io_dir:
            _candidates_s.append(Path(self.last_io_dir) / "monthly_reports")
        _save_dir = next((str(c) for c in _candidates_s if c.exists()), None)
        if not _save_dir and self.last_io_dir:
            _mr = Path(self.last_io_dir) / "monthly_reports"
            _mr.mkdir(exist_ok=True)
            _save_dir = str(_mr)
        out = filedialog.asksaveasfilename(
            defaultextension=".docx", initialfile=default,
            filetypes=[("Word documents", "*.docx")],
            initialdir=_save_dir,
        )
        if not out:
            return

        # Show spinner IMMEDIATELY after file chosen — all heavy work runs inside it
        import threading
        dlg = GeneratingDialog(self.root)
        dlg.grab_set()
        self.root.update()   # force render before blocking work starts

        result = {"error": None, "skip": False}

        def _run():
            try:
                # Log state before generation
                log.warning("monthly: START y=%s m=%s io_root=%s weekly_dir=%s",
                            y, m, get_io_root(), weekly_dir())
                wf_list = list(weekly_dir().glob("week_*_CW*.json"))
                log.warning("monthly: found %d weekly files: %s",
                            len(wf_list), [f.name for f in wf_list])

                # Read archived data for the month
                monthly_data = load_monthly_data_from_weekly_reports(y, m)
                log.warning("monthly: loaded %d entries from archives", len(monthly_data))

                # Table (edited/deleted, not yet saved) wins over the archive
                monthly_data = merge_week(monthly_data, self._table_entries_in_month(y, m),
                                          self._loaded_week_keys)
                result["skipped"] = list(skipped_archives)

                log.warning("monthly: total %d entries after merge", len(monthly_data))

                if not monthly_data:
                    result["skip"] = True
                    return   # dialog closed in finally

                # Check user name consistency
                from docx import Document as _Doc
                _doc = _Doc(self.template_path)
                _template_users = [u for u, _ in get_user_tables(_doc)]
                _config_users   = load_users_from_config()
                _missing_tpl    = [u for u in _config_users if u not in _template_users]
                _missing_cfg    = [u for u in _template_users if u not in _config_users]
                _warn = []
                if _missing_tpl:
                    _warn.append(f"In config but missing in template:\n  {', '.join(_missing_tpl)}")
                if _missing_cfg:
                    _warn.append(f"In template but missing in config:\n  {', '.join(_missing_cfg)}")
                result["warn"] = _warn
                result["monthly_data"] = monthly_data

                result["unmapped"] = generate_monthly_docx(
                    self.template_path, monthly_data, y, m, out)
                log.warning("monthly: DONE — weekly files still present: %s",
                            [f.name for f in weekly_dir().glob("week_*_CW*.json")])
            except Exception as exc:
                log.error("monthly: EXCEPTION %s", exc)
                result["error"] = exc
            finally:
                self.root.after(0, dlg.close)

        t = threading.Thread(target=_run, daemon=True)
        t.start()
        self.root.wait_window(dlg)
        t.join(timeout=120)

        try:
            if result.get("error"):
                raise result["error"]

            if result.get("skip"):
                if not messagebox.askyesno(
                    "No monthly data",
                    "No saved weekly entries found for the selected month.\n\nGenerate an empty report anyway?",
                ):
                    return
                generate_monthly_docx(self.template_path, [], y, m, out)

            if result.get("skipped"):
                messagebox.showwarning(
                    "Unreadable weekly files",
                    "These weekly files could not be read and are NOT included in the report:\n\n"
                    + "\n".join(result["skipped"]),
                )

            if result.get("unmapped"):
                messagebox.showwarning(
                    "Entries not in report",
                    "These entries could not be matched to a task row of the template "
                    "and are NOT included in the report:\n\n"
                    + "\n".join(result["unmapped"][:20])
                    + "\n\nCheck task titles/numbers in config/tasks.txt against the template.",
                )

            if result.get("warn"):
                messagebox.showwarning(
                    "User name mismatch",
                    "Warning — user names differ between config and template:\n\n" +
                    "\n".join(result["warn"]) +
                    "\n\nThe report may have incomplete totals.",
                )

            self._status(f"Monthly report generated → {Path(out).name}")
            messagebox.showinfo("Done", f"Monthly Word report generated:\n{out}")
        except Exception as e:
            log.error("monthly: %s", e)
            messagebox.showerror("Generation error", str(e))

    # -----------------------------------------------------------------------
    # Help
    # -----------------------------------------------------------------------

    def show_writing_guidelines(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Activity Writing Guidelines")
        win.resizable(True, True)

        hdr = tk.Frame(win, bg=BG_HEADER, height=46)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="◈  Activity Writing Guidelines",
                 font=("Segoe UI Semibold", 11, "bold"),
                 bg=BG_HEADER, fg="#ffffff").pack(side="left", padx=14, pady=8)
        tk.Frame(win, bg=ACCENT, height=2).pack(fill="x")

        txt = scrolledtext.ScrolledText(
            win, font=("Segoe UI", 9), wrap="word",
            bg=BG_PANEL, fg=FG_MAIN, padx=20, pady=16,
            relief="flat", borderwidth=0,
        )
        txt.pack(fill="both", expand=True)

        GUIDE = """
ACTIVITY DESCRIPTION GUIDELINES
══════════════════════════════════════════════════════════════

The quality of the monthly Word report depends directly on how
activities are described. These guidelines help keep the report
clean, professional, and free of duplicates.

──────────────────────────────────────────────────────────────
1. USE A TYPE PREFIX
──────────────────────────────────────────────────────────────
Start the description with a short prefix that identifies the
nature of the work. This makes the report scannable at a glance
and helps group related activities.

  DOC:   Document drafting, review, or update
         e.g. "DOC: Updated GSRD VCD with new VM references"

  MTG:   Meeting, teleconference, or workshop attendance
         e.g. "MTG: EPS-Sterna weekly coordination call"

  REV:   Review of a document, design, or test result
         e.g. "REV: Review of MME DAS data package"

  TEST:  Test preparation, execution, or reporting
         e.g. "TEST: S-Band RFCT test specification update"

  PLAN:  Planning, scheduling, or roadmap activities
         e.g. "PLAN: IV&V campaign planning and sequencing"

  TOOLS: Tool development, configuration, or scripting
         e.g. "TOOLS: DOORS module maintenance and update"

  COORD: Coordination with external teams or contractors
         e.g. "COORD: Follow-up with ESA on open RIDs"

──────────────────────────────────────────────────────────────
2. KEEP DESCRIPTIONS GENERAL — NOT DAY-BY-DAY
──────────────────────────────────────────────────────────────
If an activity spans multiple days, use the SAME description
every day. The report will automatically aggregate the effort
and show one entry — not a repetition of identical lines.

  ✓ CORRECT (same text, 3 days):
    "DOC: GSRD VCD update — VM traceability section"
    "DOC: GSRD VCD update — VM traceability section"
    "DOC: GSRD VCD update — VM traceability section"
    → Report shows: 3 PD | DOC: GSRD VCD update…

  ✗ AVOID (minor variations that create duplicates):
    "DOC: Updated VCD chapter 4"
    "DOC: Continued VCD updates"
    "DOC: Finalised VCD chapter 4 modifications"
    → Report shows: 3 separate lines — cluttered and redundant

──────────────────────────────────────────────────────────────
3. BE SPECIFIC ENOUGH TO BE MEANINGFUL
──────────────────────────────────────────────────────────────
The description will appear in the official monthly report sent
to EUMETSAT. It must communicate what was done without requiring
context knowledge.

  ✓ Good: "DOC: Updated GSRD §4.3 Timeliness section for ESA data transfer"
  ✗ Too vague: "Worked on document"
  ✗ Too detailed: "Updated paragraph 4.3.2, line 17, changed 'shall' to 'will',
    reviewed with Bruno, sent for review, awaiting feedback from ESA team"

Aim for one sentence that captures the objective, not the steps.

──────────────────────────────────────────────────────────────
4. NAME THE DOCUMENT OR ARTEFACT
──────────────────────────────────────────────────────────────
Always name the document, system, or artefact you are working
on. Avoid generic references.

  ✓ "DOC: MSRD VCD — added verification references for System RoD"
  ✓ "REV: EPS-Sterna IV&V Test Procedures — EARS acquisition section"
  ✗ "Updated the document"
  ✗ "Reviewed the report"

──────────────────────────────────────────────────────────────
5. ONE DESCRIPTION PER TASK PER DAY
──────────────────────────────────────────────────────────────
Each entry is for one task on one day. If you worked on two
different tasks in the same day, create two separate entries
with effort split between them (e.g. 0.5 PD each).

Do not combine multiple tasks into a single description:
  ✗ "Updated VCD and attended weekly meeting and reviewed test spec"

Instead:
  ✓ Task 2 | 0.5 PD | "DOC: GSRD VCD — VM traceability update"
  ✓ Task 16 | 0.3 PD | "MTG: EPS-Sterna weekly coordination call"
  ✓ Task 7 | 0.2 PD | "REV: S-Band RFCT test specification"

──────────────────────────────────────────────────────────────
6. CONSISTENCY ACROSS TEAM MEMBERS
──────────────────────────────────────────────────────────────
When two people work on the same activity, agree on a common
description. This avoids the report showing the same work twice
with different wording.

  Agreed: "DOC: EPS-Sterna IV&V Test Procedure — EARS acquisition"
  Fabio logs:  Task 8 | 2 PD | (above)
  Bruno logs:  Task 8 | 1 PD | (above)
  → Report shows one clean entry with combined effort.

──────────────────────────────────────────────────────────────
7. QUICK REFERENCE — PREFIX CHEAT SHEET
──────────────────────────────────────────────────────────────

  DOC    Drafting / updating / finalising a document
  REV    Reviewing a document or artefact
  MTG    Meeting / call / workshop
  TEST   Test preparation, execution, or reporting
  PLAN   Planning / scheduling
  TOOLS  Tool or script development / maintenance
  COORD  Coordination / follow-up with external parties
  ANALY  Analysis or study
  PRES   Presentation preparation or delivery
"""
        txt.insert("1.0", GUIDE.strip())
        txt.configure(state="disabled")

        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        w, h = 700, min(640, sh - 120)
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{max(0,(sh-80-h)//2)}")

    def show_user_guide(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("User Guide")
        win.resizable(True, True)

        # Header
        hdr = tk.Frame(win, bg=BG_HEADER, height=46)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="◈  User Guide", font=("Segoe UI Semibold", 11, "bold"),
                 bg=BG_HEADER, fg="#ffffff").pack(side="left", padx=14, pady=8)
        tk.Frame(win, bg=ACCENT, height=2).pack(fill="x")

        # Scrollable text
        txt = scrolledtext.ScrolledText(
            win, font=("Segoe UI", 9), wrap="word",
            bg=BG_PANEL, fg=FG_MAIN, padx=20, pady=16,
            relief="flat", borderwidth=0,
        )
        txt.pack(fill="both", expand=True, padx=0, pady=0)

        GUIDE = """
ACTIVITY REPORT TOOL — USER GUIDE
══════════════════════════════════════════════════════════════

OVERVIEW
--------
This tool helps you log daily work activities, generate weekly
summaries, and produce the monthly Word activity report for the
EPS-Sterna and EPS-Aeolus System Engineering Service.

──────────────────────────────────────────────────────────────
1. FIRST-TIME SETUP
──────────────────────────────────────────────────────────────
Place the Word template file named:
  "EPS-Sterna and EPS-Aeolus System Engineering Service Template.docx"
in the same folder as the executable.

The tool loads it automatically on startup. The header bar
shows "✔  Template loaded" when ready. You can also load it
manually via "📂 Select Word template".

Set the Working folder (📁 button, top-left) to point
to the shared network folder where weekly archives are stored.
This setting is remembered across sessions.

──────────────────────────────────────────────────────────────
2. LOGGING ACTIVITIES (daily use)
──────────────────────────────────────────────────────────────
Fill in the form at the top of the window:

  • Date        — DD/MM/YY format. Click 📅 for a calendar picker.
                  Field turns red if the format is invalid.
  • User        — your name (remembered from last session)
  • Type        — Activity / Holiday / Sick Leave
  • Effort (PD) — person-days worked (e.g. 0.5 or 1.0)
                  Warning shown if daily total exceeds 1.0 PD.
  • Task        — select from the list in config/tasks.txt
  • Description — brief description of what was done

Click "💾 Save entry" (Ctrl+S) to add the entry to the table.

To edit: double-click a row, make changes, Save again.
To duplicate: right-click → Duplicate. The entry is copied into the
form as a NEW entry; adjust it and click Save entry to add it.
To delete: select a row and press Delete.
To undo: press Ctrl+Z (up to 20 levels).

──────────────────────────────────────────────────────────────
3. WEEKLY SUMMARY
──────────────────────────────────────────────────────────────
At the end of each week:

  1. (Optional) Open Edit → Next-week plan to enter planned
     activities for the coming week. The plan is included
     automatically in the generated summary.
  2. Set the Year and CW (calendar week) in the bottom bar.
  3. Click "📋 Generate weekly summary".
  4. Copy to clipboard and paste into your email/Teams message.

The current week's data is automatically archived to
weekly_reports/ when a new week begins (detected at startup).

To correct entries in a past week:
  1. Set the Year and CW of the week to fix in the bottom bar.
  2. Click "📂 Load week" — archived entries load into the table.
  3. Make your edits (add, modify, delete entries).
  4. Click "💾 Save week" — changes are merged back into the
     archive without overwriting other users' entries.

Note: if you accidentally add an entry from a different week
while editing, Save week will detect it and route each entry
to the correct archive file automatically.

──────────────────────────────────────────────────────────────
4. MONTHLY WORD REPORT
──────────────────────────────────────────────────────────────
At the end of each month:

  1. Make sure all team members have saved their weekly data
     to the shared folder (Scenario A) or exchanged JSON files
     (Scenario B — see section 5).
  2. Set the Month field (YYYY-MM) in the bottom bar.
  3. Click "📄 Generate monthly Word".
  4. Choose where to save the output .docx file.

The tool reads all weekly archives for that month from the
Working folder and fills the Word template automatically.
You can preview totals first via "📊 Statistics".

──────────────────────────────────────────────────────────────
5. TWO USAGE SCENARIOS
──────────────────────────────────────────────────────────────

SCENARIO A — Shared network folder (recommended for teams):
  • The executable and all data live in one shared folder.
  • Only one user can have the tool open at a time (lock file).
  • Set 📁 Working folder to the shared folder itself.
  • All users read/write the same weekly_reports/ archive.
  • No manual data exchange needed.
  • Open the tool via launch.bat (sets correct working directory).

SCENARIO B — Local copies + manual sync:
  • Each user has a local copy of the tool and their own data.
  • No lock conflicts — everyone works independently.
  • To consolidate before the monthly report:
      Export: "⬆ Export" → JSON → send to the report owner.
      Import: "⬇ Import" → select received JSON file.
      Duplicate entries are skipped automatically.
  • CSV/XLSX exports are for analysis in Excel only.

──────────────────────────────────────────────────────────────
6. IMPORT/EXPORT FOLDER
──────────────────────────────────────────────────────────────
The 📁 Working folder setting controls where:
  • Weekly archives (weekly_reports/) are read and written
  • Export and Import file dialogs open by default
  • The monthly Word report reads its source data from

Set it once to your shared folder and it is remembered.
If not set, all data is stored in the tool's local folder.

──────────────────────────────────────────────────────────────
7. KEYBOARD SHORTCUTS & MOUSE
──────────────────────────────────────────────────────────────
  Ctrl+S        Save current entry
  Ctrl+Z        Undo last action
  Delete        Delete selected row
  F5            Refresh table
  Double-click  Edit selected row
  Right-click   Context menu (Edit / Duplicate / Delete)
  📅 button     Open date picker calendar

──────────────────────────────────────────────────────────────
8. MULTI-USER LOCK SYSTEM
──────────────────────────────────────────────────────────────
When running from a shared folder, only one user can have the
tool open at a time. If another user has it open, you will see
a warning with their name, PC and when they were last seen.

A running tool refreshes its lock every 5 minutes. A lock that
has not been refreshed for 15 minutes (left after a crash) is
cleared automatically.

Use "🔓 Force Unlock" only if you are certain the other user
is no longer running the tool. If they still are, their tool
detects it within 5 minutes and switches to READ-ONLY: they can
still view data and generate reports, but nothing they do is
written to the shared folder.

──────────────────────────────────────────────────────────────
9. DATA FILES & BACKUPS
──────────────────────────────────────────────────────────────
  activities_current.json   — current unsaved week entries
  activities_current.state.json — entries loaded in the table (so
                              deletions reach the archive even
                              without Save week)
  app_state.json            — UI preferences (user, folders)
  next_week_plan.json       — planned activities for next week
  activity_report.lock      — multi-user lock (auto-managed)
  weekly_reports/           — one JSON file per calendar week
  backups/                  — auto-backup every 10 saves (last 20 kept)
  activity_report.log       — errors and warnings

All files are stored in the Working folder (if set)
or in the tool's local folder. View the log via Help → View log.

──────────────────────────────────────────────────────────────
10. RESET
──────────────────────────────────────────────────────────────
File → Reset clears the current table. Use it only to start
fresh. Entries already saved in the weekly archives are NOT
deleted and still count in the monthly report.

──────────────────────────────────────────────────────────────
11. VERSION & AUTO-UPDATE
──────────────────────────────────────────────────────────────
The current version is shown in the top-right corner of the
header bar (e.g. "v1.1.0").

The tool checks for updates at startup and every 4 hours while
open. It compares its own version with VERSION.txt in the
release folder (the shared OneDrive folder).

If a newer version is available:
  • A popup appears: "New version available vX.X.X"
  • Click "⬇ Update Now" to download automatically
  • The tool downloads the new exe, shows you the exact path
    where it was saved, then closes
  • Reopen the application from that path — it will now run
    the new version

If no Working folder is set yet:
  • Set the Working folder with 📁 — the version check runs
    immediately after, no restart needed

Note: the version shown in the header is the version of the
exe you are currently running. After updating, the header
will show the new version number on next launch.
"""
        txt.insert("1.0", GUIDE.strip())
        txt.configure(state="disabled")

        # Size and centre
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        w, h = 680, min(620, sh - 120)
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{max(0,(sh-80-h)//2)}")

    def _open_text_editor(self, filepath: Path, title: str) -> None:
        """Open a simple text editor for a config file."""
        win = tk.Toplevel(self.root)
        win.title(title)
        win.resizable(True, True)
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        w, h = 700, min(500, sh - 120)
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{max(0,(sh-80-h)//2)}")

        # Header
        hdr = tk.Frame(win, bg=BG_HEADER, height=40)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"◈  {title}", font=("Segoe UI Semibold", 10, "bold"),
                 bg=BG_HEADER, fg="#ffffff").pack(side="left", padx=14, pady=8)
        tk.Frame(win, bg=ACCENT, height=2).pack(fill="x")

        txt = scrolledtext.ScrolledText(win, font=("Consolas", 9), wrap="none",
                                         bg=BG_ENTRY, fg=FG_MAIN, padx=10, pady=10)
        txt.pack(fill="both", expand=True)
        if filepath.exists():
            txt.insert("1.0", read_config_text(filepath))

        btn_frame = tk.Frame(win, bg=BG_PANEL)
        btn_frame.pack(fill="x", pady=6)

        def save_and_reload():
            if read_only:
                messagebox.showwarning("Read-only", READ_ONLY_MSG, parent=win)
                return
            try:
                filepath.write_text(txt.get("1.0", "end-1c"), encoding="utf-8")   # Tk adds a trailing newline
                self.users, self.tasks = reload_config()
                reload_holidays()
                self._update_combos()
                self._status(f"Config saved and reloaded — {len(self.users)} users, {len(self.tasks)} tasks")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Save error", str(e), parent=win)

        ttk.Button(btn_frame, text="💾  Save & Reload", command=save_and_reload,
                   style="Accent.TButton").pack(side="left", padx=12)
        ttk.Button(btn_frame, text="✖  Cancel", command=win.destroy).pack(side="left", padx=4)

    def _update_combos(self) -> None:
        """Refresh user and task dropdowns from current self.users/self.tasks."""
        if hasattr(self, "user_cb"):
            self.user_cb["values"] = self.users
            if self.users and self.user_var.get() not in self.users:
                self.user_var.set(
                    self.last_user if self.last_user in self.users else self.users[0]
                )
        if hasattr(self, "wp_cb"):
            self._populate_wp_cb()

    def reload_config_from_files(self) -> None:
        """Reload users, tasks and holidays from config files and refresh UI."""
        if not require_io_dir():
            return
        try:
            self.users, self.tasks = reload_config()
            reload_holidays()
            # Update user combobox
            if hasattr(self, "user_cb"):
                self.user_cb["values"] = self.users
                cur = self.user_var.get()
                if self.users and cur not in self.users:
                    self.user_var.set(self.last_user if self.last_user in self.users else self.users[0])
            # Update WP and task comboboxes
            if hasattr(self, "wp_cb"):
                self._populate_wp_cb()
            self._status(
                f"Config reloaded — {len(self.users)} users, "
                f"{len(self.tasks)} tasks, {len(_holidays)} holidays"
            )
            messagebox.showinfo("Config reloaded",
                f"Users: {len(self.users)}\n"
                f"Tasks: {len(self.tasks)}\n"
                f"Holidays: {len(_holidays)}\n\n"
                "Dropdowns updated.")
        except Exception as e:
            messagebox.showerror("Config error", str(e))

    def open_users_config(self) -> None:
        if not require_io_dir():
            return
        create_default_config()
        self._open_text_editor(users_file(), "Edit Users — config/users.txt")

    def open_tasks_config(self) -> None:
        if not require_io_dir():
            return
        create_default_config()
        self._open_text_editor(tasks_file(), "Edit Tasks — config/tasks.txt")

    def open_holidays_config(self) -> None:
        if not require_io_dir():
            return
        create_default_config()
        self._open_text_editor(holidays_file(), "Edit Holidays — config/holidays.txt")

    def view_log(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Application Log")
        win.geometry(f"900x{min(500, self.root.winfo_screenheight()-120)}")
        txt = scrolledtext.ScrolledText(win, font=("Consolas", 9), wrap="none")
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        log_path = log_file()
        if log_path.exists():
            txt.insert("1.0", log_path.read_text(encoding="utf-8", errors="replace"))
        else:
            txt.insert("1.0", "(Log file not found — no warnings or errors recorded yet.)")
        txt.configure(state="disabled")
        ttk.Button(win, text="Refresh",
                   command=lambda: [
                       txt.configure(state="normal"),
                       txt.delete("1.0", "end"),
                       txt.insert("1.0", log_file().read_text(encoding="utf-8", errors="replace") if log_file().exists() else "(empty)"),
                       txt.configure(state="disabled"),
                   ]).pack(pady=(0, 8))

    def _periodic_update_check(self) -> None:
        latest, available = check_version()
        if available and latest != getattr(self, "_notified_version", None):
            self._show_update_notice(latest, APP_VERSION)
        self.root.after(UPDATE_CHECK_MS, self._periodic_update_check)

    def _show_update_notice(self, latest: str, current: str) -> None:
        """Show update dialog with auto-update button."""
        self._notified_version = latest
        import threading
        win = tk.Toplevel(self.root)
        win.title("Update available")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.configure(bg=BG_PANEL)
        w, h = 460, 240
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{max(0,(sh-80-h)//2)}")

        tk.Frame(win, bg=ACCENT, height=3).pack(fill="x")
        tk.Label(win, text="🔔  New version available",
                 font=("Segoe UI", 11, "bold"), bg=BG_PANEL, fg=ACCENT).pack(pady=(16, 4))
        tk.Label(win,
                 text=f"Your version:   {current}\nLatest version: {latest}",
                 font=("Consolas", 10), bg=BG_PANEL, fg=FG_MAIN, justify="left").pack(pady=(0, 6))

        status_var = tk.StringVar(value="Click 'Update Now' to download it. The application then "
                                        "closes — reopen it to use the new version.")
        tk.Label(win, textvariable=status_var,
                 font=("Segoe UI", 8), bg=BG_PANEL, fg=FG_DIM,
                 wraplength=400, justify="center").pack(pady=(0, 8))

        btn_frame = tk.Frame(win, bg=BG_PANEL)
        btn_frame.pack()

        def do_update():
            btn_update.configure(state="disabled")
            btn_later.configure(state="disabled")
            status_var.set("Downloading… please wait.")
            win.update()

            def _run():
                try:
                    ok, msg = auto_update(
                        on_progress=lambda m: self.root.after(0, lambda: status_var.set(m))
                    )
                except Exception as exc:   # never leave the dialog stuck
                    log.error("auto_update: %s", exc)
                    ok, msg = False, str(exc)
                if ok and msg.startswith("ok_close:"):
                    exe_path = msg.split(":", 1)[1]
                    exe_dir  = str(Path(exe_path).parent)
                    def _finish(p=exe_path, d=exe_dir):
                        messagebox.showinfo(
                            "Update ready",
                            f"New version downloaded successfully!\n\n"
                            f"Saved to:\n{p}\n\n"
                            f"The application will now close.\n"
                            f"Reopen it from:\n{d}",
                            parent=win,
                        )
                        self._shutdown()
                    self.root.after(0, _finish)
                elif not ok:
                    self.root.after(0, lambda: [
                        status_var.set(f"Update failed:\n{msg}"),
                        btn_later.configure(state="normal"),
                    ])

            threading.Thread(target=_run, daemon=True).start()

        btn_update = ttk.Button(btn_frame, text="⬇  Update Now",
                                command=do_update, style="Accent.TButton", width=16)
        btn_update.pack(side="left", padx=(0, 8))

        btn_later = ttk.Button(btn_frame, text="Later",
                               command=win.destroy, width=10)
        btn_later.pack(side="left")

        tk.Frame(win, bg=ACCENT, height=3).pack(fill="x", side="bottom")

    def _about(self) -> None:
        messagebox.showinfo(
            "About",
            f"{APP_TITLE}  v{APP_VERSION}\n\n"
            "A tool for logging daily activities and generating\n"
            "weekly summaries and monthly Word activity reports\n"
            "for the EPS-Sterna and EPS-Aeolus programmes.\n\n"
            "EUMETSAT — System Engineering Service",
        )


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app  = App(root)
    root.mainloop()