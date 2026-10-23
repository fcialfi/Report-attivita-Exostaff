"""Shared fixtures. The tool is imported once; every test gets a fresh
working folder with its own config, a fixed "today", and message boxes that
record instead of blocking. UI tests (the `app` fixture) need a display:
run them with `xvfb-run -a pytest` on a headless machine."""

import importlib
import sys
from datetime import date
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

USERS = "Alice A\nBob B\n"
TASKS = ("1 | T1 | P | WP6 | AlA BoB\n"
         "2 | T2 | P | WP6 | AlA BoB\n"
         "3 | T3 | P | WP7 | AlA BoB\n"
         "1 | S1 | P | SLM | AlA BoB\n")


@pytest.fixture(scope="session")
def T():
    return importlib.import_module("Exostaff_Activity_Report_Tool")


class Today:
    """Patchable date: date.today() returns `Today.value`."""
    value = date(2026, 10, 19)


@pytest.fixture
def tool(T, tmp_path, monkeypatch):
    io = tmp_path / "share"
    (io / "config").mkdir(parents=True)
    (io / "config" / "users.txt").write_text(USERS, encoding="utf-8")
    (io / "config" / "tasks.txt").write_text(TASKS, encoding="utf-8")

    class FakeDate(date):
        @classmethod
        def today(cls):
            v = Today.value
            return cls(v.year, v.month, v.day)

    Today.value = date(2026, 10, 19)
    monkeypatch.setattr(T, "date", FakeDate)
    monkeypatch.setattr(T, "read_only", False)
    monkeypatch.setattr(T, "skipped_archives", [])
    monkeypatch.setattr(T, "APP_STATE_FILE", tmp_path / "app_state.json")
    monkeypatch.setattr(T, "check_version", lambda: (T.APP_VERSION, False))
    T._test_user = "Alice A"
    monkeypatch.setattr(T, "load_app_state", lambda: (None, [], [], T._test_user, str(io)))

    dialogs = []
    answers = {"yesno": True}

    def record(kind):
        def f(*a, **k):
            dialogs.append((kind,) + tuple(a[:2]))
            return answers["yesno"] if kind == "askyesno" else None
        return f
    for kind in ("showinfo", "showwarning", "showerror", "askyesno"):
        monkeypatch.setattr(T.messagebox, kind, record(kind))

    T.set_io_dir(str(io))
    T.reload_config()
    T.share = io
    T.dialogs = dialogs
    T.answers = answers
    T.Today = Today
    yield T
    T.set_io_dir(str(tmp_path / "unused"))   # don't leave the next test our folder


def entry(d, desc, task=1, wp="WP6", effort=0.5, user="Alice A", etype="Activity", title=None):
    return {"date": d if isinstance(d, str) else d.isoformat(), "user": user, "entry_type": etype,
            "task_no": task if etype == "Activity" else 0,
            "task_title": title if title is not None else (f"S{task}" if wp == "SLM" else f"T{task}"),
            "task_wp": wp if etype == "Activity" else "", "effort": effort, "description": desc}


class Session:
    """Drives the real App the way a user does."""

    def __init__(self, T, day, user="Alice A"):
        T.Today.value = day
        T._test_user = user
        self.T = T
        self.root = T.tk.Tk()
        self.app = T.App(self.root)

    # -- actions ---------------------------------------------------------
    def add(self, d, desc, task=1, wp="WP6", effort=0.5, user=None, etype="Activity"):
        a = self.app
        if a.edit_index is not None:
            a.cancel_edit(clear_desc=True)
        a.date_var.set(d.strftime("%d/%m/%y"))
        if user:
            a.user_var.set(user)
        a.entry_type_var.set(etype)
        if etype == "Activity":
            a.wp_var.set(wp); a._on_wp_changed()
            a.task_var.set(f"{task} - {'S' if wp == 'SLM' else 'T'}{task}")
        a.effort_var.set(str(effort))
        a.desc.delete("1.0", "end"); a.desc.insert("1.0", desc)
        a.save_activity()

    def _select(self, desc):
        i = next(i for i, x in enumerate(self.app.data) if x["description"] == desc)
        iid = next(k for k, v in self.app._display_index_map.items() if v == i)
        self.app.tree.selection_set(iid)

    def delete(self, desc):
        self._select(desc); self.app.delete_selected()

    def edit(self, desc, new_desc=None, effort=None, d=None):
        self._select(desc); self.app.load_selected_for_edit()
        if new_desc is not None:
            self.app.desc.delete("1.0", "end"); self.app.desc.insert("1.0", new_desc)
        if effort is not None:
            self.app.effort_var.set(str(effort))
        if d is not None:
            self.app.date_var.set(d.strftime("%d/%m/%y"))
        self.app.save_activity()

    def load_week(self, y, cw):
        self.app.week_year.set(str(y)); self.app.cw_var.set(str(cw))
        self.app.load_week_into_table()

    def save_week(self):
        self.app.save_current_week_button()

    def weekly(self, y, cw):
        self.app.week_year.set(str(y)); self.app.cw_var.set(str(cw))
        self.app.weekly()

    def close(self):
        self.app._shutdown()

    # -- observations ----------------------------------------------------
    def table(self):
        return sorted(x["description"] for x in self.app.data)


def archives(T):
    """{week file stem: sorted descriptions}"""
    import json
    out = {}
    for p in sorted((T.share / "weekly_reports").glob("week_*.json")):
        out[p.stem] = sorted(a["description"] for a in json.loads(p.read_text(encoding="utf-8")))
    return out


@pytest.fixture
def session(tool):
    """session(day, user=...) -> Session; closed automatically."""
    try:
        import tkinter as tk
        tk.Tk().destroy()
    except Exception as exc:   # no display
        pytest.skip(f"no display for Tk ({exc}); run with xvfb-run")
    opened = []

    def open_(day, user="Alice A"):
        s = Session(tool, day, user)
        opened.append(s)
        return s
    yield open_
    for s in opened:
        try:
            s.root.destroy()
        except Exception:
            pass
