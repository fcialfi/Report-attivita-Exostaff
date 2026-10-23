"""Model-based randomized test of the data path of the real App.

Random sequences of user actions across days and sessions, checked against
a model of which entries must exist. On a failure the sequence is shrunk to
the fewest steps that still fail and printed.

CI runs a short sample; for a long run:
    MBT_SEEDS=300 MBT_STEPS=60 xvfb-run -a pytest tests/test_model_random.py
Env:  MBT_SEEDS (default 20), MBT_STEPS (default 40), MBT_SEED0 (default 0)
"""
import importlib.util
import json
import os
import random
import sys
import tempfile
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
_spec = importlib.util.spec_from_file_location("repo_conftest", REPO / "tests" / "conftest.py")
C = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(C)

import Exostaff_Activity_Report_Tool as T  # noqa: E402

START = date(2026, 10, 19)
TASKS = [(1, "WP6"), (2, "WP6"), (3, "WP7"), (1, "SLM")]
EFF = [0.25, 0.5, 0.75, 1.0]


class Failure(AssertionError):
    pass


# ---------------------------------------------------------------- generator
def gen(rng, n):
    acts, jumped = [], False
    for _ in range(n):
        r = rng.random()
        if r < 0.30:
            acts.append(("add", rng.randint(0, 2), rng.randint(0, 6), rng.randrange(4),
                         rng.choice(EFF)))
        elif r < 0.42:
            acts.append(("edit", rng.random(), rng.random() < 0.6, rng.choice([None] + EFF),
                         rng.choice([None, None, (rng.randint(0, 2), rng.randint(0, 6))])))
        elif r < 0.50:
            acts.append(("delete", rng.random()))
        elif r < 0.58:
            acts.append(("load", rng.randint(1, 4)))
        elif r < 0.63:
            acts.append(("current",))
        elif r < 0.71:
            acts.append(("save",))
        elif r < 0.76:
            acts.append(("weekly", rng.randint(0, 3)))
        elif r < 0.82:
            acts.append(("stats", rng.random()))
        elif r < 0.97:
            acts.append(("reopen", rng.randint(0, 3), rng.random() < 0.2))
        elif not jumped:
            jumped = True
            acts.append(("yearjump",))
        else:
            acts.append(("save",))
    return acts


# ---------------------------------------------------------------- runner
def ids_of(app):
    return [a["description"] for a in app.data]


def run(actions, workdir):
    mp = pytest.MonkeyPatch()
    try:
        return _run(actions, Path(workdir), mp)
    finally:
        mp.undo()
        T.set_io_dir(str(Path(workdir) / "unused"))


def _setup(tmp, mp):
    io = tmp / "share"
    (io / "config").mkdir(parents=True)
    (io / "config" / "users.txt").write_text(C.USERS, encoding="utf-8")
    (io / "config" / "tasks.txt").write_text(C.TASKS, encoding="utf-8")

    class FakeDate(date):
        @classmethod
        def today(cls):
            v = C.Today.value
            return cls(v.year, v.month, v.day)
    mp.setattr(T, "date", FakeDate)
    mp.setattr(T, "read_only", False)
    mp.setattr(T, "skipped_archives", [])
    mp.setattr(T, "APP_STATE_FILE", tmp / "app_state.json")
    mp.setattr(T, "check_version", lambda: (T.APP_VERSION, False))
    mp.setattr(T, "load_app_state", lambda: (None, [], [], T._test_user, str(io)))
    mp.setattr(T, "auto_backup", lambda *a, **k: None)   # speed only
    dialogs, answers = [], {"yesno": True}

    def record(kind):
        def f(*a, **k):
            dialogs.append((kind,) + tuple(a[:2]))
            return answers["yesno"] if kind == "askyesno" else None
        return f
    for kind in ("showinfo", "showwarning", "showerror", "askyesno"):
        mp.setattr(T.messagebox, kind, record(kind))
    captured = []
    mp.setattr(T.StatsWindow, "__init__",
               lambda self, p, data, y, m, **k: captured.append(data))
    T.set_io_dir(str(io)); T.reload_config()
    T.share = io; T.dialogs = dialogs; T.answers = answers; T.Today = C.Today
    return captured, dialogs


def _run(actions, tmp, mp):
    captured, dialogs = _setup(tmp, mp)
    oracle = {}          # id -> (date iso, user, task, wp, effort)
    counter = [0]
    day = START
    user = "Alice A"
    log = []
    s = C.Session(T, day, user)

    def fail(msg):
        raise Failure(msg + "\nlog:\n  " + "\n  ".join(log))

    def rel_date(wk_back, wd):
        mon = day - timedelta(days=day.weekday()) - timedelta(weeks=wk_back)
        return mon + timedelta(days=wd)

    def check_table():
        ids = ids_of(s.app)
        if len(ids) != len(set(ids)):
            fail(f"duplicate ids in table: {sorted(k for k, v in Counter(ids).items() if v > 1)}")
        for a in s.app.data:
            i = a["description"]
            if i not in oracle:
                fail(f"table holds dead id {i}")
            o = oracle[i]
            got = (a["date"], a["user"], a["task_no"], a["task_wp"], float(a["effort"]))
            if got != o:
                fail(f"table entry {i}: {got} != oracle {o}")

    def pick(frac):
        ids = sorted(ids_of(s.app))
        return ids[int(frac * len(ids))] if ids else None

    def ensure_saved():
        if s.app._table_dirty:
            s.save_week(); log.append("  (auto Save week: table dirty)")

    def reopen(new_day, new_user):
        nonlocal s, day, user
        s.close()
        day, user = new_day, new_user
        s = C.Session(T, day, user)

    for act in actions:
        k = act[0]
        if k == "add":
            _, wb, wd, ti, eff = act
            d = rel_date(wb, wd)
            task, wp = TASKS[ti]
            counter[0] += 1
            i = f"id{counter[0]:03d}"
            log.append(f"[{day}] add {i} {d} T{task}/{wp} {eff} {user}")
            s.add(d, i, task=task, wp=wp, effort=eff, user=user)
            oracle[i] = (d.isoformat(), user, task, wp, eff)
        elif k == "edit":
            _, frac, newid, eff, dd = act
            tgt = pick(frac)
            if tgt is None:
                continue
            d = rel_date(*dd) if dd else None
            nid = None
            if newid:
                counter[0] += 1; nid = f"id{counter[0]:03d}"
            log.append(f"[{day}] edit {tgt} -> desc={nid} effort={eff} date={d}")
            s.edit(tgt, new_desc=nid, effort=eff, d=d)
            o = oracle.pop(tgt)
            o = ((d.isoformat() if d else o[0]), o[1], o[2], o[3], (eff if eff is not None else o[4]))
            oracle[nid or tgt] = o
        elif k == "delete":
            tgt = pick(act[1])
            if tgt is None:
                continue
            log.append(f"[{day}] delete {tgt}")
            s.delete(tgt); del oracle[tgt]
        elif k == "load":
            mon = day - timedelta(days=day.weekday()) - timedelta(weeks=act[1])
            y, cw = T.iso_year_week_for_date(mon)
            ensure_saved()
            log.append(f"[{day}] Load week {y}-CW{cw}")
            s.load_week(y, cw)
        elif k == "current":
            ensure_saved()
            log.append(f"[{day}] Current week")
            s.app.load_current_week()
        elif k == "save":
            log.append(f"[{day}] Save week")
            s.save_week()
        elif k == "weekly":
            mon = day - timedelta(days=day.weekday()) - timedelta(weeks=act[1])
            y, cw = T.iso_year_week_for_date(mon)
            log.append(f"[{day}] Weekly {y}-CW{cw}")
            s.weekly(y, cw)
        elif k == "stats":
            dates = sorted({o[0] for o in oracle.values()}) or [day.isoformat()]
            dm = date.fromisoformat(dates[int(act[1] * len(dates))])
            y, m = dm.year, dm.month
            captured.clear()
            s.app.month_var.set(f"{y}-{m:02d}"); s.app.show_stats()
            got = Counter((a["description"], a["date"], float(a["effort"])) for a in captured[0])
            exp = Counter((i, o[0], o[4]) for i, o in oracle.items()
                          if o[0][:7] == f"{y}-{m:02d}")
            log.append(f"[{day}] Stats {y}-{m:02d}")
            if got != exp:
                fail(f"stats {y}-{m:02d}: extra={sorted((got - exp).elements())} "
                     f"missing={sorted((exp - got).elements())}")
        elif k == "reopen":
            _, adv, bob = act
            nu = "Bob B" if bob else "Alice A"
            log.append(f"[{day}] close; reopen +{adv}d as {nu}")
            reopen(day + timedelta(days=adv), nu)
        elif k == "yearjump":
            if day < date(2026, 12, 28):
                log.append(f"[{day}] close; reopen 2026-12-28")
                reopen(date(2026, 12, 28), user)
                counter[0] += 1; i = f"id{counter[0]:03d}"
                log.append(f"[{day}] add {i} 2026-12-31 T1/WP6 0.5 {user}")
                s.add(date(2026, 12, 31), i, user=user); oracle[i] = ("2026-12-31", user, 1, "WP6", 0.5)
                log.append(f"[{day}] close; reopen 2027-01-04")
                reopen(date(2027, 1, 4), user)
        errs = [d for d in dialogs if d[0] == "showerror"]
        if errs:
            fail(f"error dialog: {errs}")
        check_table()

    # final: Save week + close + reopen a week later
    log.append(f"[{day}] FINAL Save week; close; reopen +8d")
    s.save_week()
    reopen(day + timedelta(days=8), "Alice A")
    check_table()
    s.close()
    arch = []
    for p in sorted((T.share / "weekly_reports").glob("week_*.json")):
        arch += json.loads(p.read_text(encoding="utf-8"))
    cnt = Counter(a["description"] for a in arch)
    dup = sorted(i for i, c in cnt.items() if c > 1)
    missing = sorted(set(oracle) - set(cnt))
    dead = sorted(set(cnt) - set(oracle))
    wrong = sorted(f"{a['description']}: {(a['date'], a['effort'])} vs {(oracle[a['description']][0], oracle[a['description']][4])}"
                   for a in arch if a["description"] in oracle
                   and (a["date"], float(a["effort"]), a["user"], a["task_no"], a["task_wp"])
                   != (oracle[a["description"]][0], oracle[a["description"]][4],
                       oracle[a["description"]][1], oracle[a["description"]][2], oracle[a["description"]][3]))
    if dup or missing or dead or wrong:
        fail(f"archive: dup={dup} missing={missing} dead={dead} wrong={wrong}")
    return len(actions)


def shrink(actions):
    def fails(acts):
        with tempfile.TemporaryDirectory() as d:
            try:
                run(acts, d); return None
            except Failure as e:
                return str(e)
    cur = list(actions); msg = fails(cur)
    changed = True
    while changed:
        changed = False
        for i in range(len(cur)):
            cand = cur[:i] + cur[i + 1:]
            m = fails(cand)
            if m:
                cur, msg, changed = cand, m, True
                break
    return cur, msg


N = int(os.environ.get("MBT_SEEDS", 20)); STEPS = int(os.environ.get("MBT_STEPS", 40))
SEED0 = int(os.environ.get("MBT_SEED0", 0))


@pytest.mark.parametrize("seed", range(SEED0, SEED0 + N))
def test_random_sequence(seed, tmp_path):
    try:
        tk = T.tk.Tk(); tk.destroy()
    except Exception as exc:
        pytest.skip(f"no display ({exc})")
    acts = gen(random.Random(seed), STEPS)
    try:
        run(acts, tmp_path)
    except Failure:
        small, msg = shrink(acts)
        pytest.fail(f"seed {seed} minimal sequence ({len(small)} actions): {small}\n{msg}")
