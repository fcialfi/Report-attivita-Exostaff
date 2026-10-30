"""Real App, driven like a user over several days and sessions. Invariant
everywhere: an entry saved and not deliberately deleted is archived exactly
once; a deleted entry never comes back. Needs a display (xvfb-run)."""

import json
from datetime import date, timedelta

from conftest import archives

MON = date(2026, 10, 19)          # CW43
NEXT_MON = date(2026, 10, 26)     # CW44


def week_of_entries(session, n=5):
    for i in range(n):
        d = MON + timedelta(days=i)
        s = session(d); s.add(d, f"d{i}"); s.close()


def test_daily_sessions_then_week_change_archives(session, tool):
    week_of_entries(session)
    s = session(NEXT_MON)
    assert archives(tool) == {"week_2026_CW43": ["d0", "d1", "d2", "d3", "d4"]}
    assert s.table() == []


def test_forgotten_entry_never_erases_an_archived_one(session, tool):
    week_of_entries(session)
    s = session(NEXT_MON)
    s.add(MON + timedelta(days=4), "fri-forgot")          # same day/task as d4
    s.save_week(); s.save_week()                           # repeated saves
    s.weekly(2026, 43)
    s.edit("fri-forgot", effort=0.25); s.save_week()
    s.close()
    s = session(NEXT_MON + timedelta(days=1)); s.save_week(); s.close()
    assert archives(tool)["week_2026_CW43"] == ["d0", "d1", "d2", "d3", "d4", "fri-forgot"]


def test_forgotten_entry_kept_without_save_week(session, tool):
    week_of_entries(session)
    s = session(NEXT_MON); s.add(MON + timedelta(days=4), "fri-forgot"); s.close()
    s = session(NEXT_MON + timedelta(days=1)); s.close()
    assert archives(tool)["week_2026_CW43"] == ["d0", "d1", "d2", "d3", "d4", "fri-forgot"]


def test_edit_and_delete_in_a_loaded_past_week(session, tool):
    week_of_entries(session)
    s = session(NEXT_MON)
    s.load_week(2026, 43)
    s.edit("d1", new_desc="d1-edited")
    s.delete("d2")
    s.close()                                              # no Save week
    s = session(NEXT_MON + timedelta(days=1)); s.close()   # applied at next start
    assert archives(tool)["week_2026_CW43"] == ["d0", "d1-edited", "d3", "d4"]


def test_current_week_after_edit_and_delete(session, tool):
    s = session(MON)
    s.add(MON, "A"); s.add(MON, "C"); s.save_week()
    s.edit("A", new_desc="B"); s.delete("C")
    s.app.load_current_week(); s.save_week()
    assert archives(tool)["week_2026_CW43"] == ["B"]


def test_two_identical_half_days_are_both_kept(session, tool):
    s = session(MON)
    s.add(MON, "Meetings"); s.add(MON, "Meetings"); s.save_week(); s.close()
    s = session(NEXT_MON)
    s.add(MON, "Meetings")                                 # a third, in the past week
    s.save_week()
    assert archives(tool)["week_2026_CW43"] == ["Meetings"] * 3


def test_same_entry_in_two_wps_is_two_entries(session, tool):
    s = session(MON)
    s.add(MON, "m", task=1, wp="WP6"); s.save_week(); s.close()
    s = session(NEXT_MON)
    s.add(MON, "m", task=1, wp="SLM"); s.save_week()
    wps = sorted(a["task_wp"] for a in tool.load_week_entries(2026, 43))
    assert wps == ["SLM", "WP6"]


def test_year_boundary(session, tool):
    for i in range(5):                                     # 28/12/2026 .. 01/01/2027
        d = date(2026, 12, 28) + timedelta(days=i)
        s = session(d); s.add(d, f"y{i}"); s.close()
    s = session(date(2027, 1, 4)); s.add(date(2027, 1, 4), "w1"); s.save_week()
    a = archives(tool)
    assert a["week_2026_CW53"] == ["y0", "y1", "y2", "y3", "y4"]
    assert a["week_2027_CW01"] == ["w1"]
    dec = tool.load_monthly_data_from_weekly_reports(2026, 12)
    jan = tool.load_monthly_data_from_weekly_reports(2027, 1)
    assert sorted(x["description"] for x in dec) == ["y0", "y1", "y2", "y3"]
    assert sorted(x["description"] for x in jan) == ["w1", "y4"]


def test_import_twice_adds_nothing_and_delete_keeps_archive(session, tool, tmp_path, monkeypatch):
    s = session(MON, user="Bob B")
    s.add(MON, "bob1", user="Bob B"); s.add(MON, "bob1", user="Bob B"); s.save_week()
    exported = tmp_path / "bob.json"
    exported.write_text(json.dumps(s.app.data), encoding="utf-8")
    s.close()
    s = session(NEXT_MON)
    monkeypatch.setattr(tool.filedialog, "askopenfilename", lambda **k: str(exported))
    s.app.import_table()                                   # already archived: skipped
    assert s.table() == []
    s.save_week()
    assert archives(tool)["week_2026_CW43"] == ["bob1", "bob1"]


def test_undo_restores_a_deleted_entry_but_not_after_save(session, tool):
    s = session(MON)
    s.add(MON, "a"); s.add(MON, "b"); s.save_week()
    s.delete("b"); s.app.undo()
    assert s.table() == ["a", "b"]
    s.delete("b"); s.save_week(); s.app.undo()             # nothing to undo after Save week
    assert archives(tool)["week_2026_CW43"] == ["a"]


def test_loading_a_week_during_an_edit_does_not_duplicate(session, tool):
    s = session(MON)
    s.add(MON, "orig"); s.save_week()
    s._select("orig"); s.app.load_selected_for_edit()
    s.app.desc.delete("1.0", "end"); s.app.desc.insert("1.0", "changed")
    s.app.load_current_week()                              # asks, then discards the edit
    s.add(MON, "new")
    assert s.table() == ["new", "orig"]


def test_weekly_summary_ignores_deleted_entries(session, tool, monkeypatch):
    s = session(MON)
    s.add(MON, "a"); s.save_week(); s.delete("a")
    captured = {}
    real = tool.generate_weekly_text
    monkeypatch.setattr(tool, "generate_weekly_text",
                        lambda data, *a, **k: captured.setdefault("d", data) and real(data, *a, **k))
    s.weekly(2026, 43)
    assert captured["d"] == [] and archives(tool)["week_2026_CW43"] == []


def test_read_only_writes_nothing(session, tool, monkeypatch):
    s = session(MON)
    s.add(MON, "a"); s.save_week()
    before = {p.name: p.read_bytes() for p in tool.share.rglob("*") if p.is_file()}
    monkeypatch.setattr(tool, "read_only", True)
    s.add(MON, "b"); s.save_week(); s.app.undo()
    after = {p.name: p.read_bytes() for p in tool.share.rglob("*") if p.is_file()}
    assert after == before


def test_statistics_see_unsaved_changes(session, tool, monkeypatch):
    s = session(MON)
    s.add(MON, "a", effort=1); s.add(MON + timedelta(days=1), "b", effort=1); s.save_week()
    s.edit("a", effort=0.5); s.delete("b"); s.add(MON, "c", effort=0.25)
    got = []
    monkeypatch.setattr(tool.StatsWindow, "__init__", lambda self, p, data, y, m, **k: got.append(data))
    s.app.month_var.set("2026-10"); s.app.show_stats()
    assert sorted((x["description"], x["effort"]) for x in got[0]) == [("a", 0.5), ("c", 0.25)]


def test_report_parameters_dialog_saves(session, tool):
    s = session(date(2026, 10, 19))
    s.app.open_report_params()
    dlg = next(w for w in s.root.winfo_children() if isinstance(w, tool.ReportParamsDialog))
    dlg._vars["EUM Technical Officer"].set("  Mario Rossi ")
    dlg._save()
    assert tool.load_report_params() == {"EUM Technical Officer": "Mario Rossi"}
