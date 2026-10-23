"""Regressions found in the final data-path review: identical copies and
multi-week writes interrupted by an error."""

from datetime import date, timedelta
from conftest import archives

MON = date(2026, 10, 19)
NEXT_MON = date(2026, 10, 26)


def test_identical_entry_lost_on_second_save_week(session, tool):
    s = session(MON)
    s.add(MON, "Meetings"); s.save_week(); s.close()
    s = session(NEXT_MON)
    s.add(MON, "Meetings")                 # second identical half-day, past week
    s.save_week()
    assert archives(tool)["week_2026_CW43"] == ["Meetings"] * 2
    s.save_week()                          # nothing changed
    assert archives(tool)["week_2026_CW43"] == ["Meetings"] * 2


def test_identical_entry_lost_after_weekly_summary(session, tool, monkeypatch):
    s = session(MON)
    s.add(MON, "Meetings"); s.save_week(); s.close()
    s = session(NEXT_MON)
    s.add(MON, "Meetings")
    s.weekly(2026, 43)                     # archives the table
    got = []
    monkeypatch.setattr(tool.StatsWindow, "__init__",
                        lambda self, p, data, y, m, **k: got.append(data))
    s.app.month_var.set("2026-10"); s.app.show_stats()
    assert len(got[0]) == 2                # archive has 2, stats must see 2


def test_partial_multiweek_save_duplicates(session, tool):
    s = session(MON)
    s.add(MON, "A"); s.add(NEXT_MON, "B")
    bad = tool.share / "weekly_reports" / "week_2026_CW44.json"
    bad.parent.mkdir(exist_ok=True); bad.write_text("{garbage", encoding="utf-8")
    s.save_week()                          # CW43 written, CW44 raises
    bad.unlink()
    s.save_week()                          # retry after fixing the file
    assert archives(tool)["week_2026_CW43"] == ["A"]


def test_identical_entry_lost_after_restart(session, tool):
    s = session(MON)
    s.add(MON, "Meetings"); s.save_week(); s.close()
    s = session(NEXT_MON)
    s.add(MON, "Meetings"); s.save_week(); s.close()
    s = session(NEXT_MON + timedelta(days=1))
    s.save_week()
    assert archives(tool)["week_2026_CW43"] == ["Meetings"] * 2


def test_partial_auto_archive_duplicates(session, tool):
    s = session(MON)
    s.add(MON - timedelta(days=7), "B")   # CW42
    s.add(MON, "A")                       # CW43
    s.close()
    bad = tool.share / "weekly_reports" / "week_2026_CW43.json"
    bad.parent.mkdir(exist_ok=True); bad.write_text("{garbage", encoding="utf-8")
    s = session(NEXT_MON); s.close()      # CW42 archived, CW43 raises
    bad.unlink()
    s = session(NEXT_MON); s.close()      # retry
    assert archives(tool)["week_2026_CW42"] == ["B"]
