"""Data model: normalisation, archive merging, table state, config parsing,
file handling. No display needed."""

import json
import os
from datetime import date

import pytest

from conftest import entry


# ── dates ───────────────────────────────────────────────────────────────────

def test_iso_weeks_across_the_year_boundary(tool):
    assert tool.iso_year_week_for_date(date(2026, 12, 31)) == (2026, 53)
    assert tool.iso_year_week_for_date(date(2027, 1, 3)) == (2026, 53)
    assert tool.iso_year_week_for_date(date(2027, 1, 4)) == (2027, 1)
    assert tool.week_after(2026, 53) == (2027, 1)
    assert tool.cw_range(2026, 53) == (date(2026, 12, 28), date(2027, 1, 1))   # Mon-Fri


@pytest.mark.parametrize("text", ["23/10/26", "23/10/2026", "2026-10-23", "23.10.2026"])
def test_safe_date_formats(tool, text):
    assert tool.safe_date(text) == date(2026, 10, 23)


# ── normalisation ───────────────────────────────────────────────────────────

def test_legacy_entry_fields(tool):
    a = tool.normalise_activity({"date": "12/10/2026", "user": "alice  a", "desc": "d",
                                 "task_no": 1, "task_title": "T1", "effort": "0,5", "task_wp": "wp06"})
    assert a["date"] == "2026-10-12"
    assert a["user"] == "Alice A"          # users.txt spelling
    assert a["description"] == "d"
    assert a["effort"] == 0.5
    assert a["task_wp"] == "WP6"


def test_missing_effort_is_one_day_null_is_zero(tool):
    assert tool.normalise_activity({"date": "2026-10-12", "user": "x"})["effort"] == 1
    assert tool.normalise_activity({"date": "2026-10-12", "user": "x", "effort": None})["effort"] == 0


def test_control_characters_removed(tool):
    a = tool.normalise_activity(entry("2026-10-12", "a\x0bb\x01c"))
    assert a["description"] == "a\nbc"


# ── archive merging ─────────────────────────────────────────────────────────

def test_merge_keeps_entries_that_were_not_loaded(tool):
    archived = [entry("2026-10-23", "d4")]
    table = [entry("2026-10-23", "forgot")]          # same day, user, task
    merged = tool.merge_week(archived, table, set())
    assert sorted(a["description"] for a in merged) == ["d4", "forgot"]


def test_merge_replaces_and_deletes_loaded_entries(tool):
    a, b, c = entry("2026-10-20", "a"), entry("2026-10-21", "b"), entry("2026-10-22", "c")
    loaded = {tool.loaded_key(x) for x in (a, b, c)}
    edited_b = dict(b, description="b2")
    merged = tool.merge_week([a, b, c], [a, edited_b], loaded)    # c deleted, b edited
    assert sorted(x["description"] for x in merged) == ["a", "b2"]


def test_merge_keeps_identical_entry_in_another_wp(tool):
    wp6 = entry("2026-10-20", "m", wp="WP6", title="Mgmt")
    slm = entry("2026-10-20", "m", wp="SLM", title="Mgmt")
    assert len(tool.merge_week([wp6], [slm], set())) == 2


def test_old_state_keys_are_honoured(tool):
    a = entry("2026-10-20", "a", user="Alice A")
    # 5-field stable key written by an older version, old user spelling
    tool.save_table_state({("20/10/2026", "alice  a", "Activity", "WP6", "1")})
    loaded = tool.load_table_state()
    assert tool._is_loaded(a, loaded)
    # 7-field key of the previous version
    assert tool._is_loaded(a, {tool.activity_key(a)})


# ── weekly archives ─────────────────────────────────────────────────────────

def test_week_archive_round_trip_and_monthly_dedup(tool):
    e1, e2 = entry("2026-10-20", "a"), entry("2026-10-20", "a")   # 2 x 0.5 PD, same text
    tool.write_week_archive(2026, 43, [e1, e2])
    data = tool.load_monthly_data_from_weekly_reports(2026, 10)
    assert len(data) == 2 and sum(x["effort"] for x in data) == 1.0


def test_onedrive_conflict_copy_is_not_counted(tool):
    tool.write_week_archive(2026, 43, [entry("2026-10-20", "a", effort=1)])
    wd = tool.share / "weekly_reports"
    (wd / "week_2026_CW43-LAPTOP.json").write_text(
        json.dumps([entry("2026-10-20", "a", effort=0.5)]), encoding="utf-8")
    data = tool.load_monthly_data_from_weekly_reports(2026, 10)
    assert [x["effort"] for x in data] == [1.0]
    assert any("LAPTOP" in s for s in tool.skipped_archives)


def test_unreadable_archive_is_reported_not_overwritten(tool):
    wd = tool.share / "weekly_reports"; wd.mkdir(exist_ok=True)
    (wd / "week_2026_CW43.json").write_text("", encoding="utf-8")
    assert tool.load_week_entries(2026, 43) == []
    assert "week_2026_CW43.json" in tool.skipped_archives
    with pytest.raises(RuntimeError):
        tool.write_week_archive(2026, 43, [entry("2026-10-20", "x")])


def test_atomic_write_retries_while_the_file_is_locked(tool, monkeypatch):
    calls = {"n": 0}
    real = os.replace

    def flaky(src, dst):
        calls["n"] += 1
        if calls["n"] < 3:
            raise PermissionError(13, "locked by OneDrive")
        return real(src, dst)
    monkeypatch.setattr(tool.os, "replace", flaky)
    monkeypatch.setattr(tool.time, "sleep", lambda s: None)
    target = tool.share / "x.json"
    tool.atomic_write_json(target, {"a": 1})
    assert json.loads(target.read_text()) == {"a": 1}
    assert not (tool.share / "x.json.tmp").exists()


def test_atomic_write_refuses_in_read_only(tool, monkeypatch):
    monkeypatch.setattr(tool, "read_only", True)
    with pytest.raises(PermissionError):
        tool.atomic_write_json(tool.share / "y.json", {})


# ── auto-archive at week change ─────────────────────────────────────────────

def test_week_change_archives_and_applies_pending_deletions(tool):
    keep, gone = entry("2026-10-20", "keep"), entry("2026-10-21", "gone")
    tool.write_week_archive(2026, 43, [keep, gone])
    cur = entry("2026-10-27", "cur-deleted")
    tool.write_week_archive(2026, 44, [cur])
    tool.write_json([keep])                                  # gone and cur deleted in table
    tool.save_table_state({tool.loaded_key(x) for x in (keep, gone, cur)})
    tool.Today.value = date(2026, 10, 28)
    data, reset, _ = tool.read_json()
    assert reset and data == []
    assert [a["description"] for a in tool.load_week_entries(2026, 43)] == ["keep"]
    assert tool.load_week_entries(2026, 44) == []


# ── config ──────────────────────────────────────────────────────────────────

def test_users_same_person_twice_keeps_first_spelling(tool):
    (tool.share / "config" / "users.txt").write_text("Fabio CIALFI\nfabio  cialfi\nBob B\n", encoding="utf-8")
    users, _ = tool.reload_config()
    assert users == ["Fabio CIALFI", "Bob B"]


def test_tasks_wp_spellings_and_no_wp(tool):
    (tool.share / "config" / "tasks.txt").write_text(
        "1 | A | P | wp6 | X\n2 | B | P | WP06 | X\n3 | C | P | WP-7 | X\n4 | D | Common | X\n",
        encoding="utf-8")
    _, tasks = tool.reload_config()
    assert sorted({wp for _, _, wp in tasks}) == ["", "WP6", "WP7"]


def test_config_files_with_bom_and_ansi(tool):
    (tool.share / "config" / "users.txt").write_bytes("﻿Nicolò R\n".encode("utf-8"))
    assert tool.load_users_from_config() == ["Nicolò R"]
    (tool.share / "config" / "users.txt").write_bytes("Nicolò R\n".encode("cp1252"))
    assert tool.load_users_from_config() == ["Nicolò R"]


def test_holiday_formats(tool):
    (tool.share / "config" / "holidays.txt").write_text(
        "2026-01-01 New Year\n01/05/2026 | Labour\n2026-12-25\tXmas\n2026-12-24 | Eve\nnonsense\n",
        encoding="utf-8")
    h = tool.load_holidays()
    assert h == {date(2026, 1, 1): "New Year", date(2026, 5, 1): "Labour",
                 date(2026, 12, 25): "Xmas", date(2026, 12, 24): "Eve"}


# ── import / export ─────────────────────────────────────────────────────────

def test_own_csv_export_reimports(tool, tmp_path):
    import csv
    p = tmp_path / "x.csv"
    row = entry("2026-10-20", 'a, b, c, d, e, f, g, h, i\nline2; "q"')
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date", "user", "entry_type", "task_wp", "task_no",
                                          "task_title", "effort", "description"], delimiter=";")
        w.writeheader(); w.writerow(row)
    ok, rejected = tool.import_activity_table(p)
    assert rejected == [] and ok[0]["description"] == row["description"]


def test_csv_saved_by_excel_in_ansi(tool, tmp_path):
    p = tmp_path / "x.csv"
    p.write_bytes("date;user;entry_type;task_no;task_title;effort;description\n"
                  "20/10/2026;Alice A;Activity;1;T1;0,5;caffè\n".encode("cp1252"))
    ok, rejected = tool.import_activity_table(p)
    assert rejected == [] and ok[0]["description"] == "caffè" and ok[0]["effort"] == 0.5


def test_xlsx_export_text_not_formula_and_total_activities_only(tool, tmp_path):
    import openpyxl
    out = tmp_path / "x.xlsx"
    tool.export_to_xlsx([entry("2026-10-20", "=> follow-up", effort=1),
                         entry("2026-10-21", "", etype="Holiday", effort=1)], str(out))
    ws = openpyxl.load_workbook(out).active
    assert ws.cell(2, 7).data_type == "s"
    assert ws.cell(4, 6).value == 1   # TOTAL row: leave excluded


# ── version check ───────────────────────────────────────────────────────────

def test_version_parsing(tool):
    assert tool._parse_version("1.2.10") > tool._parse_version("1.2.9")


def test_version_file_with_bom(T, tool, monkeypatch, tmp_path):
    rel = tmp_path / "rel"; rel.mkdir()
    (rel / "VERSION.txt").write_bytes("﻿99.0.0\r\n".encode("utf-8"))
    monkeypatch.setattr(T, "RELEASE_DIR", str(rel))
    monkeypatch.undo()   # check_version was stubbed by the fixture: use the real one
    monkeypatch.setattr(T, "RELEASE_DIR", str(rel))
    import importlib
    real = importlib.import_module("Exostaff_Activity_Report_Tool").__dict__["check_version"]
    latest, update = real()
    assert latest == "99.0.0" and update


def test_description_suggestions(tool):
    from conftest import entry
    T = tool
    data = [entry("2026-10-01", "GSTRACE update"), entry("2026-10-02", "GSTRACE update."),
            entry("2026-10-03", "gstrace  update", user="Bob B"), entry("2026-10-05", "Slides for the CDR"),
            entry("2026-10-06", "Review of RFCT documents"),
            entry("2026-10-07", "Other task", task=2), entry("2026-10-08", "Same task, other WP", wp="SLM"),
            entry("2026-10-09", "Holiday", etype="Holiday")]
    # one suggestion per activity (case, spaces, final full stop), its most used spelling, most used first
    assert T.description_suggestions(data, 1, "WP6") == [
        "GSTRACE update", "Review of RFCT documents", "Slides for the CDR"]
    # filtered on every word typed, anywhere in the text, any case
    assert T.description_suggestions(data, 1, "WP6", "cdr SLI") == ["Slides for the CDR"]
    assert T.description_suggestions(data, 1, "WP6", "gstrace update") == []   # already typed in full
    assert T.description_suggestions(data, 1, "SLM") == ["Same task, other WP"]
    assert T.description_suggestions(data, 3, "WP7") == []
    assert len(T.description_suggestions(
        [entry("2026-10-01", f"activity {k}") for k in range(20)], 1, "WP6")) == 8


def test_recent_archived_entries(tool):
    from conftest import entry
    T = tool
    T.write_week_archive(2026, 43, [entry("2026-10-19", "recent")])
    T.write_week_archive(2025, 10, [entry("2025-03-03", "too old")])
    assert [a["description"] for a in T.recent_archived_entries()] == ["recent"]


def test_update_exe_name_and_fingerprint(tool):
    from pathlib import Path
    T = tool
    exe = Path("C:/Tools/ActivityReportTool_1.2.2.exe")
    # a local copy named with an old version takes the published name
    assert T.updated_exe_path(exe, "ActivityReportTool.exe") == Path("C:/Tools/ActivityReportTool.exe")
    plain = Path("C:/Tools/My tool.exe")
    assert T.updated_exe_path(plain, "ActivityReportTool.exe") == plain   # any other name stays
    tmp = exe.with_name("ActivityReportTool_1.2.2_new.exe")
    new = Path("C:/Tools/ActivityReportTool.exe")
    bat = T.update_launcher_script(tmp, exe, new)
    assert 'move /y "' + str(tmp) + '" "' + str(new) + '"' in bat
    assert 'del /f "' + str(exe) + '"' in bat and "goto retry" in bat   # old name removed, with retries
    same = T.update_launcher_script(tmp, plain, plain)
    assert "del /f" not in same and 'move /y "' + str(tmp) + '" "' + str(plain) + '"' in same


def test_update_waits_for_the_new_exe(tool):
    T = tool
    rel = T.share
    (rel / "VERSION.txt").write_text("9.9.9", encoding="utf-8")
    (rel / "EXE_LOCATION.txt").write_text("C:\\Users\\Cialfi\\x\\ActivityReportTool.exe\n", encoding="utf-8")
    (rel / "ActivityReportTool.exe").write_bytes(b"old build")
    (rel / "EXE_SHA256.txt").write_text(T.file_sha256(rel / "ActivityReportTool.exe") + "\n", encoding="utf-8")
    exe, err = T.get_latest_exe_path()
    assert exe == rel / "ActivityReportTool.exe" and not err
    (rel / "EXE_SHA256.txt").write_text("0" * 64 + "\n", encoding="utf-8")   # VERSION synced, exe not yet
    exe, err = T.get_latest_exe_path()
    assert exe is None and "not the new version yet" in err
    (rel / "EXE_SHA256.txt").unlink()   # releases made before the fingerprint: as before
    assert T.get_latest_exe_path()[0] == rel / "ActivityReportTool.exe"


def test_report_month_choice(tool):
    from datetime import date
    T = tool
    assert T.default_report_month(date(2026, 11, 9)) == "2026-10"    # first days: last month's report
    assert T.default_report_month(date(2026, 11, 11)) == "2026-11"
    assert T.default_report_month(date(2027, 1, 5)) == "2026-12"     # across the year
    assert T.month_label_of("2026-10") == "October 2026" and T.month_from_label("October 2026") == "2026-10"
    labels = T.recent_month_labels(date(2026, 2, 20))
    assert labels[:3] == ["February 2026", "January 2026", "December 2025"] and len(labels) == 12


def test_build_does_not_modify_the_source(tmp_path):
    """build.bat (update_version.py) writes build_info.json, not the .py: git pull is never blocked."""
    import json, shutil, subprocess, sys
    from conftest import REPO
    for name in ("update_version.py", "Exostaff_Activity_Report_Tool.py"):
        shutil.copy(REPO / name, tmp_path / name)
    source = (tmp_path / "Exostaff_Activity_Report_Tool.py").read_bytes()
    run = lambda *a: subprocess.run([sys.executable, str(tmp_path / "update_version.py"), *a],
                                    capture_output=True, text=True, check=True).stdout
    default = run().strip()
    assert default.count(".") == 2                                    # the default of the source
    run("2.0.1", str(tmp_path / "share"))
    info = json.loads((tmp_path / "build_info.json").read_text(encoding="utf-8"))
    assert info["version"] == "2.0.1" and info["release_dir"].endswith("share")
    assert run().strip() == "2.0.1"                                   # next build proposes it
    assert (tmp_path / "Exostaff_Activity_Report_Tool.py").read_bytes() == source


def test_build_info_read_at_start(tmp_path):
    import json, subprocess, sys, shutil
    from conftest import REPO
    shutil.copy(REPO / "Exostaff_Activity_Report_Tool.py", tmp_path / "tool.py")
    (tmp_path / "build_info.json").write_text(json.dumps({"version": "3.4.5", "release_dir": "~/X"}))
    out = subprocess.run([sys.executable, "-c", "import tool; print(tool.APP_VERSION, tool.RELEASE_DIR)"],
                         cwd=tmp_path, capture_output=True, text=True)
    assert out.stdout.split()[-2:] == ["3.4.5", "~/X"], out.stderr


def test_build_proposes_the_latest_version(tmp_path):
    """The version proposed by build.bat: the highest of the last build here, the one published
    in the shared folder and the default of the source (not an old value of the source)."""
    import json, shutil, subprocess, sys
    from conftest import REPO
    for name in ("update_version.py", "Exostaff_Activity_Report_Tool.py"):
        shutil.copy(REPO / name, tmp_path / name)
    share = tmp_path / "share"; share.mkdir()
    (tmp_path / "build_config.txt").write_text(f"EXE_SHARE={share}\nDATA_SHARE={share}\n", encoding="utf-8")
    run = lambda *a: subprocess.run([sys.executable, str(tmp_path / "update_version.py"), *a],
                                    capture_output=True, text=True, check=True).stdout.strip()
    (share / "VERSION.txt").write_text("7.0.2\n", encoding="utf-8")
    assert run() == "7.0.2"                                  # published, no build here yet
    (tmp_path / "build_info.json").write_text(json.dumps({"version": "7.1.0"}))
    assert run() == "7.1.0"                                  # built here, not yet published
    (share / "VERSION.txt").write_text("not a version", encoding="utf-8")
    assert run() == "7.1.0"


def test_holidays_in_the_eumetsat_list_format(tool):
    from datetime import date
    T = tool
    (T.share / "config" / "holidays.txt").write_text(
        "# 2026\n2026-12-25 | Christmas Day\n"
        "# 2027\n"
        "Friday, 1 January \tNew Year’s Day\n"
        "\t\n"
        "Monday, 8 February\tRosenmontag\n"
        "Friday, 24 December pm (1/2 day)\tChristmas Eve\n"
        "Friday, 31 December   New Year’s Eve\n"
        "2027-05-06 | Ascension Day\n"
        "Someday\tnot a date\n", encoding="utf-8")
    T.reload_holidays()
    assert T._holidays[date(2026, 12, 25)] == "Christmas Day"
    assert T._holidays[date(2027, 1, 1)] == "New Year’s Day"
    assert T._holidays[date(2027, 2, 8)] == "Rosenmontag"
    assert T._holidays[date(2027, 12, 24)] == "Christmas Eve, pm (1/2 day)"
    assert T._holidays[date(2027, 12, 31)] == "New Year’s Eve"
    assert date(2027, 5, 6) in T._holidays and len(T._holidays) == 6
    assert T.holiday_lines_ignored == ["Someday\tnot a date"]
    # without a year line above, a date without year is not guessed
    (T.share / "config" / "holidays.txt").write_text("Monday, 8 February\tRosenmontag\n", encoding="utf-8")
    T.reload_holidays()
    assert T._holidays == {} and len(T.holiday_lines_ignored) == 1


OUTLOOK_CSV = (
    '﻿"Subject","Start Date","Start Time","End Date","End Time","All day event","Reminder on/off",'
    '"Reminder Date","Reminder Time","Meeting Organizer","Required Attendees","Optional Attendees",'
    '"Meeting Resources","Billing Information","Categories","Description","Location","Mileage","Priority",'
    '"Private","Sensitivity","Show time as"\n'
    '"Christmas Day","25/12/2026","00:00:00","26/12/2026","00:00:00","True","False",,,,,,,,"Holiday",,,,"Normal","False","Normal","3"\n'
    '"EPS-Sterna Ground Segment Engineering Team Weekly Meeting","10/11/2026","10:00:00","10/11/2026","11:00:00","False",,,,,,,,,,,,,"Normal","False","Normal","2"\n'
    '"Upd1: EPS Sterna CCB","12/11/2026","14:00:00","12/11/2026","15:30:00","False",,,,,,,,,,,,,"Normal","False","Normal","2"\n'
    '"Maintenance of M-Files","11/11/2026","08:00:00","11/11/2026","12:00:00","False",,,,,,,,,,,,,"Normal","False","Normal","2"\n'
    '"Optional talk","13/11/2026","12:00:00","13/11/2026","13:00:00","False",,,,,,,,,,,,,"Normal","False","Normal","1"\n'
    '"Next week","17/11/2026","09:00:00","17/11/2026","10:00:00","False",,,,,,,,,,,,,"Normal","False","Normal","2"\n')


def test_meetings_from_an_outlook_csv(tool, tmp_path):
    from datetime import datetime
    T = tool
    path = tmp_path / "calendar.csv"
    path.write_text(OUTLOOK_CSV, encoding="utf-8")
    week = T.read_outlook_csv(path, datetime(2026, 11, 9), datetime(2026, 11, 16))
    assert [m["subject"][:12] for m in week] == ["EPS-Sterna G", "Maintenance ", "Upd1: EPS St", "Optional tal"]
    words = T.parse_exclude_words(T.DEFAULT_MEETING_EXCLUDE)
    assert [T.meeting_preselected(m, words) for m in week] == [True, False, True, False]   # excluded, tentative
    assert T.clean_meeting_subject("Upd1: EPS Sterna CCB") == "EPS Sterna CCB"
    assert T.clean_meeting_subject("RE: FW:  Review") == "Review"
    assert [T.meeting_effort(m) for m in week] == [0.2, 0.5, 0.2, 0.2]   # hours / 8, rounded up to 0.1
    half = {"start": week[0]["start"], "end": week[0]["start"] + (week[0]["end"] - week[0]["start"]) / 2}
    assert T.meeting_effort(half) == 0.1                                  # 30 min
    task = T.meeting_task([(1, "T1", "WP6"), (22, "Participation in any necessary meeting", "WP6")])
    entries = T.meeting_entries([week[0], week[2]], "Alice A", task)
    assert [(e["date"], e["task_no"], e["effort"], e["description"]) for e in entries] == [
        ("2026-11-10", 22, 0.2, "EPS-Sterna Ground Segment Engineering Team Weekly Meeting"),
        ("2026-11-12", 22, 0.2, "EPS Sterna CCB")]


def test_meeting_error_points_to_the_new_outlook(tool):
    T = tool
    text = T.meeting_error_text(Exception("Outlook.Application.GetNamespace"), "outlook")
    assert "new Outlook" in text and "New Outlook" in text and "GetNamespace" in text
    assert T.meeting_error_text(ValueError("bad date"), "csv") == "The meetings could not be read:\nbad date"
