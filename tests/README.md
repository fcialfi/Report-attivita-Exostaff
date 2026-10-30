# Tests

Automated checks of the data path (entries → table → weekly archives →
reports), run by GitHub Actions on every pull request and on main.

```bash
pip install -r requirements-dev.txt
pytest                      # Windows / a desktop: the Tk tests open hidden windows
xvfb-run -a pytest          # headless Linux
```

- `test_core.py` — normalisation, archive merging, table state, config files,
  import/export, version check.
- `test_workflows.py` — the real App driven over several days and sessions
  (week change, forgotten entries, past-week edits, import, undo, year
  boundary, read-only). Invariant: a saved entry is archived exactly once and
  a deleted one never comes back.
- `test_reports.py` — monthly Word report on the real template, weekly
  summary, task trend, next-week plan.
- `test_lock.py` — the single-user lock between two PCs.
- `test_regressions.py` — identical copies of an entry, multi-week writes
  interrupted by an error, a state file older than the table (OneDrive)
  and the clean-up of an archive with duplicated entries.
- `test_model_random.py` — random sequences of user actions over days and
  sessions checked against a model of which entries must exist (20 short
  sequences in CI; `MBT_SEEDS=300 MBT_STEPS=60` for a long run).

Every bug fixed in the tool gets a test here, so it cannot come back.
