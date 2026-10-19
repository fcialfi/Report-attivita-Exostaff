"""Helper script called by build.bat to read/update APP_VERSION (and RELEASE_DIR)
in the main script.

    python update_version.py                       -> print current version
    python update_version.py 1.2.0 [release_dir]   -> set version (+ release folder)
"""
import re
import sys
from pathlib import Path

SCRIPT = "Exostaff_Activity_Report_Tool.py"


def home_relative(path: str) -> str:
    """C:\\Users\\<me>\\OneDrive - x\\Folder -> ~/OneDrive - x/Folder, so the
    path resolves on every user's PC (each has their own home folder)."""
    p = Path(path)
    try:
        return "~/" + p.resolve().relative_to(Path.home().resolve()).as_posix()
    except ValueError:
        return p.as_posix()   # e.g. a network drive N:\... — same for everybody


with open(SCRIPT, "r", encoding="utf-8") as f:
    s = f.read()

if len(sys.argv) == 1:
    m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', s)
    print(m.group(1) if m else "1.0.0")
    sys.exit(0)

new_ver = sys.argv[1].strip()
if not re.fullmatch(r"\d+\.\d+\.\d+", new_ver):
    print(f"[ERROR] Version must look like 1.2.0 (got {new_ver!r})")
    sys.exit(1)
s = re.sub(r'APP_VERSION\s*=\s*"[^"]*"', f'APP_VERSION       = "{new_ver}"', s, count=1)

if len(sys.argv) > 2:
    rel = home_relative(sys.argv[2])
    s = re.sub(r'RELEASE_DIR\s*=\s*"[^"]*"', lambda _m: f'RELEASE_DIR       = "{rel}"', s, count=1)
    print(f"[OK] Release folder: {rel}")

with open(SCRIPT, "w", encoding="utf-8") as f:
    f.write(s)
print(f"[OK] Version updated to {new_ver}")
