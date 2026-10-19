"""Helper script called by build.bat to read/update APP_VERSION in the main script."""
import re
import sys

SCRIPT = "Exostaff_Activity_Report_Tool.py"

if len(sys.argv) == 1:
    # Read current version
    with open(SCRIPT, "r", encoding="utf-8") as f:
        s = f.read()
    m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', s)
    print(m.group(1) if m else "1.0.0")

elif len(sys.argv) == 2:
    # Write new version
    new_ver = sys.argv[1]
    with open(SCRIPT, "r", encoding="utf-8") as f:
        s = f.read()
    s = re.sub(r'APP_VERSION\s*=\s*"[^"]*"',
               f'APP_VERSION       = "{new_ver}"', s)
    with open(SCRIPT, "w", encoding="utf-8") as f:
        f.write(s)
    print(f"[OK] Version updated to {new_ver}")
