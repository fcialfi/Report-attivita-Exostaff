"""Helper script called by build.bat to read/set the version and the release
folder of the build. They go to build_info.json (not in git), which build.bat
bundles into the exe: the source file is never modified, so `git pull` is
never blocked by a build.

    python update_version.py                       -> print current version
    python update_version.py 1.2.0 [release_dir]   -> set version (+ release folder)
"""
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = HERE / "Exostaff_Activity_Report_Tool.py"
BUILD_INFO = HERE / "build_info.json"


def home_relative(path: str) -> str:
    """C:\\Users\\<me>\\OneDrive - x\\Folder -> ~/OneDrive - x/Folder, so the
    path resolves on every user's PC (each has their own home folder)."""
    p = Path(path)
    try:
        return "~/" + p.resolve().relative_to(Path.home().resolve()).as_posix()
    except ValueError:
        return p.as_posix()   # e.g. a network drive N:\... — same for everybody


def read_info() -> dict:
    try:
        return json.loads(BUILD_INFO.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def current_version() -> str:
    """The version of the last build on this PC, else the default in the source."""
    version = read_info().get("version")
    if version:
        return version
    m = re.search(r'APP_VERSION\s*=\s*_BUILD\.get\("version",\s*"([^"]+)"\)', SCRIPT.read_text(encoding="utf-8"))
    return m.group(1) if m else "1.0.0"


def main(argv) -> int:
    if not argv:
        print(current_version())
        return 0
    new_ver = argv[0].strip()
    if not re.fullmatch(r"\d+\.\d+\.\d+", new_ver):
        print(f"[ERROR] Version must look like 1.2.0 (got {new_ver!r})")
        return 1
    info = read_info()
    info["version"] = new_ver
    if len(argv) > 1:
        info["release_dir"] = home_relative(argv[1])
        print(f"[OK] Release folder: {info['release_dir']}")
    BUILD_INFO.write_text(json.dumps(info, indent=2) + "\n", encoding="utf-8")
    print(f"[OK] Version updated to {new_ver}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
