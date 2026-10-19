"""Read build_config.txt and print EXE_SHARE, DATA_SHARE, LAST_VERSION on separate lines."""
import sys
import os

config_file = sys.argv[1] if len(sys.argv) > 1 else "build_config.txt"
cfg = {}
if os.path.exists(config_file):
    for line in open(config_file, encoding="utf-8"):
        line = line.strip()
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip()

print(cfg.get("EXE_SHARE", ""))
print(cfg.get("DATA_SHARE", ""))
print(cfg.get("LAST_VERSION", "1.0.0"))
