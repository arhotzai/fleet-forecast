"""
Extract ONLY the `Settings` mappings from a Lessors Update workbook into
assets/market_settings.json — the single artefact the Market Intelligence
section keeps. No fleet data, no IBA rows and no relationship tables are read.

    python tools/extract_market_settings.py "~/Desktop/Lessors Update 26.08.xlsx"

Settings layout (one pair of columns per mapping):
    A:B  manager name      -> lessor group
    D:E  aircraft model    -> family
    G:H  IBA status        -> Active / On Order / Other
    J:K  lessor group      -> funding / main shareholders
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))
from fleet_erp.market.core import settings_from_workbook  # noqa: E402

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_OUT = os.path.join(_ROOT, "assets", "market_settings.json")
_KEYS = ("managers", "models", "statuses", "funding")


def extract(path: str) -> dict:
    with open(os.path.expanduser(path), "rb") as f:
        cfg = settings_from_workbook(f.read())
    if cfg is None:
        raise SystemExit(f"No usable 'Settings' sheet in {path}")
    return {k: v for k, v in cfg.items() if not k.startswith("_")}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    cfg = extract(sys.argv[1])
    with open(_OUT, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=1, ensure_ascii=False)
    print(f"wrote {_OUT}")
    for k in _KEYS:
        print(f"  {k}: {len(cfg[k])}")
    print("  families:", ", ".join(cfg["familyOrder"]))
