from __future__ import annotations

import json
import os
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "docs" / "community.json"


def write_output(name: str, value: str) -> None:
    output = os.environ.get("GITHUB_OUTPUT")
    if output:
        with open(output, "a", encoding="utf-8") as handle:
            handle.write(f"{name}={value}\n")


data = json.loads(CONFIG.read_text(encoding="utf-8"))
status = str(data.get("status", "pending")).lower()
valid_until = data.get("valid_until")
needs_reminder = status != "active" or not valid_until
reason = "二维码尚未配置"

if status == "active" and valid_until:
    expiry = date.fromisoformat(valid_until)
    remaining = (expiry - date.today()).days
    needs_reminder = remaining <= 2
    reason = f"距离失效还有 {remaining} 天" if remaining >= 0 else f"已过期 {-remaining} 天"

write_output("needs_reminder", str(needs_reminder).lower())
write_output("reason", reason)
print(json.dumps({"needs_reminder": needs_reminder, "reason": reason}, ensure_ascii=False))
