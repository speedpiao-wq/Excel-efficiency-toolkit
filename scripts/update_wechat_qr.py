from __future__ import annotations

import argparse
import json
import shutil
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "docs" / "community.json"
TARGET = ROOT / "docs" / "assets" / "wechat-group-current.png"


parser = argparse.ArgumentParser(description="Update the current WeChat group QR used by the stable community page.")
parser.add_argument("image", type=Path, help="Path to the new QR screenshot in PNG format")
parser.add_argument("--valid-until", required=True, help="Expected expiry date in YYYY-MM-DD format")
parser.add_argument("--dry-run", action="store_true", help="Validate and show the change without writing files")
args = parser.parse_args()

source = args.image.resolve()
if not source.is_file():
    parser.error(f"QR image does not exist: {source}")
if source.suffix.lower() != ".png":
    parser.error("QR image must be a PNG file")

expiry = date.fromisoformat(args.valid_until)
if expiry < date.today():
    parser.error("valid-until cannot be in the past")

config = json.loads(CONFIG.read_text(encoding="utf-8"))
config.update(
    {
        "status": "active",
        "qr_image": "assets/wechat-group-current.png",
        "valid_until": expiry.isoformat(),
        "updated_at": date.today().isoformat(),
        "message": "扫描二维码加入微信群；如二维码失效，请使用 GitHub 备用入口。",
    }
)

if args.dry_run:
    print(json.dumps({"target": str(TARGET), "config": config}, ensure_ascii=False, indent=2))
else:
    shutil.copy2(source, TARGET)
    CONFIG.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Updated QR image: {TARGET}")
    print(f"Expected expiry: {expiry.isoformat()}")
