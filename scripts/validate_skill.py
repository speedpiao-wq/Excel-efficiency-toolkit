from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skill" / "excel-global-addin-maker"


def fail(message: str) -> None:
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


skill_md = SKILL / "SKILL.md"
if not skill_md.is_file():
    fail("SKILL.md is missing")

text = skill_md.read_text(encoding="utf-8")
frontmatter = re.match(r"^---\s*\n(.*?)\n---\s*\n", text, re.DOTALL)
if not frontmatter:
    fail("SKILL.md has no valid YAML frontmatter")
metadata = frontmatter.group(1)
if not re.search(r"(?m)^name:\s*excel-global-addin-maker\s*$", metadata):
    fail("Skill name is missing or incorrect")
if not re.search(r"(?m)^description:\s*\S.+$", metadata):
    fail("Skill description is missing")

required = [
    SKILL / "agents" / "openai.yaml",
    SKILL / "scripts" / "build_textbox_formatter.ps1",
    SKILL / "references" / "excel-addin-safety.md",
    SKILL / "references" / "beginner-delivery.md",
    SKILL / "assets" / "icon.svg",
]
for path in required:
    if not path.is_file():
        fail(f"Required skill file is missing: {path.relative_to(ROOT)}")

openai_yaml = (SKILL / "agents" / "openai.yaml").read_text(encoding="utf-8")
if "$excel-global-addin-maker" not in openai_yaml:
    fail("openai.yaml default prompt does not mention the skill")

for html in (ROOT / "docs").glob("*.html"):
    content = html.read_text(encoding="utf-8")
    if "file:///" in content or re.search(r"[A-Za-z]:\\", content):
        fail(f"Public page contains a local absolute path: {html.relative_to(ROOT)}")

print("Repository Skill and public documentation validation passed.")
