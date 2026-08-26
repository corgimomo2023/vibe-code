from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

PROJECT_ROOT = Path(os.environ.get("VIBE_CODE_ROOT", Path(__file__).resolve().parents[2]))
REPORTS = PROJECT_ROOT / "wiki/product/reports"
MANIFEST = REPORTS / "vibe-code-report-artifact-manifest.json"
ARTIFACTS = (
    "vibe-code-no-code-poc-prd-and-delivery-plan.docx",
    "vibe-code-poc-technical-design-and-runtime-architecture.docx",
    "vibe-code-poc-data-specification-and-schema.xlsx",
)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


missing = [name for name in ARTIFACTS if not (REPORTS / name).is_file()]
if missing:
    raise SystemExit(f"missing report artifacts: {missing}")

payload = {
    "schema": "vibe-code-report-artifact-manifest/v1",
    "artifacts": {name: sha256(REPORTS / name) for name in ARTIFACTS},
}
MANIFEST.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
print({"manifest": str(MANIFEST), "artifacts": payload["artifacts"]})
