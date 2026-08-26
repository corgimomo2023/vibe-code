#!/usr/bin/env python3
"""Recalculate the generated XLSX with LibreOffice and preserve formula caches."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(os.environ.get("VIBE_CODE_ROOT", Path(__file__).resolve().parents[2]))
WORKBOOK = PROJECT_ROOT / "wiki/product/reports/vibe-code-poc-data-specification-and-schema.xlsx"


def main() -> None:
    if not WORKBOOK.is_file():
        raise SystemExit(f"workbook not found: {WORKBOOK}")
    if shutil.which("soffice") is None:
        raise SystemExit("LibreOffice 'soffice' is required for formula recalculation")

    with tempfile.TemporaryDirectory(prefix="vibe-xlsx-recalc-") as tmp:
        out_dir = Path(tmp) / "out"
        profile = Path(tmp) / "lo-profile"
        out_dir.mkdir()
        subprocess.run(
            [
                "soffice",
                "--headless",
                f"-env:UserInstallation={profile.as_uri()}",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(out_dir),
                str(WORKBOOK),
            ],
            check=True,
        )
        recalculated = out_dir / WORKBOOK.name
        if not recalculated.is_file():
            raise SystemExit(f"LibreOffice did not produce {recalculated}")
        shutil.copy2(recalculated, WORKBOOK)

    print(f"recalculated: {WORKBOOK}")


if __name__ == "__main__":
    main()
