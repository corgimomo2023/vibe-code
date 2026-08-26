# Vibe Office report generation

Generate the reviewed Office deliverables from repository source:

```bash
uv run --isolated --frozen --extra reports python scripts/report_generation/generate_vibe_er.py
uv run --isolated --frozen --extra reports python scripts/report_generation/generate_vibe_docs.py
uv run --isolated --frozen --extra reports python scripts/report_generation/generate_vibe_data_spec.py
uv run --isolated --frozen --extra reports python scripts/report_generation/recalc_xlsx.py
uv run --isolated --frozen --extra reports python scripts/report_generation/write_vibe_report_manifest.py
uv run --isolated --frozen --extra reports python scripts/report_generation/export_vibe_reports_to_markdown.py
VIBE_CODE_ROOT="$PWD" uv run --isolated --frozen --extra test python scripts/report_generation/verify_vibe_reports.py
uv run --isolated --frozen --extra test pytest -o addopts='' -q
```

`VIBE_CODE_ROOT` may point to an isolated repository-shaped directory for reproducibility tests. `VIBE_REPORT_ER_DIR` and `VIBE_REPORT_DIAGRAM_DIR` may override temporary build directories; otherwise the OS temporary directory is used. Final outputs are written under `wiki/product/reports/`. `recalc_xlsx.py` requires LibreOffice and must run before verification so the five formula caches are populated. The exporter writes three deterministic Markdown companions for fast diff/search/AI review; Office remains the visual authority.
