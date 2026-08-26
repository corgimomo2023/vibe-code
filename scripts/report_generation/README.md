# Vibe Office report generation

Generate the reviewed Office deliverables from repository source:

```bash
python scripts/report_generation/generate_vibe_er.py
python scripts/report_generation/generate_vibe_docs.py
python scripts/report_generation/generate_vibe_data_spec.py
python scripts/report_generation/recalc_xlsx.py
python scripts/report_generation/verify_vibe_reports.py
```

`VIBE_CODE_ROOT` may point to an isolated repository-shaped directory for reproducibility tests. `VIBE_REPORT_ER_DIR` and `VIBE_REPORT_DIAGRAM_DIR` may override temporary build directories; otherwise the OS temporary directory is used. Final outputs are written under `wiki/product/reports/`. `recalc_xlsx.py` requires LibreOffice and must run before verification so the five formula caches are populated.
