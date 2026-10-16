#!/usr/bin/env python3
"""Validate the public release without needing the private source files."""

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parent
data = json.loads((ROOT / "summary.json").read_text(encoding="utf-8"))
allowed = {
    "schema_version", "as_of", "source_date", "total", "current_year", "approved",
    "active", "year_missing", "years", "stages", "units", "regions", "months_2026", "stages_2026",
}
assert set(data) == allowed
for key in ("years", "stages", "units", "regions", "months_2026", "stages_2026"):
    assert all(set(row) == {"label", "count"} and isinstance(row["count"], int) for row in data[key])
for key in ("years", "stages", "units", "regions"):
    assert sum(row["count"] for row in data[key]) == data["total"]
for key in ("months_2026", "stages_2026"):
    assert sum(row["count"] for row in data[key]) == data["current_year"]
assert next(row["count"] for row in data["stages"] if row["label"] == "승인") == data["approved"]
assert next(row["count"] for row in data["years"] if row["label"] == "미기록") == data["year_missing"]
assert data["total"] == 3419 and data["current_year"] == 53
assert (ROOT / "summary.json").stat().st_size < 6000
for forbidden in ("applications.json", "source-tools.json", "state.sqlite3", "외부사업_통합관리_v19.xlsx"):
    assert not (ROOT / forbidden).exists(), forbidden
print("v19 public release: aggregate counts and file allowlist passed")
