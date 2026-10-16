#!/usr/bin/env python3
"""Build the public, aggregate-only v19 snapshot from an authorized local export.

The input is never copied to the public directory. Only enumerated dimensions and
counts are written; row identifiers, names, notes, and contact details are excluded.
"""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path


STAGES = ("접수", "검토", "협의", "보완", "보류", "승인", "반려", "철회 요청", "확인 필요")
UNITS = ("방법론 신청", "사업 평가", "감축량 인증", "사업 비공개 신청", "인증 비공개 신청")
REGIONS = ("국내", "국외", "공통")
ACTIVE_STAGES = ("접수", "검토", "협의", "보완", "보류", "철회 요청", "확인 필요")


def ordered(counter, keys):
    return [{"label": key, "count": counter.get(key, 0)} for key in keys]


def build(source):
    data = json.loads(source.read_text(encoding="utf-8-sig"))
    rows = data["rows"]
    if len(rows) != 3419:
        raise ValueError(f"Expected the checked v19 snapshot (3419 applications), got {len(rows)}")
    stages = Counter(str(row.get("stage") or "확인 필요") for row in rows)
    units = Counter(str(row.get("unit") or "") for row in rows)
    regions = Counter(str(row.get("region") or "") for row in rows)
    if set(stages) - set(STAGES) or set(units) - set(UNITS) or set(regions) - set(REGIONS):
        raise ValueError("An unapproved category appeared; review it before publication")

    years = Counter(str(row.get("year") or "미기록") for row in rows)
    if set(years) - {str(year) for year in range(2014, 2027)} - {"미기록"}:
        raise ValueError("An unexpected application year appeared")
    current = [row for row in rows if str(row.get("year")) == "2026"]
    current_stages = Counter(str(row.get("stage") or "확인 필요") for row in current)
    months = Counter(int(row.get("month") or 0) for row in current)
    if set(months) - set(range(1, 13)):
        raise ValueError("An unexpected 2026 month appeared")

    result = {
        "schema_version": "v19-public-aggregate-1",
        "as_of": data["as_of"],
        "source_date": data["source_date"],
        "total": len(rows),
        "current_year": len(current),
        "approved": stages["승인"],
        "active": sum(stages[stage] for stage in ACTIVE_STAGES),
        "year_missing": years["미기록"],
        "years": ordered(years, [str(year) for year in range(2014, 2027)] + ["미기록"]),
        "stages": ordered(stages, STAGES),
        "units": ordered(units, UNITS),
        "regions": ordered(regions, REGIONS),
        "months_2026": [{"label": f"{month:02d}", "count": months[month]} for month in range(1, 13)],
        "stages_2026": ordered(current_stages, STAGES),
    }
    assert sum(item["count"] for item in result["years"]) == result["total"]
    assert sum(item["count"] for item in result["stages"]) == result["total"]
    assert sum(item["count"] for item in result["units"]) == result["total"]
    assert sum(item["count"] for item in result["months_2026"]) == result["current_year"]
    assert sum(item["count"] for item in result["stages_2026"]) == result["current_year"]
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="local v19 data/applications.json")
    parser.add_argument("--output", type=Path, default=Path(__file__).with_name("summary.json"))
    args = parser.parse_args()
    summary = build(args.source)
    args.output.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Public aggregate: {summary['total']} applications; {args.output}")


if __name__ == "__main__":
    main()
