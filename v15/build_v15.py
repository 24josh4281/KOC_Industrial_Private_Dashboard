"""Build three traceable static pages from the checked management/ORS snapshot."""
from __future__ import annotations

import argparse
import base64
import gzip
import hashlib
import json
import os
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
V13 = next((path for path in (ROOT.parent / 'v13', ROOT.parent / '2026-10-15_meeting_v13') if path.exists()), None)
V14 = next((path for path in (ROOT.parent / 'v14', ROOT.parent / '2026-10-15_meeting_v14') if path.exists()), None)
if V14 is None:
    raise FileNotFoundError('v14 원본 행 모델이 필요합니다.')
sys.path.insert(0, str(V14))
from ors_model import collect, events_in, packed_json  # noqa: E402
from model_v15 import (build_management, build_ors, link_candidates,
                       management_stats, month_counts, stage_counts)  # noqa: E402


def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + '.building')
    temporary.write_text(json.dumps(value, ensure_ascii=False, separators=(',', ':')),
                         encoding='utf-8')
    os.replace(temporary, path)


def get_source(db):
    if V13 is None:
        raise FileNotFoundError('v13 기준 자료가 없습니다.')
    if db:
        result = subprocess.run([sys.executable, str(V13 / 'refresh_outputs.py'), '--db', str(db)],
                                text=True, encoding='utf-8', capture_output=True, check=True)
        snapshot = Path(json.loads(result.stdout.strip().splitlines()[-1])['snapshot'])
    else:
        snapshot = V13
    for name in ('dashboard-data.json', 'ors-full-catalog.gz.b64'):
        if not (snapshot / name).exists():
            raise FileNotFoundError(snapshot / name)
    return snapshot


def manifest_version(snapshot, as_of):
    digest = hashlib.sha256()
    for path in (snapshot / 'dashboard-data.json', snapshot / 'ors-full-catalog.gz.b64',
                 ROOT / 'model_v15.py', ROOT / 'app.css', ROOT / 'app.js',
                 ROOT / 'statistics.html', ROOT / 'progress.html', ROOT / 'case.html'):
        digest.update(path.name.encode())
        if path.name == 'dashboard-data.json':
            payload = json.loads(path.read_text(encoding='utf-8'))
            payload.get('refresh_snapshot', {}).pop('created_at', None)
            source = json.dumps(payload, ensure_ascii=False, sort_keys=True,
                                separators=(',', ':')).encode('utf-8')
        elif path.name == 'ors-full-catalog.gz.b64':
            source = gzip.decompress(base64.b64decode(path.read_bytes()))
        else:
            source = path.read_bytes()
        digest.update(hashlib.sha256(source).digest())
    digest.update(as_of.encode('ascii'))
    return digest.hexdigest()[:16]


def build(snapshot, output):
    data = json.loads((snapshot / 'dashboard-data.json').read_text(encoding='utf-8'))
    catalog = packed_json(snapshot / 'ors-full-catalog.gz.b64')
    progress = collect(data, catalog)
    as_of = date.today().isoformat() if data.get('refresh_snapshot') else data['as_of']
    freshness = {'analysis_date': as_of,
                 'management_date': data.get('refresh_snapshot', {}).get('management_snapshot_date') or '2026-10-13',
                 'ors_date': data.get('refresh_snapshot', {}).get('latest_ors_date') or
                             max((item.get('date', '') for item in data.get('ors_imports', [])), default='')}
    management, management_details = build_management(data, as_of)
    ors, ors_details = build_ors(progress, catalog, as_of, events_in)
    candidates = link_candidates(management, ors)
    for record in management:
        management_details[record['id']]['match_candidate'] = record['match_candidate']
    for record in ors:
        ors_details[record['raw_excel_sheet']][str(record['source_row'])]['match_candidate'] = record['match_candidate']

    latest_industry = [row for row in ors if row['is_latest'] and row['scope'] == '산업']
    industry_intake = [row for row in latest_industry if row['intake_dates_2026']]
    industry_activity = [row for row in latest_industry if row['event_count_2026']]
    statistics = management_stats(data, management)
    statistics.update({'as_of': as_of, 'freshness': freshness, 'source': '관리 Tool Excel 2종',
                       'ors_reference': {'source_occurrences': len(ors),
                                         'latest_industry_rows': len(latest_industry),
                                         'industry_intake_2026': len(industry_intake),
                                         'industry_activity_2026': len(industry_activity)}})
    progress_data = {
        'as_of': as_of, 'freshness': freshness, 'management': management, 'ors': ors,
        'overview': {'management_2026': len(management),
                     'ors_industry_2026_intake': len(industry_intake),
                     'ors_industry_2026_activity': len(industry_activity),
                     'ors_all_occurrences': len(ors),
                     'ors_latest_industry': len(latest_industry),
                     'candidate_matches': candidates,
                     'management_monthly': month_counts(management),
                     'ors_monthly': month_counts(industry_intake),
                     'management_stages': stage_counts(management),
                     'ors_stages': stage_counts(industry_activity)}}
    version = manifest_version(snapshot, as_of)
    output.mkdir(parents=True, exist_ok=True)
    write_json(output / 'data' / 'statistics.json', statistics)
    write_json(output / 'data' / 'progress.json', progress_data)
    write_json(output / 'details' / 'management.json', management_details)
    for sheet, details in ors_details.items():
        write_json(output / 'details' / f'{sheet}.json', details)
    for name in ('statistics.html', 'progress.html', 'case.html', 'app.css', 'app.js'):
        if (ROOT / name).resolve() != (output / name).resolve():
            shutil.copy2(ROOT / name, output / name)
    write_json(output / 'manifest.json', {'version': version, 'as_of': as_of,
                'freshness': freshness,
                'source_snapshot': snapshot.name, 'ors_file_occurrences': len(catalog),
                'ors_row_occurrences': len(ors), 'management_2026': len(management)})
    return {'version': version, 'output': str(output), 'as_of': as_of,
            'management_2026': len(management), 'ors_rows': len(ors),
            'ors_2026_intake_industry': len(industry_intake),
            'ors_2026_activity_industry': len(industry_activity),
            'matched_candidates_not_merged': candidates,
            'detail_sheets': len(ors_details)}


def main():
    parser = argparse.ArgumentParser(description='v15 통계·진행·개별 상세 화면 생성')
    parser.add_argument('--db', type=Path, help='월별 ORS 반영 DB')
    parser.add_argument('--output', type=Path, help='출력 폴더; 기본값은 v15 또는 v15/snapshots/판본')
    args = parser.parse_args()
    snapshot = get_source(args.db)
    output = args.output or (ROOT / 'snapshots' / snapshot.name if args.db else ROOT)
    print(json.dumps(build(snapshot, output), ensure_ascii=False))


if __name__ == '__main__':
    main()
