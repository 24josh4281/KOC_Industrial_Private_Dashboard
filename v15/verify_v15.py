"""Check that the published site covers every source row and each detail link resolves."""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from urllib.parse import parse_qs, urlsplit


def read(path):
    return json.loads(path.read_text(encoding='utf-8'))


def verify(site):
    root = site / 'v15'
    manifest = read(root / 'manifest.json')
    stats = read(root / 'data' / 'statistics.json')
    progress = read(root / 'data' / 'progress.json')
    management = read(root / 'details' / 'management.json')
    assert stats['ledger_rows'] == 3419, '관리 Tool 원장 전체 행 누락'
    assert len(progress['management']) == len(management) == stats['management_2026'] == 53
    assert sum(x['count'] for x in stats['all_stage']) == 3419
    assert sum(x['count'] for x in stats['current_stage']) == 53
    assert sum(x['count'] for x in stats['monthly_2026']) == 53
    assert len(progress['ors']) == manifest['ors_row_occurrences']
    detail_files = sorted((root / 'details').glob('ORS*.json'))
    ors_details = {path.stem: read(path) for path in detail_files}
    assert sum(len(group) for group in ors_details.values()) == len(progress['ors'])
    assert len({row['detail_url'] for row in progress['ors']}) == len(progress['ors'])
    for row in progress['management'] + progress['ors']:
        link = urlsplit(row['detail_url'])
        assert link.path == 'case.html'
        query = parse_qs(link.query)
        if row['source'] == 'ORS':
            detail = ors_details[query['sheet'][0]][query['row'][0]]
            assert detail['id'] == row['id'] and detail['source_row'] == row['source_row']
            assert detail['raw_fields'], f"원본 필드 없음: {row['id']}"
        else:
            assert management[query['id'][0]]['id'] == row['id']
    latest = [r for r in progress['ors'] if r['is_latest'] and r['scope'] == '산업']
    assert sum(bool(r['intake_dates_2026']) for r in latest) == progress['overview']['ors_industry_2026_intake']
    assert sum(bool(r['event_count_2026']) for r in latest) == progress['overview']['ors_industry_2026_activity']
    assert sum(x['count'] for x in progress['overview']['management_stages']) == 53
    assert sum(x['count'] for x in progress['overview']['ors_stages']) == progress['overview']['ors_industry_2026_activity']
    assert stats['freshness'] == progress['freshness'] == manifest['freshness']
    for name in ('statistics.html', 'progress.html', 'case.html', 'app.js', 'app.css'):
        assert (root / name).stat().st_size > 100, f'{name} 파일 누락'
    workbook = site / 'v15' / '외부사업_통합관리_v15.xlsx'
    with zipfile.ZipFile(workbook) as archive:
        assert archive.testzip() is None, 'Excel ZIP 손상'
        chart_count = sum(bool(re.search(r'^xl/charts/chart\d+\.xml$', name))
                          for name in archive.namelist())
        assert chart_count >= 8, f'Excel 그래프 부족: {chart_count}'
    return {'management_ledger_rows': stats['ledger_rows'],
            'management_2026_details': len(management),
            'ors_source_rows': len(progress['ors']),
            'ors_detail_sheets': len(detail_files),
            'ors_2026_industry_intake': progress['overview']['ors_industry_2026_intake'],
            'ors_2026_industry_activity': progress['overview']['ors_industry_2026_activity'],
            'workbook_charts': chart_count, 'version': manifest['version']}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--site', type=Path, required=True)
    print(json.dumps(verify(parser.parse_args().site), ensure_ascii=False))
