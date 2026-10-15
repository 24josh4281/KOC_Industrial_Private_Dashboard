"""Validate monthly ORS ZIPs and build the public site from one database snapshot."""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
V13, V14, V15 = (ROOT / name for name in ('v13', 'v14', 'v15'))


def run(*args):
    result = subprocess.run([sys.executable, *map(str, args)],
                            capture_output=True, text=True, encoding='utf-8')
    if result.returncode:
        raise RuntimeError(f"{' '.join(map(str, args))}: {result.stderr.strip()}")
    return json.loads(result.stdout.strip().splitlines()[-1])


def incoming_sources(root, legacy):
    return sorted([*root.joinpath('incoming').rglob('*.zip'),
                   *legacy.joinpath('incoming').rglob('*.zip')])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--site', required=True, type=Path)
    parser.add_argument('--work', required=True, type=Path)
    args = parser.parse_args()
    site, work = args.site.resolve(), args.work.resolve()
    if site == work or site == ROOT or work == ROOT:
        raise SystemExit('산출·작업 폴더는 저장소와 서로 분리해야 합니다.')
    work.mkdir(parents=True, exist_ok=True)
    site.mkdir(parents=True, exist_ok=True)
    db = work / 'working.sqlite3'
    if not db.exists():
        shutil.copy2(V13 / 'external_projects_v13.sqlite3', db)
    accepted = []
    for source in incoming_sources(ROOT, V13):
        preview = run(V13 / 'monthly_ors.py', source, '--db', db)
        if preview['duplicate']:
            accepted.append({'file': str(source.relative_to(ROOT)), 'result': 'already_registered'})
            continue
        report = json.loads(Path(preview['report']).read_text(encoding='utf-8'))
        if report['totals']['review_required']:
            raise SystemExit(f"검증 보류: {source.name} / 검토 필요 {report['totals']['review_required']}개 시트. 기존 공개 판본을 유지합니다.")
        committed = run(V13 / 'monthly_ors.py', source, '--commit', '--db', db)
        accepted.append({'file': str(source.relative_to(ROOT)), 'result': 'registered',
                         'rows': committed['rows']})
    workbook = run(V14 / 'build_v14.py', '--db', db)
    output14 = Path(workbook['output'])
    target14 = site / 'v14'
    target14.mkdir(parents=True, exist_ok=True)
    for name in ('통계_대시보드_v14.html', '진행현황_대시보드_v14.html'):
        shutil.copy2(output14 / name, target14 / name)
    shutil.copy2(V14 / '외부사업_통합관리_v14.xlsx',
                 target14 / '외부사업_통합관리_v14.xlsx')
    target13 = site / 'v13'
    target13.mkdir(parents=True, exist_ok=True)
    for name in ('대시보드_v13.html', '외부사업_통합관리_전체자료_v13_검증완료.xlsx'):
        shutil.copy2(V13 / name, target13 / name)
    shutil.copytree(V13 / 'source-archive', target13 / 'source-archive',
                    dirs_exist_ok=True)
    built = run(V15 / 'build_v15.py', '--db', db, '--output', site / 'v15')
    shutil.copy2(output14 / '외부사업_통합관리_v14.xlsx',
                 site / 'v15' / '외부사업_통합관리_v15.xlsx')
    shutil.copy2(V15 / 'index.html', site / 'index.html')
    (site / '.nojekyll').touch()
    manifest = json.loads((site / 'v15' / 'manifest.json').read_text(encoding='utf-8'))
    progress = json.loads((site / 'v15' / 'data' / 'progress.json').read_text(encoding='utf-8'))
    statistics = json.loads((site / 'v15' / 'data' / 'statistics.json').read_text(encoding='utf-8'))
    if len(progress['management']) != 53 or statistics['ledger_rows'] != 3419:
        raise SystemExit('관리 Tool 기준 원장 검증 실패. 배포하지 않습니다.')
    if len(progress['ors']) != manifest['ors_row_occurrences']:
        raise SystemExit('ORS 원본행 수가 상세 판본과 일치하지 않습니다. 배포하지 않습니다.')
    if not (site / 'v15' / 'case.html').exists():
        raise SystemExit('상세 페이지가 빠졌습니다. 배포하지 않습니다.')
    print(json.dumps({'site': str(site), 'zip_imports': accepted,
                      'manifest': manifest, 'build': built,
                      'workbook_sheets': workbook['excel']['sheets']}, ensure_ascii=False))


if __name__ == '__main__':
    main()
