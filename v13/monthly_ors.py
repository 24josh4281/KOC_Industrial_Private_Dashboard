"""Safe, repeatable ORS ZIP intake. Use --commit only after inspecting the JSON preview."""
import argparse
import gzip
import hashlib
import json
import shutil
import sqlite3
import sys
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'python'))
legacy_deps=ROOT.parent/'2026-10-15_meeting_v9/work/pydeps'
if legacy_deps.exists():sys.path.insert(0,str(legacy_deps))
from kets_dashboard.ors_refresh import read_exports, dated_filename
from ors_full import read_full_zip


def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for part in iter(lambda: f.read(1024 * 1024), b''): h.update(part)
    return h.hexdigest()


def scope(name):
    return '산업' if '산업' in name else '발전' if '발전' in name else '기후' if '기후' in name else '확인 필요'


def main():
    p = argparse.ArgumentParser(description='월별 ORS ZIP을 원본 보존 DB에 단계적으로 등록')
    p.add_argument('zip_file', type=Path, help='새로 수령한 ORS ZIP 경로')
    p.add_argument('--commit', action='store_true', help='검증 보고서 확인 후 원본 스냅샷을 DB에 등록')
    p.add_argument('--db', type=Path, help='별도 시험 DB. 기본값은 state/working.sqlite3')
    args = p.parse_args()
    src = args.zip_file.resolve(strict=True)
    if src.suffix.lower() != '.zip': raise SystemExit('ZIP 파일만 받습니다.')
    zip_sha = digest(src)
    with zipfile.ZipFile(src) as z:
        members = [i.filename for i in z.infolist() if not i.is_dir()]
        if len(members) != len(set(members)): raise SystemExit('ZIP 내부 파일명이 중복됩니다. 원본을 확인하세요.')
    full = read_full_zip(src)
    parse_issue = None
    try: exports = read_exports([src])
    except (ValueError, RuntimeError) as exc:
        exports = []
        parse_issue = str(exc)
    parsed = {(e['sha256'],e['sheet']) for e in exports}
    for raw in full:
        if (raw['sha256'],raw['sheet']) not in parsed:
            exports.append({'filename':raw['filename'],'sheet':raw['sheet'],'sha256':raw['sha256'],'as_of':dated_filename(raw['filename']),'kind':None,'columns':[],'rows':[]})
    if not full: raise SystemExit('ZIP 안에 XLS 원본이 없습니다.')
    dbpath=args.db
    if dbpath is None:
        state=ROOT/'state';state.mkdir(exist_ok=True)
        dbpath=state/'working.sqlite3'
        if not dbpath.exists():
            if args.commit:shutil.copy2(ROOT/'external_projects_v13.sqlite3',dbpath)
            else:dbpath=ROOT/'external_projects_v13.sqlite3'
    con = sqlite3.connect(dbpath)
    con.execute('PRAGMA foreign_keys=ON')
    duplicate = con.execute('SELECT 1 FROM source_files WHERE sha=?', (zip_sha,)).fetchone() is not None
    checks = []
    for e in exports:
        s = scope(e['filename'])
        issues = []
        if parse_issue: issues.append('표 구조/원본 오류: '+parse_issue[:180])
        if e['as_of'] is None: issues.append('파일명 기준일 미확인')
        elif e['as_of'] > date.today().isoformat(): issues.append('미래 기준일')
        if e['kind'] is None: issues.append('자동 분류 미지원: 원본 보존만')
        if s == '확인 필요': issues.append('소관 범위 확인 필요')
        if e['kind'] and s != '확인 필요':
            previous = con.execute('SELECT as_of,header_json,filename FROM ors_members WHERE kind=? AND scope=? ORDER BY as_of DESC LIMIT 1', (e['kind'], s)).fetchone()
            if previous:
                if e['as_of'] and previous[0] and e['as_of'] < previous[0]: issues.append(f'과거 기준일: 최신 {previous[0]}')
                old_headers = json.loads(previous[1])
                old_labels = [x['label'].strip() for x in old_headers]
                new_labels = [x['label'].strip() for x in e['columns']]
                if new_labels != old_labels: issues.append('열 구조 변경: 매핑 재검토')
        checks.append({'member': e['filename'], 'sheet': e['sheet'], 'sha256': e['sha256'], 'as_of': e['as_of'], 'kind': e['kind'], 'scope': s, 'rows': len(e['rows']), 'columns': len(e['columns']), 'issues': issues, 'promotion': '보류' if issues else '원본 스냅샷 등록 가능'})
    report = {'zip': src.name, 'sha256': zip_sha, 'size': src.stat().st_size, 'duplicate': duplicate, 'mode': 'commit' if args.commit else 'dry_run', 'files': checks, 'totals': {'members': len(exports), 'rows': sum(len(e['rows']) for e in exports), 'full_source_cells': sum(x['cells'] for x in full), 'review_required': sum(bool(x['issues']) for x in checks)}, 'policy': '원본 ZIP·행은 보존합니다. 기존 관리 원장·업무 의견·사무국 안내를 자동 덮어쓰지 않습니다. 보류 자료는 자동 최신 현황에 승격하지 않습니다.'}
    reports = ROOT / 'imports'
    reports.mkdir(exist_ok=True)
    report_path = reports / f'{zip_sha[:12]}_{"commit" if args.commit else "preview"}.json'
    if args.commit and not duplicate:
        backup_dir=(ROOT if args.db is None else dbpath.parent)/'backups';backup_dir.mkdir(exist_ok=True)
        backup=backup_dir/f'{datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")}_{zip_sha[:12]}_before.sqlite3'
        backup_con=sqlite3.connect(backup)
        con.backup(backup_con);backup_con.close()
        report['backup']=str(backup)
        archive_root=ROOT/'source-archive' if args.db is None else dbpath.parent/'source-archive'
        dest_dir = archive_root / 'monthly' / (max((x['as_of'] or 'unknown' for x in checks))[:7])
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f'{zip_sha[:12]}_{src.name}'
        if not dest.exists(): shutil.copy2(src, dest)
        if digest(dest) != zip_sha: raise SystemExit('보관본 해시가 원본과 다릅니다. DB를 변경하지 않았습니다.')
        try:
            con.execute('BEGIN IMMEDIATE')
            con.execute('INSERT INTO source_files VALUES(?,?,?,?,?)', (zip_sha, src.name, src.stat().st_size, 'ors_zip', max((x['as_of'] or '' for x in checks))))
            con.execute('CREATE TABLE IF NOT EXISTS ors_raw_members(member_sha TEXT NOT NULL,sheet TEXT NOT NULL,rows_gzip BLOB NOT NULL,cell_count INTEGER NOT NULL,nrows INTEGER NOT NULL,ncols INTEGER NOT NULL,error_count INTEGER NOT NULL,PRIMARY KEY(member_sha,sheet))')
            for raw in full:
                raw_bytes=json.dumps(raw['rows'],ensure_ascii=False,separators=(',',':')).encode('utf-8')
                con.execute('INSERT OR IGNORE INTO ors_raw_members VALUES(?,?,?,?,?,?,?)',(raw['sha256'],raw['sheet'],gzip.compress(raw_bytes,compresslevel=9),raw['cells'],raw['nrows'],raw['ncols'],raw['errors']))
            for e, check in zip(exports, checks):
                con.execute('INSERT OR IGNORE INTO ors_members VALUES(?,?,?,?,?,?,?)', (e['sha256'], e['sheet'], e['filename'], e['kind'], check['scope'], e['as_of'], json.dumps(e['columns'], ensure_ascii=False)))
                con.execute('INSERT OR IGNORE INTO ors_occurrences VALUES(?,?,?,?)', (zip_sha, e['sha256'], e['sheet'], e['filename']))
                con.executemany('INSERT OR IGNORE INTO ors_rows VALUES(?,?,?,?)', [(e['sha256'], e['sheet'], r['row'], json.dumps(r, ensure_ascii=False)) for r in e['rows']])
            con.execute('INSERT INTO import_runs VALUES(?,?,?,?)', (zip_sha, 'ors_zip', datetime.now(timezone.utc).isoformat(), json.dumps(report, ensure_ascii=False)))
            con.commit()
        except Exception:
            con.rollback()
            raise
    con.close()
    if not report_path.exists(): report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'report': str(report_path), 'duplicate': duplicate, 'mode': report['mode'], 'members': len(exports), 'rows': report['totals']['rows'], 'review_required': report['totals']['review_required']}, ensure_ascii=False))


if __name__ == '__main__': main()
