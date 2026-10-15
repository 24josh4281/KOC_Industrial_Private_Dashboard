"""Traceable ORS progress rows for the two v14 dashboards.

Every source occurrence remains a row. `is_latest` selects the newest file
per data type and scope for charts, without deleting older snapshots.
"""
import base64
import gzip
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

DATE = re.compile(r'(?<!\d)(20\d{2})[-./](\d{1,2})[-./](\d{1,2})')
EVENT = re.compile(r'^\s*(20\d{2}[-./]\d{1,2}[-./]\d{1,2})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\s*-\s*(.*)$')


def packed_json(path: Path):
    return json.loads(gzip.decompress(base64.b64decode(path.read_text(encoding='ascii'))))


def date_iso(value):
    match = DATE.search(str(value or ''))
    if not match:
        return ''
    y, m, d = map(int, match.groups())
    if not (1 <= m <= 12 and 1 <= d <= 31):
        return ''
    return f'{y:04d}-{m:02d}-{d:02d}'


def infer_kind(meta, filename):
    if meta.get('kind'):
        return meta['kind']
    if '국외' in filename and '인증실적' in filename:
        return 'disclosure_credit_foreign'
    if '국외' in filename:
        return 'disclosure_project_foreign'
    return 'unclassified'


def events_in(history):
    events = []
    for line in str(history or '').splitlines():
        match = EVENT.match(line.strip())
        if not match:
            continue
        date = date_iso(match.group(1))
        if date:
            parts = [part.strip() for part in match.group(2).split(' - ')]
            events.append({'date': date, 'phase': parts[0],
                           'action': parts[-1], 'text': line.strip()})
    return events


def intake_events(direct_date, events):
    """Return source events, never an inferred count of unique applications."""
    if direct_date:
        return [{'date': direct_date, 'basis': '명시 신청일'}]
    project = []
    for event in events:
        phase, action = event['phase'], event['action']
        if ('신청' in phase and '보완' not in phase and '재제출' not in phase
                and action == '접수'):
            project.append({'date': event['date'], 'basis': f'{phase} 접수 이력'})
    if project:
        return project
    # Certification applications have a two-part "date - 신청" history. The
    # first such event is the intake; later repeats can be supplementation.
    first = next((e for e in events if e['action'] == '신청'), None)
    return [{'date': first['date'], 'basis': '최초 신청 이력'}] if first else []


def status_bucket(raw):
    first = str(raw or '').splitlines()[0].strip()
    first = re.sub(r'^\[보완\d+차\]\s*', '', first)
    return first or '상태 미기록'


def collect(data, catalog):
    meta = {(x['file'], x['sha256'], x['sheet']): x for x in data['ors_imports']}
    chosen = {}
    for item in catalog:
        m = meta.get((item['filename'], item['sha256'], item['sheet']))
        if m is None:
            raise ValueError(f'ORS 출처 메타데이터 누락: {item["filename"]}')
        group = (infer_kind(m, item['filename']), m['scope'], item['sheet'])
        chosen[group] = max(m['date'], chosen.get(group, ''))
    result = []
    for sheet_no, item in enumerate(catalog, 1):
        m = meta[(item['filename'], item['sha256'], item['sheet'])]
        kind = infer_kind(m, item['filename'])
        is_latest = m['date'] == chosen[(kind, m['scope'], item['sheet'])]
        headers = {str(value).strip(): col for col, value, *_ in item['rows'][0][1]}

        def field(values, *labels):
            for label in labels:
                col = headers.get(label)
                if col and values.get(col) not in (None, ''):
                    return values[col]
            return ''

        for source_row, cells in item['rows'][1:]:
            values = {col: value for col, value, *_ in cells}
            history = str(field(values, '이력'))
            events = events_in(history)
            direct_date = date_iso(field(values, '신청일'))
            status_date = date_iso(field(values, '처리일자', '처리일'))
            intakes = intake_events(direct_date, events)
            intakes_2026 = [x for x in intakes if x['date'].startswith('2026-')]
            activity = [x['text'] for x in events if x['date'].startswith('2026-')]
            if direct_date.startswith('2026-') and not any(x['date']==direct_date for x in events):
                activity.append(f'{direct_date} - 원본 신청일')
            if status_date.startswith('2026-') and not any(x['date']==status_date for x in events):
                activity.append(f'{status_date} - 원본 처리일: {status_bucket(field(values,"처리상태","진행상태"))}')
            title = field(values, '[신청정보] 신청사업명',
                          '[사업정보]사업명(단위사업명)', '사업명')
            number = field(values, '외부사업 고유번호', '외부사업 등록번호',
                           '인증실적 고유번호', '번호')
            status = field(values, '처리상태', '진행상태')
            applicant = field(values, '[신청자] 사업자명', '업체명')
            result.append({
                'id': f'{item["container"]}:{item["filename"]}:{item["sha256"][:16]}:{item["sheet"]}:{source_row}',
                'source_file': item['filename'], 'source_zip': item['container'],
                'source_sheet': item['sheet'], 'source_row': source_row,
                'raw_excel_sheet': f'ORS{sheet_no:02d}', 'snapshot_date': m['date'],
                'is_latest': is_latest, 'kind': kind, 'scope': m['scope'],
                'title': str(title or ''), 'applicant': str(applicant or ''),
                'number': str(number or ''), 'application_type': str(field(values, '신청유형') or ''),
                'status': str(status or ''), 'status_bucket': status_bucket(status),
                'status_date': status_date,
                'intake_dates_2026': [x['date'] for x in intakes_2026],
                'intake_basis_2026': [x['basis'] for x in intakes_2026],
                'event_count_2026': len(activity),
                'last_event_date': events[-1]['date'] if events else '',
                'last_event': events[-1]['text'] if events else '',
                'activity_2026': activity,
            })
    return result


def summarize(data, progress):
    latest = [x for x in progress if x['is_latest']]
    industry = [x for x in latest if x['scope'] == '산업']
    intake = [x for x in latest if x['intake_dates_2026']]
    industry_intake = [x for x in intake if x['scope'] == '산업']
    history = [x for x in latest if x['event_count_2026']]
    yearly = Counter(r['year'] for r in data['rows'] if r.get('year') and r['year'].isdigit()
                     and int(r['year']) >= 2015)
    month_2026 = Counter(r['month'] for r in data['cases_2026'] if r.get('month'))
    intake_month = Counter(date[5:7] for x in industry_intake for date in x['intake_dates_2026'])
    status_2026 = Counter(x['status_bucket'] for x in industry
                          if x['intake_dates_2026'] or x['event_count_2026'])
    return {
        'source_occurrence_rows': len(progress), 'latest_ors_rows': len(latest),
        'latest_industry_rows': len(industry), 'ors_2026_intake_rows_all_scopes': len(intake),
        'ors_2026_intake_events_all_scopes': sum(len(x['intake_dates_2026']) for x in intake),
        'ors_2026_intake_rows_industry': len(industry_intake),
        'ors_2026_intake_events_industry': sum(len(x['intake_dates_2026']) for x in industry_intake),
        'ors_2026_activity_rows_all_scopes': len(history),
        'ors_2026_activity_rows_industry': sum(1 for x in industry if x['event_count_2026']),
        'management_rows': len(data['rows']), 'management_2026_rows': len(data['cases_2026']),
        'management_yearly': dict(sorted(yearly.items())),
        'management_month_2026': {f'{m:02d}': month_2026.get(f'{m:02d}', 0) for m in range(1,13)},
        'industry_intake_month_2026': {f'{m:02d}': intake_month.get(f'{m:02d}', 0) for m in range(1,13)},
        'industry_2026_status': dict(status_2026.most_common()),
        'latest_by_scope': dict(Counter(x['scope'] for x in latest)),
        'latest_by_kind': dict(Counter(x['kind'] for x in latest)),
    }
