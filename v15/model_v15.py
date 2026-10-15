"""Business views for the management ledger and traceable ORS source rows."""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from datetime import date


STAGES = ('접수', '심사', '의견 수렴', '보완', '승인·통과', '반려', '후속 관리', '확인 필요')
KIND_NAMES = {
    'domestic_project': '국내 사업신청', 'domestic_credit': '국내 인증신청',
    'foreign_project': '국외 사업신청', 'foreign_credit': '국외 인증신청',
    'disclosure_project_domestic': '비공개 사업승인',
    'disclosure_credit_domestic': '비공개 인증승인',
    'disclosure_project_foreign': '국외 비공개 사업',
    'disclosure_credit_foreign': '국외 비공개 인증',
    'unclassified': '분류 확인 필요',
}


def stage_for(value):
    status = str(value or '').replace(' ', '')
    if any(word in status for word in ('반려', '불승인', '기각')):
        return '반려'
    if any(word in status for word in ('심의승인', '검토승인', '최종승인', '승인')):
        return '승인·통과'
    if any(word in status for word in ('보완', '재제출')):
        return '보완'
    if any(word in status for word in ('협의', '의견')):
        return '의견 수렴'
    if any(word in status for word in ('검토', '심사', '평가', '심의')):
        return '심사'
    if any(word in status for word in ('착수', '사후', '운영')):
        return '후속 관리'
    if any(word in status for word in ('신청', '접수', '갱신')):
        return '접수'
    return '확인 필요'


def month_counts(records, key='intake_date'):
    counts = Counter(str(record.get(key) or '')[:7] for record in records
                     if str(record.get(key) or '').startswith('2026-'))
    return [{'month': f'2026-{month:02d}', 'count': counts[f'2026-{month:02d}']}
            for month in range(1, 13)]


def stage_counts(records):
    counter = Counter(record['stage'] for record in records)
    return [{'stage': stage, 'count': counter[stage]} for stage in STAGES]


def clean(value):
    return '' if value is None else str(value).strip()


def norm_title(value):
    return re.sub(r'[^가-힣A-Za-z0-9]', '', clean(value)).lower()


def days_old(start, as_of):
    try:
        return (date.fromisoformat(as_of[:10]) - date.fromisoformat(start[:10])).days
    except (ValueError, TypeError):
        return None


def field(raw, *names):
    for name in names:
        value = raw.get(name)
        if value not in (None, ''):
            return clean(value)
    return ''


def attention_for(source, stage, status, *, intake_date='', last_date='', agency='',
                  reviewer='', issues=(), opinion_status='', guidance_status='', as_of=''):
    """Separate observed source issues from internal follow-up suggestions."""
    result = []
    seen = set()

    def add(label, basis, level='확인', origin='내부 확인 제안'):
        if label not in seen:
            result.append({'label': label, 'basis': basis, 'level': level, 'origin': origin})
            seen.add(label)

    for issue in issues or ():
        if issue:
            add(clean(issue), '관리 Tool 원본의 이슈 열', '확인', '원본 기록')
    if source == '관리 Tool':
        if not reviewer:
            add('검토기관 확인', '관리 Tool 검토기관 미기록')
        if stage == '보완' and guidance_status in ('', '미기록', '확인 필요'):
            add('신청자 안내 기록 필요', '보완 상태이나 사무국 안내가 기록되지 않음', '조치')
        if stage in ('심사', '의견 수렴') and opinion_status in ('', '미기록', '확인 필요'):
            add('의견 검토 기록 필요', '심사·협의 상태이나 의견 검토가 기록되지 않음', '조치')
    else:
        if stage == '보완':
            add('보완 사유·안내 확인', 'ORS 상태가 보완/재제출이며 세부 사유는 원문 확인 필요', '조치')
        if stage == '의견 수렴':
            add('협의 의견 확인', 'ORS 상태가 협의·의견 단계', '조치')
        if stage == '반려':
            add('반려 사유 확인', 'ORS 상태가 반려로 표시됨', '조치')
        if not agency:
            add('관장기관 원문 미기록', 'ORS의 관장기관 칸이 비어 있음')
        if not intake_date and last_date.startswith('2026-'):
            add('최초 접수일 확인', '2026 처리 이력은 있으나 신규 접수 근거일이 확인되지 않음')
    since = days_old(last_date or intake_date, as_of)
    if since is not None and since >= 30 and stage not in ('승인·통과', '반려', '후속 관리'):
        add('30일 이상 변화 확인', f'마지막 확인 날짜로부터 {since}일 · 내부 점검 기준', '점검')
    return result


def build_management(data, as_of):
    source_by_case = {row.get('case_key'): row for row in data['rows'] if row.get('case_key')}
    records, details = [], {}
    for case in data['cases_2026']:
        key = case['case_key']
        full = source_by_case.get(key, {})
        workflow = data.get('workflow', {}).get(key, {})
        status = clean(case.get('raw_status'))
        stage = stage_for(status)
        reviewer = clean(case.get('reviewer') or full.get('reviewer'))
        agency = clean(full.get('original_assignment') or reviewer)
        intake = clean(case.get('apply_date'))
        last = clean(case.get('processed')) or intake
        flags = attention_for('관리 Tool', stage, status, intake_date=intake,
            last_date=last, agency=agency, reviewer=reviewer, issues=case.get('issues'),
            opinion_status=clean(workflow.get('opinion_status')),
            guidance_status=clean(workflow.get('guidance_status')), as_of=as_of)
        record = {
            'id': key, 'source': '관리 Tool', 'title': clean(case.get('title')),
            'applicant': clean(full.get('company')), 'kind': clean(case.get('unit')),
            'intake_date': intake, 'last_date': last, 'status': status, 'stage': stage,
            'agency': agency, 'reviewer': reviewer, 'risk': flags,
            'opinion': clean(workflow.get('opinion_text')),
            'guidance': clean(workflow.get('guidance_text')),
            'next_action': clean(workflow.get('next_action') or case.get('next_action')),
            'source_file': clean(case.get('source_file')),
            'source_sheet': clean(case.get('source_sheet')),
            'source_row': case.get('source_row'),
            'detail_url': f'case.html?source=management&id={key}',
            'match_candidate': None,
        }
        records.append(record)
        details[key] = {
            **record, 'source_key': clean(case.get('source_key')),
            'application_type': clean(full.get('type')),
            'sector': clean(full.get('sector')), 'method': clean(full.get('method')),
            'region': clean(case.get('region')), 'owner': clean(case.get('owner')),
            'first_agency': clean(full.get('original_assignment')),
            'assignments': full.get('assignments') or case.get('assignments') or [],
            'notes': clean(full.get('notes') or case.get('notes')),
            'history': clean(full.get('history') or case.get('history')),
            'visit': clean(case.get('visit')), 'supplement_round': clean(case.get('supplement_round')),
            'opinion_status': clean(workflow.get('opinion_status')),
            'guidance_status': clean(workflow.get('guidance_status')),
            'intake_status': clean(workflow.get('intake_status')),
            'workflow_updated_at': clean(workflow.get('updated_at')),
            'raw_fields': [{'label': label, 'value': value} for label, value in full.items()
                           if value not in (None, '', [], {})],
        }
    return records, details


def build_ors(progress, catalog, as_of, events_in):
    records, details_by_sheet = [], {}
    index = {(x['raw_excel_sheet'], x['source_row']): x for x in progress}
    for sheet_index, item in enumerate(catalog, 1):
        sheet_code = f'ORS{sheet_index:02d}'
        headers = {cell[0]: clean(cell[1]) for cell in item['rows'][0][1]}
        sheet_details = {}
        for row_no, cells in item['rows'][1:]:
            original = index[(sheet_code, row_no)]
            raw = {headers.get(c[0], str(c[0])): c[1] for c in cells}
            status = original['status_bucket']
            stage = stage_for(status)
            agency = field(raw, '관장기관', '검토기관')
            intake = original['intake_dates_2026'][0] if original['intake_dates_2026'] else ''
            last = original['last_event_date'] or original['status_date'] or intake
            flags = attention_for('ORS', stage, status, intake_date=intake,
                                  last_date=last, agency=agency, as_of=as_of)
            url = f'case.html?source=ors&sheet={sheet_code}&row={row_no}'
            record = {
                'id': original['id'], 'source': 'ORS', 'title': original['title'],
                'applicant': original['applicant'], 'kind': KIND_NAMES.get(original['kind'], original['kind']),
                'intake_date': intake, 'intake_dates_2026': original['intake_dates_2026'],
                'intake_basis_2026': original['intake_basis_2026'],
                'last_date': last, 'status': original['status'], 'stage': stage,
                'agency': agency, 'reviewer': '', 'risk': flags,
                'opinion': '', 'guidance': '', 'next_action': '',
                'source_file': original['source_file'], 'source_zip': original['source_zip'],
                'source_sheet': original['source_sheet'], 'source_row': row_no,
                'raw_excel_sheet': sheet_code, 'scope': original['scope'],
                'kind_code': original['kind'], 'is_latest': original['is_latest'],
                'event_count_2026': original['event_count_2026'],
                'snapshot_date': original['snapshot_date'], 'detail_url': url,
                'match_candidate': None,
            }
            records.append(record)
            timeline = events_in(field(raw, '이력'))
            sheet_details[str(row_no)] = {
                **record, 'number': original['number'],
                'application_type': original['application_type'],
                'project_type': field(raw, '사업유형', '사업구분'),
                'project_name': field(raw, '[사업정보] 사업명', '[사업정보]사업명(단위사업명)', '사업명'),
                'method': field(raw, '적용방법론', '[정책사업정보] 적용방법론'),
                'sector': field(raw, '사업분야'),
                'location': field(raw, '사업시행장소'),
                'expected_reduction': field(raw, '연평균 온실가스 배출감축량(흡수량)'),
                'first_agency': agency,
                'agency_events': [event for event in timeline if '기관' in event['phase']],
                'timeline': [{**event, 'stage': stage_for(event['phase'] + ' ' + event['action'])}
                             for event in timeline],
                'raw_fields': [{'label': headers.get(cell[0], str(cell[0])),
                                'value': cell[1], 'column': cell[0]} for cell in cells],
            }
        details_by_sheet[sheet_code] = sheet_details
    return records, details_by_sheet


def link_candidates(management, ors):
    """Only annotate unique title+date candidates. Never merge or count as verified."""
    m_by = defaultdict(list)
    o_by = defaultdict(list)
    for item in management:
        if item['intake_date']:
            m_by[(norm_title(item['title']), item['intake_date'])].append(item)
    for item in ors:
        if item['is_latest'] and item['scope'] == '산업':
            for day in item['intake_dates_2026']:
                o_by[(norm_title(item['title']), day)].append(item)
    matches = 0
    for key, m_items in m_by.items():
        o_items = o_by.get(key, [])
        if len(m_items) == len(o_items) == 1:
            m_items[0]['match_candidate'] = {'source': 'ORS', 'url': o_items[0]['detail_url']}
            o_items[0]['match_candidate'] = {'source': '관리 Tool', 'url': m_items[0]['detail_url']}
            matches += 1
    return matches


def management_stats(data, management):
    ledger = data['rows']
    yearly = Counter(row['year'] for row in ledger if clean(row.get('year')).isdigit()
                     and int(row['year']) >= 2015)
    statuses = Counter(stage_for(row.get('raw_status')) for row in ledger)
    units = Counter(clean(row.get('unit')) or '분류 미기록' for row in ledger)
    sectors = Counter(clean(row.get('sector')) or '분야 미기록' for row in ledger)
    methods = Counter(clean(row.get('method')) for row in ledger if clean(row.get('method')))
    reviewers = Counter(clean(row.get('reviewer')) or '검토기관 미기록' for row in ledger)
    intake = [record for record in management if record['intake_date'].startswith('2026-')]
    current_statuses = Counter(clean(record['status']) or '상태 미기록' for record in intake)
    current_reviewers = Counter(clean(record['reviewer']) or '검토기관 미기록' for record in intake)
    approved = sum(1 for row in ledger if stage_for(row.get('raw_status')) == '승인·통과')
    return {
        'ledger_rows': len(ledger), 'yearly': [{'year': year, 'count': yearly[str(year)]}
                      for year in range(2015, max([2026]+[int(y) for y in yearly])+1)],
        'management_2026': len(intake), 'approved_rows': approved,
        'current_stage': stage_counts(intake), 'all_stage': stage_counts(
            [{'stage': stage_for(row.get('raw_status'))} for row in ledger]),
        'monthly_2026': month_counts(intake),
        'units': units.most_common(), 'sectors': sectors.most_common(12),
        'methods': methods.most_common(12), 'reviewers': reviewers.most_common(12),
        'statuses_2026': current_statuses.most_common(),
        'reviewers_2026': current_reviewers.most_common(),
        'year_missing_or_early': len(ledger)-sum(yearly.values()),
        'source_files': sorted({clean(row.get('source_file')) for row in ledger}),
    }
