"""Traceable operator worklist. Suggestions are never presented as assigned tasks."""
from __future__ import annotations

import re
from collections import Counter
from datetime import date

ACTIVE = {'접수', '심사', '의견 수렴', '보완', '확인 필요'}
SUGGESTIONS = {
    '접수': '접수 사실·신청 회차·담당 배정 확인',
    '심사': '검토기관 의견과 심사 일정 확인',
    '의견 수렴': '회신기관 의견·다음 통보 확인',
    '보완': '보완 요청 내용·회신 예정일·신청자 안내 확인',
    '승인·통과': '승인 근거와 종결 처리 확인',
    '반려': '반려 사유·신청자 안내 확인',
    '후속 관리': '착수·운영 후속 조치 확인',
    '확인 필요': '원본 처리상태·신청 회차 확인',
}


def text(value):
    return '' if value is None else str(value).strip()


def field(detail, *names):
    for item in detail.get('raw_fields', []):
        if text(item.get('label')) in names and text(item.get('value')):
            return text(item['value'])
    return ''


def round_label(record, detail):
    explicit = text(detail.get('supplement_round'))
    if explicit:
        return explicit, '관리 Tool 원본'
    mentions = re.findall(r'보완\s*([1-9])\s*차', text(record.get('status')))
    if mentions:
        return f'{max(map(int, mentions))}차 표기', 'ORS 상태 원문'
    return '', '원본 미기록'


def age(last_date, as_of):
    try:
        days = (date.fromisoformat(as_of[:10]) - date.fromisoformat(last_date[:10])).days
        return days if days >= 0 else None
    except (ValueError, TypeError):
        return None


def work_item(record, detail, note, as_of):
    management = record['source'] == '관리 Tool'
    stage = record['stage']
    age_days = age(record.get('last_date') or record.get('intake_date'), as_of)
    flags = record.get('risk') or []
    manual_action = text(note.get('next_action')) if management else ''
    action = manual_action or text(record.get('next_action')) or SUGGESTIONS.get(stage, SUGGESTIONS['확인 필요'])
    action_origin = '담당자 입력' if manual_action else '원본 기반 내부 제안'
    owner = text(note.get('owner')) if management else ''
    source_owner = text(detail.get('owner')) if management else ''
    if not owner and source_owner:
        owner = source_owner
        owner_origin = '관리 Tool 원본'
    else:
        owner_origin = '담당자 입력' if owner else '담당자 미기록'
    if any(x.get('level') == '조치' for x in flags) and stage in ACTIVE:
        priority = '조치 확인'
    elif age_days is not None and age_days >= 30 and stage in ACTIVE:
        priority = '장기 계류 점검'
    else:
        priority = '일반 확인'
    original_issues = [text(x.get('label')) for x in flags if x.get('origin') == '원본 기록']
    notes = text(detail.get('notes')) if management else ''
    special = notes or (' / '.join(original_issues) if original_issues else
                        ' / '.join(text(x.get('label')) for x in flags[:2]))
    special_origin = ('관리 Tool 원본' if notes else '원본 기록' if original_issues
                      else '내부 확인 제안' if flags else '표시 없음')
    supplement, supplement_origin = round_label(record, detail)
    reviewer = text(record.get('reviewer')) if management else field(detail, '검토기관', '평가기관')
    authority = field(detail, '관장기관') if not management else ''
    first_reviewer = text(detail.get('first_agency')) if management else field(detail, '최초 검토기관')
    visit = text(detail.get('visit')) if management else field(detail, '현장방문 여부', '현장 방문 여부')
    return {
        'id': record['id'], 'source': record['source'], 'title': record['title'],
        'applicant': record.get('applicant', ''), 'kind': record.get('kind', ''),
        'stage': stage, 'source_status': record.get('status', ''),
        'intake_date': record.get('intake_date', ''), 'last_date': record.get('last_date', ''),
        'age_days': age_days, 'internal_30day_check': bool(age_days is not None and age_days >= 30 and stage in ACTIVE),
        'priority': priority, 'next_action': action, 'action_origin': action_origin,
        'owner': owner, 'owner_origin': owner_origin,
        'review_org': reviewer, 'authority_org': authority,
        'first_review_org': first_reviewer, 'site_visit': visit,
        'supplement_round': supplement, 'supplement_origin': supplement_origin,
        'special': special, 'special_origin': special_origin, 'risk': flags,
        'intake_status': text(note.get('intake_status')) if management else '',
        'opinion_status': text(note.get('opinion_status')) if management else '',
        'opinion_text': text(note.get('opinion_text')) if management else '',
        'guidance_status': text(note.get('guidance_status')) if management else '',
        'guidance_text': text(note.get('guidance_text')) if management else '',
        'workflow_updated_at': text(note.get('updated_at')) if management else '',
        'due_basis': '법정·계약 기한 기준 미설정',
        'detail_url': record['detail_url'],
        'source_file': record.get('source_file', ''),
        'source_sheet': record.get('source_sheet', ''),
        'source_row': record.get('source_row'),
        'raw_excel_sheet': record.get('raw_excel_sheet', ''),
    }


def make_worklist(progress, management_details, ors_details, manual_notes):
    as_of = progress['as_of']
    rows = []
    for record in progress['management']:
        key = record['id']
        rows.append(work_item(record, management_details[key], manual_notes.get(key, {}), as_of))
    for record in progress['ors']:
        if not (record.get('is_latest') and record.get('scope') == '산업' and record.get('event_count_2026', 0) > 0):
            continue
        detail = ors_details[record['raw_excel_sheet']][str(record['source_row'])]
        rows.append(work_item(record, detail, {}, as_of))
    rows.sort(key=lambda x: (x['last_date'] or x['intake_date'] or '', x['id']), reverse=True)
    management = [x for x in rows if x['source'] == '관리 Tool']
    ors = [x for x in rows if x['source'] == 'ORS']
    summary = {
        'source_rows': len(rows), 'management_rows': len(management), 'ors_activity_rows': len(ors),
        'priority': dict(Counter(x['priority'] for x in rows)),
        'stages': dict(Counter(x['stage'] for x in rows)),
        'owner_missing': sum(not x['owner'] for x in rows),
        'manual_actions': sum(x['action_origin'] == '담당자 입력' for x in rows),
        'manual_opinions': sum(bool(x['opinion_text']) for x in management),
        'manual_guidance': sum(bool(x['guidance_text']) for x in management),
        'internal_30day_check': sum(x['internal_30day_check'] for x in rows),
        'supplement': sum(x['stage'] == '보완' for x in rows),
        'management_review_org': dict(Counter(x['review_org'] or '검토기관 미기록' for x in management).most_common()),
    }
    return {'as_of': as_of, 'freshness': progress['freshness'],
            'summary': summary, 'rows': rows,
            'interpretation': '출처별 원본행 대기열. 동일 신청의 관리 Tool·ORS 중복 가능성을 제거한 고유 건수 아님.'}
