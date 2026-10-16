'use strict';

const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const fmt = value => Number(value).toLocaleString('ko-KR');
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
let summary;

function chart(items, tone = '', unit = '건') {
  const max = Math.max(1, ...items.map(item => item.count));
  return `<div class="barlist">${items.map(item => {
    const width = item.count ? Math.max(1.5, item.count / max * 100) : 0;
    return `<div class="bar ${tone} ${item.count ? '' : 'empty'}" aria-label="${esc(item.label)} ${fmt(item.count)}${unit}"><span class="name" title="${esc(item.label)}">${esc(item.label)}</span><span class="track"><span class="fill" style="--bar:${width.toFixed(2)}%"></span></span><span class="value">${fmt(item.count)}</span></div>`;
  }).join('')}</div>`;
}

function card(label, value, note, primary = false) {
  return `<article class="kpi ${primary ? 'primary' : ''}"><span class="label">${esc(label)}</span><strong>${fmt(value)}</strong><small>${esc(note)}</small></article>`;
}

function panel(title, text, body) {
  return `<section class="panel"><h2>${esc(title)}</h2><p>${esc(text)}</p>${body}</section>`;
}

function renderStatistics() {
  const s = summary;
  const years = s.years.filter(item => item.count > 0);
  return `<div class="kpis">${card('전체 신청 내역', s.total, '관리 도구의 신청 단위', true)}${card('2026년 신청', s.current_year, '신청 연도 기준')}${card('승인 상태', s.approved, '전체 기간 · 현재 단계')}${card('검토·처리 진행', s.active, '전체 기간 · 현재 단계')}</div>
    <h2 class="section-title">신청 흐름과 구성</h2><div class="grid">
      ${panel('신청 연도별 현황', '신청일 연도 기준 · 막대 길이는 최댓값 대비 실제 비율', chart(years, 'navy') + `<div class="key-note">신청 연도가 비어 있는 ${fmt(s.year_missing)}건은 “미기록”으로 표시했습니다.</div>`)}
      ${panel('현재 처리 단계', '각 신청의 최신 분류 · 단계별 합계는 전체 신청과 일치', chart(s.stages))}
    </div><h2 class="section-title">신청 업무와 국내외 구성</h2><div class="grid equal">
      ${panel('업무 유형별 신청', '원본 관리 도구의 업무 구분을 유지', chart(s.units, 'green'))}
      ${panel('국내외 구분', '방법론 신청은 공통으로 분류', chart(s.regions, 'amber'))}
    </div>`;
}

function renderProgress() {
  const s = summary;
  const count = label => s.stages_2026.find(item => item.label === label)?.count ?? 0;
  const ongoing = ['접수', '검토', '협의', '보완', '보류', '철회 요청', '확인 필요'].reduce((sum, label) => sum + count(label), 0);
  const throughMonth = Math.min(12, Number(String(s.source_date).slice(5, 7)) || 12);
  const months = s.months_2026.slice(0, throughMonth).map(item => ({...item, label: `${Number(item.label)}월`}));
  return `<div class="kpis">${card('2026년 신청', s.current_year, '신청 연도 기준', true)}${card('진행 중', ongoing, '접수·검토·협의·보완 등')}${card('승인 상태', count('승인'), '현재 처리 단계')}${card('반려', count('반려'), '현재 처리 단계')}</div>
    <h2 class="section-title">2026년 월별 접수와 단계</h2><div class="grid">
      ${panel('월별 신청 현황', `${esc(s.source_date)} 수신본 · 신청일 월 기준`, chart(months, 'navy') + '<div class="key-note">월별 합계는 2026년 신청 53건과 일치합니다. 접수일이 아닌 신청일 기준입니다.</div>')}
      ${panel('현재 처리 단계별 신청', '2026년 신청 53건의 현재 상태', chart(s.stages_2026.filter(item => item.count > 0)))}
    </div><h2 class="section-title">담당자 확인을 위한 해석</h2><div class="grid equal">
      ${panel('우선 확인할 단계', '이 화면에는 신청별 의견·다음 조치가 포함되지 않습니다.', `<div class="steps"><div class="step"><span>보완</span><strong>${fmt(count('보완'))}</strong></div><div class="step"><span>검토</span><strong>${fmt(count('검토'))}</strong></div><div class="step"><span>접수</span><strong>${fmt(count('접수'))}</strong></div><div class="step"><span>보류·협의</span><strong>${fmt(count('보류') + count('협의'))}</strong></div></div><div class="key-note">담당자 배정, 안내 사항, 위험 판단은 로컬 v19 작업본의 개별 신청에서 확인해야 합니다.</div>`)}
      ${panel('월별 숫자 한눈에 보기', '0건도 표시하며, 막대 그래프와 같은 집계입니다.', `<div class="month-grid">${months.map(item => `<div class="month"><small>${esc(item.label)}</small><strong>${fmt(item.count)}</strong></div>`).join('')}</div>`)}
    </div>`;
}

function renderAbout() {
  const s = summary;
  return `<div class="about-list">
    <div><strong>자료 기준</strong><p>관리 도구 수신일 ${esc(s.source_date)}, 분석 기준일 ${esc(s.as_of)}. 총 ${fmt(s.total)}건의 신청을 원본의 신청 단위로 집계했습니다. 사업별 묶음 수와 혼동하지 마세요.</p></div>
    <div><strong>공개 범위</strong><p>연도·월·업무·처리 단계·국내외별 건수만 제공합니다. 사업명, 신청기관, 담당자, 검토 의견, 사무국 안내, 연락처, 원본 파일은 이 공개판에 없습니다.</p></div>
    <div><strong>업데이트 방식</strong><p>이 화면은 고정된 검증 집계입니다. GitHub의 incoming 폴더에 ORS ZIP을 올려도 v19 공개 통계가 즉시 바뀌지 않습니다. 원본 검증과 공개용 집계 재생성이 필요합니다.</p></div>
    <div><strong>담당자용 기능</strong><p>신청 등록·수정, 사업별 상세, 검토 의견·안내, Excel 가져오기·내보내기 및 변경 이력은 로컬 서버가 실행되는 v19 작업본에서만 제공됩니다.</p></div>
  </div>`;
}

function render() {
  if (!summary) return;
  const view = ['statistics', 'progress', 'about'].includes(location.hash.slice(1)) ? location.hash.slice(1) : 'statistics';
  const labels = {
    statistics: ['MANAGEMENT STATISTICS', '신청 통계', '신청관리 도구의 전체 흐름을 집계로 살펴봅니다.'],
    progress: ['2026 APPLICATION PROGRESS', '2026 진행 현황', '올해 접수된 신청의 월별 규모와 현재 처리 단계를 봅니다.'],
    about: ['DATA & ACCESS', '자료 기준과 이용', '공개 범위와 담당자용 작업본의 차이를 확인합니다.'],
  };
  const [eyebrow, title, description] = labels[view];
  $('#eyebrow').textContent = eyebrow;
  $('#title').textContent = title;
  $('#description').textContent = description;
  $('#asof').textContent = `관리 도구 ${summary.source_date} · 분석 ${summary.as_of}`;
  $$('[data-view]').forEach(link => {
    link.classList.toggle('active', link.dataset.view === view);
    if (link.dataset.view === view) link.setAttribute('aria-current', 'page');
    else link.removeAttribute('aria-current');
  });
  $('#content').innerHTML = view === 'statistics' ? renderStatistics() : view === 'progress' ? renderProgress() : renderAbout();
}

window.addEventListener('hashchange', render);
fetch('./summary.json', {cache: 'no-store'}).then(response => {
  if (!response.ok) throw new Error('집계 파일을 찾을 수 없습니다.');
  return response.json();
}).then(data => {summary = data; render();}).catch(error => {
  $('#content').innerHTML = `<div class="error">집계를 불러오지 못했습니다. ${esc(error.message)}</div>`;
});
