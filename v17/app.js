const $ = selector => document.querySelector(selector);
const fmt = value => Number(value || 0).toLocaleString('ko-KR');
const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
const show = value => typeof value === 'object' && value !== null ? JSON.stringify(value, null, 2) : String(value ?? '');
const has = value => value !== undefined && value !== null && String(value).trim() !== '';
async function getJson(path, fresh=false) {
  const response = await fetch(path + (fresh ? `?v=${Date.now()}` : ''), {cache:fresh?'no-store':'default'});
  if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`);
  return response.json();
}
function pill(stage) {
  const color = ({'접수':'blue','심사':'blue','의견 수렴':'teal','보완':'amber','승인·통과':'green','반려':'red','후속 관리':'teal'})[stage] || '';
  return `<span class="pill ${color}">${esc(stage || '확인 필요')}</span>`;
}
function kpi(label,value,note,accent=false) {
  return `<article class="kpi ${accent?'accent':''}"><span class="label">${esc(label)}</span><strong>${fmt(value)}</strong><small>${esc(note)}</small></article>`;
}
function enhanceNav(){
  const nav=$('.nav');if(!nav)return;
  const excel=nav.querySelector('a[href$=".xlsx"]');
  if(excel)excel.href='외부사업_통합관리_v17.xlsx';
  if(!nav.querySelector('a[href="work.html"]')){
    const link=document.createElement('a');link.href='work.html';link.innerHTML='☷ &nbsp;담당자 작업현황';
    nav.insertBefore(link,excel||null);
  }
}
function bars(pairs,{color='',limit=12}={}) {
  const data=pairs.slice(0,limit),max=Math.max(1,...data.map(item=>Number(item[1]||0)));
  return data.length ? data.map(([label,value])=>`<div class="barrow"><span class="name" title="${esc(label)}">${esc(label)}</span><span class="track"><i class="fill ${color}" style="width:${Math.max(1,100*Number(value||0)/max)}%"></i></span><b>${fmt(value)}</b></div>`).join('') : '<div class="empty">표시할 자료가 없습니다.</div>';
}
function monthColumns(left,right=null) {
  const l=left.map(x=>x.count),r=right?right.map(x=>x.count):[],max=Math.max(1,...l,...r);
  return left.map((x,i)=>`<div class="month" title="${esc(x.month)} · ${fmt(x.count)}${right?` / ${fmt(right[i].count)}`:''}"><div class="columns"><i class="blue" style="height:${Math.max(2,140*x.count/max)}px"></i>${right?`<i style="height:${Math.max(2,140*right[i].count/max)}px"></i>`:''}</div><span>${i+1}월</span></div>`).join('');
}
function stagePairs(items){return items.map(x=>[x.stage,x.count]);}
function toast(message,auto=false){const box=$('#update-toast');if(!box)return;box.className='toast visible';box.innerHTML=`${esc(message)} <button type="button">새로고침</button>`;box.querySelector('button').onclick=()=>location.reload();if(auto)setTimeout(()=>location.reload(),8000);}
async function liveCheck(page){
  try {
    const initial=await getJson('manifest.json',true);
    const dates=initial.freshness||{};
    document.querySelectorAll('.asof').forEach(node=>node.textContent=`관리 Tool ${dates.management_date||'확인 필요'} · ORS ${dates.ors_date||'확인 필요'}`);
    setInterval(async()=>{try{const current=await getJson('manifest.json',true);if(current.version!==initial.version){toast('새 ORS 판본이 게시되었습니다.',page!=='case');}}catch(error){}},30000);
  } catch(error){const stamp=$('.stamp');if(stamp){stamp.classList.add('warn');stamp.textContent='갱신 상태 확인 불가';}}
}
function renderStatistics(data){
  $('#stats-kpis').innerHTML=kpi('관리 Tool 전체 원장행',data.ledger_rows,'고유 사업 수와 다름',true)+kpi('2026 신청일 확인',data.management_2026,'관리 Tool 원장행')+kpi('승인 상태 원장행',data.approved_rows,'원본 처리상태 기준')+kpi('연도 미확인·2014 이전',data.year_missing_or_early,'연도별 추이에서 제외');
  $('#year-bars').innerHTML=bars(data.yearly.map(x=>[x.year,x.count]),{color:'blue',limit:30});
  $('#all-stage-bars').innerHTML=bars(stagePairs(data.all_stage));
  $('#monthly-chart').innerHTML=monthColumns(data.monthly_2026);
  $('#current-stage-bars').innerHTML=bars(stagePairs(data.current_stage),{color:'amber'});
  $('#unit-bars').innerHTML=bars(data.units,{limit:8});
  $('#sector-bars').innerHTML=bars(data.sectors,{limit:8,color:'blue'});
  $('#reviewer-bars').innerHTML=bars(data.reviewers,{limit:8});
  $('#status-2026-bars').innerHTML=bars(data.statuses_2026,{limit:12,color:'amber'});
  $('#reviewer-2026-bars').innerHTML=bars(data.reviewers_2026,{limit:12,color:'blue'});
  $('#method-bars').innerHTML=bars(data.methods,{limit:12});
  $('#year-table').innerHTML=data.yearly.map(x=>`<tr><td>${esc(x.year)}</td><td><b>${fmt(x.count)}</b></td><td>신청 연도가 확인된 관리 Tool 원장행</td></tr>`).join('');
  $('#stats-method').innerHTML=`통계의 주원천은 <b>${esc(data.source)}</b>이며, 총 ${fmt(data.ledger_rows)}행입니다. 관리 Tool 기준일은 ${esc(data.freshness?.management_date)}, ORS 수신본 기준일은 ${esc(data.freshness?.ors_date)}입니다. ORS 최신 산업부문은 2026 접수 근거 ${fmt(data.ors_reference.industry_intake_2026)}행·2026 활동 ${fmt(data.ors_reference.industry_activity_2026)}행으로 별도 출처입니다. 신청 회차가 확정 연결되지 않아 두 원장을 더하지 않습니다.`;
}
function currentOrs(data){return data.ors.filter(row=>row.is_latest && row.scope==='산업' && row.event_count_2026>0);}
function listFor(view,data){
  if(view==='management')return data.management;
  if(view==='intake')return data.ors.filter(row=>row.is_latest&&row.scope==='산업'&&row.intake_dates_2026.length);
  if(view==='activity')return currentOrs(data);
  if(view==='orsall')return data.ors;
  return [...data.management,...currentOrs(data)];
}
function attentionCounts(rows){
  const counts={};for(const row of rows)for(const flag of row.risk||[])counts[flag.label]=(counts[flag.label]||0)+1;
  return Object.entries(counts).sort((a,b)=>b[1]-a[1]);
}
function queue(rows,stages){
  return rows.filter(row=>stages.includes(row.stage)).sort((a,b)=>(b.last_date||'').localeCompare(a.last_date||'')).slice(0,7).map(row=>`<a class="queue-item" href="${esc(row.detail_url)}">${pill(row.stage)}<span><strong>${esc(row.title||'사업명 미기록')}</strong><small>${esc(row.risk?.[0]?.label||row.status||'원본 상태 확인')} · ${esc(row.source)}</small></span><time>${esc(row.last_date||row.intake_date||'')}</time></a>`).join('')||'<div class="empty">해당 단계의 기록이 없습니다.</div>';
}
function csvDownload(filename,rows){
  const safe=value=>{let text=String(value??'');if(/^[=+@\-\t\r]/.test(text))text="'"+text;return '"'+text.replaceAll('"','""')+'"'};
  const blob=new Blob(['\uFEFF'+rows.map(row=>row.map(safe).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'});
  const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=filename;link.click();setTimeout(()=>URL.revokeObjectURL(link.href),1500);
}
function renderProgress(data){
  const overview=data.overview;
  $('#management-stage-count').textContent=fmt(overview.management_2026);
  $('#ors-stage-count').textContent=fmt(overview.ors_industry_2026_activity);
  $('#progress-kpis').innerHTML=kpi('관리 Tool 2026 접수',overview.management_2026,'의견·안내 업무 원장',true)+kpi('ORS 2026 신규 접수',overview.ors_industry_2026_intake,'산업부문 최신 · 접수 근거')+kpi('ORS 2026 처리 이력',overview.ors_industry_2026_activity,'보완·승인·반려 포함')+kpi('전체 ORS 원본행',overview.ors_all_occurrences,'모든 수신본 보존');
  $('#intake-month-chart').innerHTML=monthColumns(overview.management_monthly,overview.ors_monthly);
  const active=[...data.management,...currentOrs(data)];
  $('#attention-summary').innerHTML=bars(attentionCounts(active).slice(0,7),{color:'amber'});
  $('#management-stages').innerHTML=bars(stagePairs(overview.management_stages),{color:'blue'});
  $('#ors-stages').innerHTML=bars(stagePairs(overview.ors_stages));
  $('#queue-supplement').innerHTML=queue(active,['보완']);
  $('#queue-opinion').innerHTML=queue(active,['의견 수렴','심사']);
  const stageSelect=$('#stage');stageSelect.insertAdjacentHTML('beforeend',[...new Set(active.map(x=>x.stage))].map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join(''));
  const monthSelect=$('#month');monthSelect.insertAdjacentHTML('beforeend',Array.from({length:12},(_,i)=>`<option value="2026-${String(i+1).padStart(2,'0')}">${i+1}월</option>`).join(''));
  let page=0,filtered=[];
  function filter(){
    const view=$('#view').value,stage=$('#stage').value,month=$('#month').value,attention=$('#attention').value,query=$('#search').value.trim().toLowerCase();
    filtered=listFor(view,data).filter(row=>(stage==='all'||row.stage===stage)&&(month==='all'||(row.intake_date||'').startsWith(month))&&(attention==='all'||(attention==='yes')===(row.risk?.length>0))&&(!query||[row.title,row.applicant,row.agency,row.reviewer,row.status,row.source_file,row.kind,row.opinion,row.guidance,row.next_action,...(row.risk||[]).map(x=>x.label)].some(value=>String(value||'').toLowerCase().includes(query))));
    filtered.sort((a,b)=>(b.last_date||b.intake_date||'').localeCompare(a.last_date||a.intake_date||''));
  }
  function draw(reset=false){
    if(reset)page=0;filter();const pages=Math.max(1,Math.ceil(filtered.length/30));page=Math.max(0,Math.min(page,pages-1));
    $('#work-note').textContent=`${fmt(filtered.length)}행 · 출처별 원본행`;
    $('#pages').textContent=`${page+1} / ${pages}`;$('#prev').disabled=page===0;$('#next').disabled=page>=pages-1;
    const rows=filtered.slice(page*30,page*30+30);
    $('#work-rows').innerHTML=rows.length?rows.map(row=>`<tr tabindex="0" role="link" data-href="${esc(row.detail_url)}"><td class="nowrap"><b>${esc(row.intake_date||'접수일 미확인')}</b><br><span class="tiny">최근 ${esc(row.last_date||'미기록')}</span></td><td class="rowtitle"><a href="${esc(row.detail_url)}">${esc(row.title||'사업명 미기록')}</a><small>${esc(row.source)} · ${esc(row.kind)} · ${esc(row.applicant||'신청자 미기록')}</small></td><td>${pill(row.stage)}<br><span class="tiny">${esc(String(row.status||'').split('\n')[0].slice(0,35))}</span></td><td>${esc(row.agency||row.reviewer||'원본 미기록')}</td><td>${row.risk?.length?row.risk.slice(0,2).map(flag=>`<span class="pill ${flag.level==='조치'?'amber':''}">${esc(flag.label)}</span>`).join(' '):'<span class="tiny">표시 없음</span>'}${row.guidance?`<small>${esc(row.guidance.slice(0,70))}</small>`:''}</td><td><a class="button small ghost" href="${esc(row.detail_url)}">상세 →</a></td></tr>`).join(''):'<tr><td colspan="6" class="empty">조건에 맞는 행이 없습니다.</td></tr>';
    sessionStorage.setItem('koc-v15-filters',JSON.stringify({view:$('#view').value,stage:$('#stage').value,month:$('#month').value,attention:$('#attention').value,query:$('#search').value}));
  }
  const saved=JSON.parse(sessionStorage.getItem('koc-v15-filters')||'{}');
  for(const [id,value] of Object.entries(saved))if($('#'+id)&&value!==undefined)$('#'+id).value=value;
  for(const id of ['view','stage','month','attention'])$('#'+id).addEventListener('change',()=>draw(true));
  $('#search').addEventListener('input',()=>draw(true));
  $('#prev').onclick=()=>{page--;draw()};$('#next').onclick=()=>{page++;draw()};
  $('#work-rows').addEventListener('click',event=>{const row=event.target.closest('tr[data-href]');if(row&&!event.target.closest('a'))location.href=row.dataset.href;});
  $('#work-rows').addEventListener('keydown',event=>{if(event.key==='Enter'&&event.target.matches('tr[data-href]'))location.href=event.target.dataset.href;});
  $('#export').onclick=()=>csvDownload('외부사업_진행현황_v17.csv',[
      ['출처','접수일','최근이력일','사업명','신청자','유형','현재단계','원본상태','기관','확인사항','안내내용','원본파일','원본시트','원본행'],
      ...filtered.map(row=>[row.source,row.intake_date,row.last_date,row.title,row.applicant,row.kind,row.stage,row.status,row.agency,(row.risk||[]).map(x=>x.label).join(' / '),row.guidance,row.source_file,row.source_sheet,row.source_row])]);
  draw(true);
}
function fact(label,value){return `<div class="fact"><span>${esc(label)}</span><strong>${has(value)?esc(show(value)):'원본 미기록'}</strong></div>`;}
function callout(title,text,type=''){return `<div class="callout ${type}"><b>${esc(title)}</b><p>${esc(text||'원문 확인 필요')}</p></div>`;}
function rawGrid(fields){return fields.filter(item=>has(item.value)).map(item=>`<div class="rawfield"><span>${esc(item.label)}${item.column?` · ${esc(item.column)}`:''}</span><strong>${esc(show(item.value))}</strong></div>`).join('')||'<div class="empty">기록된 필드가 없습니다.</div>';}
const stageMeaning={'접수':'신청·접수로 표시된 단계입니다. 접수 확정 여부와 기산일은 원본에서 확인합니다.','심사':'검토·평가가 진행 또는 필요하다고 표시된 단계입니다. 결과는 확정되지 않았습니다.','의견 수렴':'협의·의견 처리 상태입니다. 회신기관과 의견 내용은 별도 원문을 확인해야 합니다.','보완':'추가 자료나 재제출 관련 상태입니다. 요청 사유·회신 내용·안내 여부를 확인합니다.','승인·통과':'원본 처리상태에 승인이 표시됩니다. 승인서의 효력·범위는 원문 확인이 필요합니다.','반려':'원본 처리상태에 반려가 표시됩니다. 사유와 후속 신청 가능성은 원문 확인이 필요합니다.','후속 관리':'착수·운영 등 승인 이후의 후속 상태일 수 있습니다. 해당 신청의 최종 결과와 구분합니다.','확인 필요':'원본 처리상태가 표준 단계에 연결되지 않았습니다.'};
function renderCase(item){
  document.title=`${item.title||'신청 상세'} | 외부사업`;
  const isOrs=item.source==='ORS';
  const flags=item.risk||[];
  const timeline=isOrs?(item.timeline||[]):[];
  const sources=`${item.source_file||''} · ${item.source_sheet||''} · ${item.source_row||''}행`;
  $('#case-content').innerHTML=`<header class="detail-hero"><div class="eyebrow">${esc(item.source)} / APPLICATION DETAIL</div><h1>${esc(item.title||'사업명 미기록')}</h1><p class="subtitle">${esc(sources)}</p><div class="detail-meta">${pill(item.stage)}<span class="pill">${esc(item.kind||'유형 미기록')}</span><span class="pill">${esc(item.intake_date||'접수일 확인 필요')}</span></div></header>
    <div class="grid wide-left"><section class="panel"><div class="panel-head"><div><h2>무엇을 신청했나</h2><p>원본 필드명과 값을 그대로 표시합니다.</p></div></div><div class="facts">${fact('신청자/기업',item.applicant)}${fact('신청 유형',item.application_type||item.kind)}${fact('사업·프로젝트 번호',item.number||item.source_key)}${fact('사업 분야',item.sector)}${fact('적용방법론',item.method)}${fact('사업 장소',item.location||item.region)}${fact('연평균 감축량(원본 표기)',item.expected_reduction)}${fact('접수일',item.intake_date)}${fact('최근 이력일',item.last_date)}</div></section><section class="panel"><div class="panel-head"><div><h2>현재 단계와 담당 기관</h2><p>기관의 역할은 원본 표기를 우선합니다.</p></div></div>${callout('원본 처리상태',item.status||'미기록')}${callout('단계의 업무 의미 · 분류 설명',stageMeaning[item.stage]||stageMeaning['확인 필요'])}${callout(isOrs?'관장기관 (원본)':'원본 배정·검토기관',item.agency||item.reviewer||'원본 미기록')}${isOrs?'<p class="metric-note">관장기관은 소관 기관이며 검토 수행기관 또는 최초 검토기관이라고 단정하지 않습니다.</p>':''}${item.first_agency&&item.first_agency!==item.agency?callout('최초 검토기관 기록',item.first_agency):''}${item.agency_events?.length?item.agency_events.map(event=>callout('기관 지정 이력',event.text)).join(''):''}${item.match_candidate?`<div class="notice">같은 사업명·접수일의 1:1 <a href="${esc(item.match_candidate.url)}">${esc(item.match_candidate.source)} 대조 후보</a>가 있습니다. 신청 회차 동일성은 검증되지 않았습니다.</div>`:''}</section></div>
    <div class="section-title"><h2>특이사항·의견·사무국 안내</h2><p>원본 기록과 내부 확인 제안을 구분합니다.</p></div><div class="grid"><section class="panel"><h2>확인할 사항</h2>${flags.length?flags.map(flag=>callout(`${flag.label} · ${flag.origin}`,flag.basis,flag.level==='조치'?'amber':flag.level==='점검'?'red':'')).join(''):'<div class="notice">현재 자동으로 표시할 확인사항이 없습니다. 누락된 원본 기록이 없다는 의미는 아닙니다.</div>'}</section><section class="panel"><h2>검토 의견·안내 기록</h2>${callout('의견 검토',item.opinion||'원본 또는 실무 입력에 기록되지 않음')}${callout('사무국 안내',item.guidance||'원본 또는 실무 입력에 기록되지 않음')}${callout('다음 조치',item.next_action||'확인 필요')}${item.notes?callout('관리 Tool 특이사항',item.notes):''}</section></div>
    <div class="section-title"><h2>신청에서 현재까지</h2><p>${isOrs?'ORS 원본 이력의 날짜·문구':'관리 Tool에 기록된 처리·배정 내역'}</p></div><div class="grid wide-left"><section class="panel"><h2>처리 타임라인</h2>${timeline.length?`<div class="timeline">${timeline.map(event=>`<div class="event"><time>${esc(event.date)}</time><b>${esc(event.phase||event.action)} ${pill(event.stage)}</b><p>${esc(event.text)}</p></div>`).join('')}</div>`:item.history?`<div class="callout"><p style="white-space:pre-wrap">${esc(item.history)}</p></div>`:'<div class="notice">원본에 날짜별 처리 이력이 없습니다.</div>'}</section><section class="panel"><h2>기관·현장·출처</h2>${fact(isOrs?'원본 관장기관':'현재 검토기관',isOrs?item.agency:item.reviewer)}${isOrs?fact('검토기관(별도 원본 칸)',item.reviewer):''}${fact('최초 검토기관',item.first_agency)}${fact('기관 배정 이력',item.assignments)}${fact('현장 방문',item.visit)}${fact('보완 차수',item.supplement_round)}${fact('원본 ZIP',item.source_zip)}${fact('원본 Excel 시트',item.raw_excel_sheet||item.source_sheet)}<p class="metric-note">기관명·역할이 원본에서 비어 있으면 임의로 채우지 않습니다. 모든 원본 셀은 연결된 Excel 원본 시트에서 확인할 수 있습니다.</p></section></div>
    <div class="section-title"><h2>원본 필드 전체</h2><p>원본의 저장값을 추적용으로 표시합니다.</p></div><section class="panel"><div class="rawgrid">${rawGrid(item.raw_fields||[])}</div></section>`;
}
function priorityPill(value){
  const color=value==='조치 확인'?'priority-high':value==='장기 계류 점검'?'priority-age':'priority-normal';
  return `<span class="pill ${color}">${esc(value)}</span>`;
}
function renderWork(data){
  const summary=data.summary,all=data.rows;
  $('#work-kpis').innerHTML=kpi('출처별 2026 업무행',summary.source_rows,'관리 Tool·ORS 중복 가능',true)+kpi('조치 확인',summary.priority['조치 확인']||0,'원본·안내 기록 점검')+kpi('내부 30일 점검',summary.internal_30day_check,'법정 기한 아님')+kpi('담당자 미기록',summary.owner_missing,'기관과 담당자 구분');
  $('#work-stage-bars').innerHTML=bars(Object.entries(summary.stages),{color:'blue'});
  $('#work-priority-bars').innerHTML=bars(Object.entries(summary.priority),{color:'amber'});
  $('#work-reviewer-bars').innerHTML=bars(Object.entries(summary.management_review_org),{color:'blue'});
  $('#work-priority').insertAdjacentHTML('beforeend',Object.keys(summary.priority).map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join(''));
  $('#work-stage').insertAdjacentHTML('beforeend',Object.keys(summary.stages).map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join(''));
  let page=0,filtered=[];
  const rank={'조치 확인':0,'장기 계류 점검':1,'일반 확인':2};
  function draw(reset=false){
    if(reset)page=0;
    const source=$('#work-source').value,priority=$('#work-priority').value,stage=$('#work-stage').value,owner=$('#work-owner').value,query=$('#work-search').value.trim().toLowerCase();
    filtered=all.filter(row=>(source==='all'||row.source===source)&&(priority==='all'||row.priority===priority)&&(stage==='all'||row.stage===stage)&&(owner==='all'||(owner==='missing')===!row.owner)&&(!query||[row.title,row.applicant,row.owner,row.review_org,row.authority_org,row.next_action,row.special,row.source_status,row.supplement_round].some(x=>String(x||'').toLowerCase().includes(query))));
    filtered.sort((a,b)=>(rank[a.priority]??3)-(rank[b.priority]??3)||(b.age_days??-1)-(a.age_days??-1)||(b.last_date||'').localeCompare(a.last_date||''));
    const pages=Math.max(1,Math.ceil(filtered.length/25));page=Math.max(0,Math.min(page,pages-1));
    $('#work-count').textContent=`${fmt(filtered.length)}행 · 출처별 원본행`;
    $('#work-pages').textContent=`${page+1} / ${pages}`;$('#work-prev').disabled=page===0;$('#work-next').disabled=page>=pages-1;
    const rows=filtered.slice(page*25,page*25+25);
    $('#work-table').innerHTML=rows.length?rows.map(row=>`<tr tabindex="0" role="link" data-href="${esc(row.detail_url)}"><td>${priorityPill(row.priority)}</td><td class="rowtitle"><a href="${esc(row.detail_url)}">${esc(row.title||'사업명 미기록')}</a><small>${esc(row.source)} · ${esc(row.stage)} · ${esc(row.applicant||'신청자 미기록')}</small></td><td class="nowrap">${esc(row.last_date||'미기록')}<small>${row.age_days===null?'경과일 미산출':`${fmt(row.age_days)}일 경과`}</small></td><td>${esc(row.owner||'담당 미기록')}<small>${esc(row.review_org||row.authority_org||'기관 미기록')} · ${esc(row.owner_origin)}</small></td><td class="work-action">${esc(row.next_action)}<small>${esc(row.action_origin)} · ${esc(row.due_basis)}</small></td><td>${esc((row.special||'표시 없음').slice(0,100))}<small>${esc(row.special_origin)} · ${esc(row.supplement_round||'보완차수 미기록')}</small></td><td><a class="button small ghost" href="${esc(row.detail_url)}">상세 →</a></td></tr>`).join(''):'<tr><td colspan="7" class="empty">조건에 맞는 행이 없습니다.</td></tr>';
  }
  for(const id of ['work-source','work-priority','work-stage','work-owner'])$('#'+id).addEventListener('change',()=>draw(true));
  $('#work-search').addEventListener('input',()=>draw(true));
  $('#work-prev').onclick=()=>{page--;draw()};$('#work-next').onclick=()=>{page++;draw()};
  $('#work-table').addEventListener('click',event=>{const row=event.target.closest('tr[data-href]');if(row&&!event.target.closest('a'))location.href=row.dataset.href;});
  $('#work-table').addEventListener('keydown',event=>{if(event.key==='Enter'&&event.target.matches('tr[data-href]'))location.href=event.target.dataset.href;});
  $('#work-export').onclick=()=>csvDownload('외부사업_담당자_작업현황_v17.csv',[
    ['출처','사업·신청명','신청자','현재 단계','원본 상태','우선순위','최근 이력일','경과일','보완차수','담당자','검토기관','관장기관','다음 조치','조치 출처','특이사항','특이사항 근거','의견 상태','안내 상태','현장 방문','기한 기준','원본 파일','원본 시트','원본 행'],
    ...filtered.map(x=>[x.source,x.title,x.applicant,x.stage,x.source_status,x.priority,x.last_date,x.age_days,x.supplement_round,x.owner,x.review_org,x.authority_org,x.next_action,x.action_origin,x.special,x.special_origin,x.opinion_status,x.guidance_status,x.site_visit,x.due_basis,x.source_file,x.source_sheet,x.source_row])]);
  draw(true);
}
async function renderWorkCase(item){
  const data=await getJson('data/work.json');
  const task=data.rows.find(row=>row.id===item.id&&row.source===item.source);
  if(!task)return;
  $('#case-content').insertAdjacentHTML('beforeend',`<div class="section-title detail-work"><h2>담당자 작업현황·다음 단계</h2><p>업무 입력과 원본 기반 확인 제안을 구분합니다.</p></div><section class="panel"><div class="facts">${fact('확인 우선순위',task.priority)}${fact('업무 담당자',task.owner||'담당자 미기록')}${fact('담당 근거',task.owner_origin)}${fact('최근 이력 이후 경과일',task.age_days===null?'미산출':`${task.age_days}일`)}${fact('보완 차수',task.supplement_round)}${fact('보완 차수 근거',task.supplement_origin)}${fact('법정·계약 기한',task.due_basis)}</div>${callout(`다음 조치 · ${task.action_origin}`,task.next_action)}${task.special?callout(`특이사항·확인사항 · ${task.special_origin}`,task.special):''}${task.workflow_updated_at?`<p class="metric-note">실무 입력 수정일: ${esc(task.workflow_updated_at)}</p>`:''}</section>`);
}
async function loadCase(){
  const p=new URLSearchParams(location.search),source=p.get('source');let item;
  if(source==='management'&&/^[0-9a-f]{20,40}$/.test(p.get('id')||'')){
    const details=await getJson('details/management.json');item=details[p.get('id')];
  }else if(source==='ors'&&/^ORS\d{2,3}$/.test(p.get('sheet')||'')&&/^\d+$/.test(p.get('row')||'')){
    const details=await getJson(`details/${p.get('sheet')}.json`);item=details[p.get('row')];
  }
  if(!item)throw new Error('해당 원본행을 찾지 못했습니다. 진행현황에서 다시 선택하세요.');
  renderCase(item);
  await renderWorkCase(item);
}
async function main(){
  const page=document.body.dataset.page;enhanceNav();liveCheck(page);
  try{
    if(page==='statistics')renderStatistics(await getJson('data/statistics.json'));
    else if(page==='progress')renderProgress(await getJson('data/progress.json'));
    else if(page==='work')renderWork(await getJson('data/work.json'));
    else if(page==='case')await loadCase();
  }catch(error){const root=page==='case'?'#case-content':'.main';$(root).insertAdjacentHTML('beforeend',`<div class="notice">자료를 불러오지 못했습니다: ${esc(error.message)}</div>`);}
}
main();
