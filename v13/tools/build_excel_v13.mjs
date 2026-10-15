import fs from 'node:fs/promises';
import zlib from 'node:zlib';
import {SpreadsheetFile} from '@oai/artifact-tool';

const arg=name=>process.argv.find(x=>x.startsWith(`--${name}=`))?.slice(name.length+3);
const base=arg('root')||'B:/CODEX/KOC_Industrial/outputs/2026-10-15_meeting_v13';
const old=arg('base-workbook')||'B:/CODEX/KOC_Industrial/outputs/2026-10-15_meeting_v12/외부사업_통합관리_검토본_v12.xlsx';
const sample=process.argv.includes('--sample');
const data=JSON.parse(await fs.readFile(arg('data')||base+'/dashboard-data.json','utf8'));
const packed=await fs.readFile(base+'/source-catalog.gz.b64','ascii');
const source=JSON.parse(zlib.gunzipSync(Buffer.from(packed,'base64')).toString('utf8'));
const ors=JSON.parse(zlib.gunzipSync(Buffer.from(await fs.readFile(arg('ors')||base+'/ors-full-catalog.gz.b64','ascii'),'base64')).toString('utf8'));
const wb=await SpreadsheetFile.importXlsx(await fs.readFile(old));
const output=arg('output')||base+(sample?'/work/엑셀_원본샘플6_v13.xlsx':'/외부사업_통합관리_전체자료_v13_검증완료.xlsx');
const navy='#102A43', teal='#147D86', blue='#3559D9', pale='#EAF3F5';
const put=(s,addr,v)=>s.getRange(addr).values=[[v]];
const sanitize=v=>typeof v==='string'&&/^[=+@]/.test(v)?"'"+v:v;
const colName=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const index=wb.worksheets.add('원본 안내');
index.showGridLines=false;
put(index,'A1','원본 전체 데이터 · 2026.10.13 관리 Tool');
put(index,'A2','원본01~36은 관리 Excel, ORS01~20은 ZIP 추출 XLS의 저장값입니다. 수식은 별도 목록에 보존합니다. 원본 서식·개체는 동봉 원본을 확인하세요.');
put(index,'A3','원본 행번호와 열번호를 유지합니다. 분류 원장 3,419행과 원본 전체 관리 483,492셀 + ORS 286,725셀은 서로 다른 범위입니다.');
index.getRange('A1:E3').format={font:{name:'Malgun Gothic',size:11,color:navy},rowHeight:27};
index.getRange('A1').format.font={name:'Malgun Gothic',size:16,bold:true,color:navy};
index.getRange('A5:E5').values=[['조회 시트','원본 파일','원본 시트','값·수식 셀','원본 수식']];
index.getRange('A5:E5').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'},rowHeight:30};
const mapping=[], formulas=[],special=[];
let n=0;
for(const book of source.books){
  for(const src of book.sheets){
    n++;
    if(sample&&src.name!=='(비공개)국내평가')continue;
    const name='원본'+String(n).padStart(2,'0');
    mapping.push([name,book.name,src.name,src.cells,src.formulas]);
    const ws=wb.worksheets.add(name);
    ws.showGridLines=false;
    ws.getRange('A1:D3').values=[['원본 파일',book.name,'',''],['원본 시트',src.name,'',''],['읽는 법','A열은 원본 행번호 · B열부터 원본 A열~','수식 셀은 저장값 표시','원본 수식 시트 참고']];
    ws.getRange('A1:D3').format={font:{name:'Malgun Gothic',size:10,color:navy},rowHeight:24};
    ws.getRange('A4:B4').values=[['원본 행','원본 A열']];
    ws.getRange('A4:B4').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'}};
    const maxcol=Math.max(1,...src.rows.flatMap(([_,c])=>c.map(x=>x[0])));
    let batch=[];
    const flush=()=>{
      if(!batch.length)return;
      const start=batch[0][0], vals=batch.map(x=>x[1]);
      ws.getRangeByIndexes(start+3,0,vals.length,maxcol+1).values=vals;
      batch=[];
    };
    for(const [row,cells] of src.rows){
      const line=Array(maxcol+1).fill(null);line[0]=row;
      for(const [col,value,formula,error] of cells){
        line[col]=sanitize(value);
        if(formula)formulas.push([book.name,src.name,`${colName(col)}${row}`,sanitize(formula),sanitize(value),error?'원본 오류':'']);
        if(value===''||error)special.push([book.name,src.name,`${colName(col)}${row}`,value===''?'빈 문자열 저장':'원본 오류 저장',error?String(value??''):null]);
      }
      if(batch.length&&row!==batch[batch.length-1][0]+1)flush();
      batch.push([row,line]);
      if(batch.length>=150)flush();
    }
    flush();
    ws.freezePanes.freezeRows(4);
    ws.getRange('A:A').format.columnWidth=12;
    ws.getRange('B:B').format.columnWidth=30;
  }
}
for(let j=0;j<ors.length;j++){
  if(sample&&j>0)break;
  const item=ors[j],name='ORS'+String(j+1).padStart(2,'0');
  mapping.push([name,item.container,item.filename+' / '+item.sheet,item.cells,0]);
  const ws=wb.worksheets.add(name);ws.showGridLines=false;
  ws.getRange('A1:D3').values=[['원본 ZIP',item.container,'',''],['원본 XLS',item.filename,item.sheet,''],['읽는 법','A열은 원본 행번호 · B열부터 원본 A열~','XLS 저장값','원본 XLS 수식·서식은 ZIP 참고']];
  ws.getRange('A1:D3').format={font:{name:'Malgun Gothic',size:10,color:navy},rowHeight:24};
  ws.getRange('A4:B4').values=[['원본 행','원본 A열']];ws.getRange('A4:B4').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'}};
  let batch=[];
  const flush=()=>{if(!batch.length)return;ws.getRangeByIndexes(batch[0][0]+3,0,batch.length,item.ncols+1).values=batch.map(x=>x[1]);batch=[];};
  for(const [row,cells] of item.rows){
    const line=Array(item.ncols+1).fill(null);line[0]=row;
    for(const [col,value] of cells){
      line[col]=sanitize(value);
      if(value==='')special.push([item.container+' / '+item.filename,item.sheet,`${colName(col)}${row}`,'ORS 빈 문자열 저장',null]);
      else if(typeof value==='string'&&value.includes('\r'))special.push([item.container+' / '+item.filename,item.sheet,`${colName(col)}${row}`,'ORS 줄바꿈 원문(JSON)',JSON.stringify(value)]);
    }
    if(batch.length&&row!==batch[batch.length-1][0]+1)flush();
    batch.push([row,line]);if(batch.length>=150)flush();
  }
  flush();ws.freezePanes.freezeRows(4);ws.getRange('A:A').format.columnWidth=12;ws.getRange('B:B').format.columnWidth=30;
}
if(mapping.length)index.getRangeByIndexes(5,0,mapping.length,5).values=mapping;
index.getRange('A:A').format.columnWidth=16;index.getRange('B:C').format.columnWidth=53;
const formulaSheet=wb.worksheets.add('원본 수식');
formulaSheet.showGridLines=false;
put(formulaSheet,'A1','원본 수식 텍스트 목록 · 재계산용 수식이 아닙니다');
formulaSheet.getRange('A3:F3').values=[['원본 파일','원본 시트','원본 셀','수식 원문','저장값','오류']];
formulaSheet.getRange('A3:F3').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'}};
for(let i=0;i<formulas.length;i+=500){const chunk=formulas.slice(i,i+500);formulaSheet.getRangeByIndexes(i+3,0,chunk.length,6).values=chunk;}
formulaSheet.freezePanes.freezeRows(3);
formulaSheet.getRange('A:B').format.columnWidth=42;formulaSheet.getRange('D:D').format.columnWidth=70;
const specialSheet=wb.worksheets.add('원본 특수셀');
specialSheet.showGridLines=false;
put(specialSheet,'A1','원본에서 빈 문자열로 저장된 셀·기존 오류·ORS 줄바꿈 원문을 주소별 보존');
specialSheet.getRange('A3:E3').values=[['원본 파일','원본 시트','원본 셀','원본 저장 상태','원문 보존값(JSON/오류)']];
specialSheet.getRange('A3:E3').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'}};
for(let i=0;i<special.length;i+=500){const chunk=special.slice(i,i+500);specialSheet.getRangeByIndexes(i+3,0,chunk.length,5).values=chunk;}
specialSheet.freezePanes.freezeRows(3);
specialSheet.getRange('A:B').format.columnWidth=42;

const q=wb.worksheets.add('2026 접수·검토·안내');
q.showGridLines=false;
put(q,'A1','2026년 신청 접수·의견·사무국 안내');
put(q,'A2','원본에서 확인된 신청일 53행만 표시합니다. 미기록은 미실시 확정이 아닙니다. 노란 열을 작성한 뒤 가져오기 도구로 DB에 반영하세요.');
put(q,'A3','접수·의견·안내 판단은 업무 메모입니다. 공식 ORS 상태·법정 기한과 별개로 관리합니다.');
q.getRange('A1:U3').format={font:{name:'Malgun Gothic',size:10,color:navy},rowHeight:27};
q.getRange('A1').format.font={name:'Malgun Gothic',size:16,bold:true,color:navy};
const heads=['업무 연결키','출처 행 ID','업무','사업·신청명','관리키 원문','신청일','신청월','원본 처리상태','내부 진행상태','담당 원문','원본 진행내용·비고','원본 이력','접수 확인','의견 검토 상태','의견·검토 내용','사무국 안내 내용','안내 상태','다음 조치','업무 담당자','수정일','원본 위치'];
q.getRange('A5:U5').values=[heads];
q.getRange('A5:U5').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'},rowHeight:34,wrapText:true};
const records=data.cases_2026.map(r=>{const note=data.workflow[r.case_key];return [r.case_key,r.record_id,r.unit,r.title,r.source_key,r.apply_date,r.month,r.raw_status,r.stage,r.owner,r.notes,r.history,note.intake_status,note.opinion_status,note.opinion_text,note.guidance_text,note.guidance_status,note.next_action,note.owner,note.updated_at,`${r.source_file} / ${r.source_sheet} ${r.source_row}행`].map(sanitize);});
q.getRangeByIndexes(5,0,records.length,21).values=records;
q.getRange(`M6:T${5+records.length}`).format.fill='#FFF4D6';
q.getRange('D:D').format.columnWidth=56;q.getRange('K:L').format.columnWidth=50;q.getRange('O:P').format.columnWidth=56;q.getRange('U:U').format.columnWidth=48;
q.freezePanes.freezeRows(5);q.freezePanes.freezeColumns(4);
q.tables.add(`A5:U${5+records.length}`,true,'Workflow2026');
q.getRange(`M6:M${5+records.length}`).dataValidation={rule:{type:'list',values:['확인 필요','접수 확인','인입 보류','해당 없음']}};
q.getRange(`N6:N${5+records.length}`).dataValidation={rule:{type:'list',values:['미기록','검토 예정','검토 중','의견 정리','검토 완료','확인 필요']}};
q.getRange(`Q6:Q${5+records.length}`).dataValidation={rule:{type:'list',values:['미기록','안내 초안','내부 확인','안내 완료','해당 없음']}};

const activity=data.ors_2026_activity||[];
const aq=wb.worksheets.add('2026 ORS 이력후보');aq.showGridLines=false;
put(aq,'A1','2026년 ORS 이력 확인 후보 · 9월 산업부문 스냅샷');
put(aq,'A2','갱신·보완·처리 사건을 포함하므로 신규 신청 건수로 합산하지 않습니다. 관리 Tool 53행과의 동일 신청 연결은 미확정입니다.');
aq.getRange('A4:I4').values=[['원본 XLS','원본 행','업무 종류','번호','사업·신청명','처리상태','2026 이력 원문','이력 줄수','원본 SHA-256']];
aq.getRange('A4:I4').format={fill:navy,font:{name:'Malgun Gothic',size:10,bold:true,color:'#FFFFFF'},rowHeight:35,wrapText:true};
if(activity.length){aq.getRangeByIndexes(4,0,activity.length,9).values=activity.map(x=>[x.source_file,x.source_row,x.kind,x.number,x.title,x.status,x.events_2026.join('\n'),x.event_count,x.source_sha256].map(sanitize));aq.tables.add(`A4:I${activity.length+4}`,true,'OrsActivity2026');}
aq.freezePanes.freezeRows(4);aq.getRange('A:A').format.columnWidth=48;aq.getRange('E:E').format.columnWidth=56;aq.getRange('G:G').format.columnWidth=75;aq.getRange(`G5:G${activity.length+4}`).format.wrapText=true;

const file=await SpreadsheetFile.exportXlsx(wb);
await file.save(output);
console.log(JSON.stringify({output,sourceSheets:mapping.length,orsSheets:ors.length,formulaRows:formulas.length,specialRows:special.length,cases2026:records.length,orsActivity2026:activity.length}));
