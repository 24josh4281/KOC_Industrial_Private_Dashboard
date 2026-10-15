"""Create a new, synchronized Excel/HTML snapshot from the working SQLite DB."""
import argparse,base64,gzip,hashlib,json,re,shutil,sqlite3,subprocess,sys
from datetime import datetime,timezone
from pathlib import Path

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
from kets_dashboard.ors_refresh import dated_filename

def saved(path,value):path.write_text(json.dumps(value,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
def activity_for(occurrences,con):
    latest={}
    for x in occurrences:
        if x['scope']!='산업' or x['kind'] not in ('domestic_project','domestic_credit'):continue
        key=x['kind'];dt=x['date'] or ''
        if key not in latest or dt>latest[key]['date']:latest[key]=x
    out=[]
    for x in latest.values():
        headers={c['label']:c['col'] for c in x['columns']};hc=headers.get('이력')
        tc=headers.get('[신청정보] 신청사업명') or headers.get('[사업정보]사업명(단위사업명)')
        sc=headers.get('처리상태');kc=headers.get('번호')
        if not hc:continue
        for rowno,payload in con.execute('SELECT row_no,payload FROM ors_rows WHERE member_sha=? AND sheet=? ORDER BY row_no',(x['sha256'],x['sheet'])):
            row=json.loads(payload);cells={c['col']:c['value'] for c in row['cells']}
            text=str(cells.get(hc) or '')
            events=[line.strip() for line in re.split(r'\r?\n|_x000D_',text) if re.search(r'(?<!\d)2026[-./]\d{1,2}[-./]\d{1,2}',line)]
            if events:out.append({'source_file':x['file'],'source_row':rowno,'kind':x['kind'],'number':cells.get(kc),'title':cells.get(tc),'status':cells.get(sc),'events_2026':events,'event_count':len(events),'source_sha256':x['sha256']})
    return sorted(out,key=lambda a:(a['source_file'],a['source_row']))

def main():
    p=argparse.ArgumentParser(description='월별 ORS/업무 메모 반영 후 새 Excel·HTML 판본 생성')
    p.add_argument('--db',type=Path,default=ROOT/'state/working.sqlite3')
    p.add_argument('--skip-excel',action='store_true',help='웹·JSON만 먼저 생성')
    args=p.parse_args()
    dbpath=args.db if args.db.exists() else ROOT/'external_projects_v13.sqlite3'
    con=sqlite3.connect(dbpath)
    fk=con.execute('PRAGMA foreign_key_check').fetchall();integrity=con.execute('PRAGMA integrity_check').fetchone()[0]
    if fk or integrity!='ok':raise RuntimeError(f'DB 검사 실패: {integrity}, 외래키 {len(fk)}건')
    files=con.execute('SELECT sha,name,kind,source_date FROM source_files ORDER BY sha').fetchall()
    notes=con.execute('SELECT case_key,intake_status,opinion_status,opinion_text,guidance_text,guidance_status,next_action,owner,updated_at FROM workflow_notes ORDER BY case_key').fetchall()
    runs=con.execute('SELECT input_sha FROM import_runs ORDER BY input_sha').fetchall()
    fingerprint=hashlib.sha256(json.dumps([files,notes,runs],ensure_ascii=False).encode()).hexdigest()[:12]
    latest=max((x[3] or '' for x in files),default='2026-10-15')
    dest=ROOT/'snapshots'/f'{latest}_{fingerprint}'
    if dest.exists():
        excel=dest/'외부사업_통합관리_전체자료.xlsx'
        if args.skip_excel or excel.exists():
            print(json.dumps({'unchanged':True,'snapshot':str(dest),'excel_exists':excel.exists()},ensure_ascii=False));return
        subprocess.run([sys.executable,str(ROOT/'tools/refresh_excel.py'),f'--data={dest/"dashboard-data.json"}',f'--ors={dest/"ors-full-catalog.gz.b64"}',f'--output={excel}'],check=True)
        verification=dest/'verification.json'
        current=json.loads(verification.read_text(encoding='utf-8')) if verification.exists() else {}
        current['excel_built']=True;saved(verification,current)
        print(json.dumps({'snapshot':str(dest),'excel_built':True,'resumed':True},ensure_ascii=False));return
    dest.mkdir(parents=True)
    data=json.loads((ROOT/'dashboard-data.json').read_text(encoding='utf-8'))
    for key,intake,opinion,optext,guidance,gstatus,next_action,owner,updated in notes:
        if key not in data['workflow']:raise ValueError(f'원장에 없는 업무 연결키: {key}')
        data['workflow'][key]={'intake_status':intake,'opinion_status':opinion,'opinion_text':optext or '', 'guidance_text':guidance or '', 'guidance_status':gstatus,'next_action':next_action or '', 'owner':owner or '', 'updated_at':updated or ''}
    occurrences=[];full=[]
    sql='''SELECT sf.name,o.member_name,o.member_sha,o.sheet,m.kind,m.scope,m.header_json,r.rows_gzip,r.cell_count,r.nrows,r.ncols,r.error_count FROM ors_occurrences o JOIN source_files sf ON sf.sha=o.zip_sha JOIN ors_members m ON m.member_sha=o.member_sha AND m.sheet=o.sheet JOIN ors_raw_members r ON r.member_sha=o.member_sha AND r.sheet=o.sheet ORDER BY sf.source_date,o.member_name'''
    for zipname,name,sha,sheet,kind,scope,header,blob,cells,nrows,ncols,errors in con.execute(sql):
        dt=dated_filename(name)
        occurrences.append({'zip':zipname,'file':name,'sha256':sha,'sheet':sheet,'date':dt,'kind':kind,'scope':scope,'columns':json.loads(header),'rows':con.execute('SELECT COUNT(*) FROM ors_rows WHERE member_sha=? AND sheet=?',(sha,sheet)).fetchone()[0]})
        full.append({'container':zipname,'filename':name,'sha256':sha,'sheet':sheet,'nrows':nrows,'ncols':ncols,'cells':cells,'errors':errors,'rows':json.loads(gzip.decompress(blob))})
    # The checked v13 workbook fixed ORS01~20 in the original ZIP order. Keep
    # those sheet identities stable; append each subsequent monthly occurrence.
    baseline=json.loads(gzip.decompress(base64.b64decode((ROOT/'ors-full-catalog.gz.b64').read_text(encoding='ascii'))))
    rank={(x['container'],x['filename'],x['sha256'],x['sheet']):i for i,x in enumerate(baseline)}
    paired=list(zip(occurrences,full))
    paired.sort(key=lambda pair:(rank.get((pair[1]['container'],pair[1]['filename'],pair[1]['sha256'],pair[1]['sheet']),len(rank)),pair[0]['date'],pair[0]['file']))
    occurrences=[pair[0] for pair in paired];full=[pair[1] for pair in paired]
    if any((full[i]['container'],full[i]['filename'],full[i]['sha256'],full[i]['sheet'])!=(x['container'],x['filename'],x['sha256'],x['sheet']) for i,x in enumerate(baseline)):
        raise RuntimeError('검증된 기준 ORS 20개 원본이 DB에서 누락되거나 교체되었습니다.')
    data['ors_imports']=[{k:v for k,v in x.items() if k!='columns'} for x in occurrences]
    data['ors_2026_activity']=activity_for(occurrences,con)
    data['source_coverage'].update({'ors_full_occurrences':len(full),'ors_full_source_cells':sum(x['cells'] for x in full),'ors_unique_rows':con.execute('SELECT COUNT(*) FROM ors_rows').fetchone()[0],'ors_occurrences':len(occurrences),'ors_2026_activity_rows':len(data['ors_2026_activity']),'workflow_notes_entered':len(notes)})
    data['refresh_snapshot']={'created_at':datetime.now(timezone.utc).isoformat(),'latest_ors_date':max((x['date'] or '' for x in occurrences),default=''),'db_fingerprint':fingerprint,'management_snapshot_date':'2026-10-13','note':'ORS 수신 원본과 실무 입력은 갱신. 관리 Tool 분류 원장은 새 관리 Excel 반영 전까지 기존 2026-10-13 스냅샷.'}
    con.close()
    jsonpath=dest/'dashboard-data.json';saved(jsonpath,data)
    orspath=dest/'ors-full-catalog.gz.b64';orspath.write_text(base64.b64encode(gzip.compress(json.dumps(full,ensure_ascii=False,separators=(',',':')).encode('utf-8'),compresslevel=9)).decode('ascii'),encoding='ascii')
    html=dest/'대시보드.html';excel=dest/'외부사업_통합관리_전체자료.xlsx'
    page=(ROOT/'대시보드_v13.html').read_text(encoding='utf-8')
    page=re.sub(r'<script id="data" type="application/json">.*?</script>',lambda _:'<script id="data" type="application/json">'+json.dumps(data,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c')+'</script>',page,count=1,flags=re.S)
    page=re.sub(r'<script id="raw-ors" type="application/octet-stream">.*?</script>',lambda _:'<script id="raw-ors" type="application/octet-stream">'+orspath.read_text(encoding='ascii')+'</script>',page,count=1,flags=re.S)
    page=page.replace('source-archive/','../../source-archive/').replace('외부사업_통합관리_전체자료_v13_검증완료.xlsx',excel.name)
    html.write_text(page,encoding='utf-8')
    excel_ok=False
    if not args.skip_excel:
        subprocess.run([sys.executable,str(ROOT/'tools/refresh_excel.py'),f'--data={jsonpath}',f'--ors={orspath}',f'--output={excel}'],check=True)
        excel_ok=True
    result={'snapshot':str(dest),'latest_ors_date':data['refresh_snapshot']['latest_ors_date'],'ors_occurrences':len(full),'workflow_notes':len(notes),'excel_built':excel_ok,'db_integrity':integrity,'foreign_key_errors':len(fk)}
    saved(dest/'verification.json',result)
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':main()
