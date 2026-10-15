"""Read every stored XLS cell from ORS ZIP members, including header/metadata rows."""
import hashlib
import zipfile
from pathlib import Path

def read_full_zip(path):
    import xlrd
    path=Path(path)
    result=[]
    with zipfile.ZipFile(path) as z:
        entries=[e for e in z.infolist() if not e.is_dir() and e.filename.lower().endswith('.xls')]
        if sum(e.file_size for e in entries)>200_000_000:raise ValueError('ORS ZIP 내부 XLS 합계가 200 MB를 넘습니다.')
        if len({e.filename for e in entries})!=len(entries):raise ValueError('ORS ZIP 내부 파일명 중복')
        for entry in entries:
            payload=z.read(entry)
            member_sha=hashlib.sha256(payload).hexdigest()
            wb=xlrd.open_workbook(file_contents=payload)
            for sh in wb.sheets():
                rows=[];cell_count=0;error_count=0
                for r in range(sh.nrows):
                    cells=[]
                    for c in range(sh.ncols):
                        cell=sh.cell(r,c)
                        if cell.ctype==xlrd.XL_CELL_EMPTY:continue
                        value=cell.value
                        if cell.ctype==xlrd.XL_CELL_DATE:value=xlrd.xldate_as_datetime(value,wb.datemode).isoformat()
                        if cell.ctype==xlrd.XL_CELL_ERROR:
                            value=f'원본 오류 코드 {value}';error_count+=1
                        if isinstance(value,float) and value.is_integer():value=int(value)
                        cells.append([c+1,value,cell.ctype])
                        cell_count+=1
                    if cells:rows.append([r+1,cells])
                result.append({'container':path.name,'filename':entry.filename,'sha256':member_sha,'sheet':sh.name,'nrows':sh.nrows,'ncols':sh.ncols,'cells':cell_count,'errors':error_count,'rows':rows})
    return result
