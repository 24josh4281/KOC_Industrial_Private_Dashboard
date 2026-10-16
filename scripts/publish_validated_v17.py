"""Record hashes of the committed v17 assets, after the existing QA steps pass."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess

REPOSITORY = '24josh4281/KOC_Industrial_Private_Dashboard'
ALLOWED = {'.html','.css','.js','.json','.xlsx','.pdf'}

def git(root, *args):
    return subprocess.check_output(['git','-C',str(root),*args])

def build_pointer(root):
    revision=git(root,'rev-parse','HEAD').decode().strip()
    if not re.fullmatch('[0-9a-f]{40}',revision):
        raise ValueError('Invalid committed revision')
    manifest=json.loads(git(root,'show',revision+':v17/manifest.json'))
    assets={}
    for raw in git(root,'ls-tree','-r','-z',revision,'--','v17').split(b'\0'):
        if not raw:
            continue
        metadata,filename=raw.split(b'\t',1)
        name=filename.decode('utf8')
        if Path(name).suffix not in ALLOWED or any(p.startswith('.') or p=='__pycache__' for p in Path(name).parts):
            continue
        if metadata.split()[0] not in (b'100644',b'100755'):
            raise ValueError('Only regular committed files are supported: '+name)
        blob=git(root,'show',revision+':'+name)
        if len(blob)>25*1024*1024:
            raise ValueError('Asset exceeds the supported 25 MiB limit: '+name)
        assets[name]={'sha256':hashlib.sha256(blob).hexdigest(),'bytes':len(blob)}
    required={'v17/statistics.html','v17/progress.html','v17/work.html','v17/case.html',
              'v17/app.js','v17/app.css','v17/manifest.json','v17/data/statistics.json',
              'v17/data/progress.json','v17/data/work.json','v17/외부사업_통합관리_v17.xlsx'}
    if not required.issubset(assets):
        raise ValueError('Required dashboard assets missing')
    return {'schema_version':1,'source_repository':REPOSITORY,'source_ref':revision,
            'dashboard':'v17','version':manifest['version'],'as_of':manifest['as_of'],
            'freshness':manifest['freshness'],
            'validation':{'status':'passed','checks':['verify_v15','verify_v16','verify_v17'],
                          'github_run_id':os.environ.get('GITHUB_RUN_ID','local-verification')},
            'files':assets}

def publish(root):
    pointer=build_pointer(root)
    latest=root/'sync/latest-v17.json'
    if latest.exists():
        previous=json.loads(latest.read_text(encoding='utf8'))
        if (previous.get('schema_version')==1 and previous.get('files')==pointer['files']
                and previous.get('version')==pointer['version']):
            print('Unchanged validated v17 assets; no new sync pointer.')
            return False
    payload=json.dumps(pointer,ensure_ascii=False,sort_keys=True,separators=(',',':'))+'\n'
    historical=root/'sync/v17'/f'{pointer["source_ref"]}.json'
    historical.parent.mkdir(parents=True,exist_ok=True)
    if historical.exists() and historical.read_text(encoding='utf8')!=payload:
        raise ValueError('Refusing to overwrite an immutable release descriptor')
    if not historical.exists():
        historical.write_text(payload,encoding='utf8')
    latest.write_text(payload,encoding='utf8')
    print(f'Validated v17 descriptor: {pointer["source_ref"]}, {len(pointer["files"])} files.')
    return True

def record_status(root, status):
    target=root/'sync/refresh-status.json'
    target.parent.mkdir(parents=True,exist_ok=True)
    latest=root/'sync/latest-v17.json'
    pointer=json.loads(latest.read_text(encoding='utf8')) if latest.exists() else {}
    payload={'schema_version':1,'status':status,'source_ref':pointer.get('source_ref'),
             'version':pointer.get('version'),'github_run_id':os.environ.get('GITHUB_RUN_ID'),
             'input_commit':os.environ.get('GITHUB_SHA')}
    target.write_text(json.dumps(payload,ensure_ascii=False,sort_keys=True)+'\n',encoding='utf8')

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--repo',type=Path,default=Path.cwd())
    parser.add_argument('--status',choices=['success','failure','cancelled'])
    args=parser.parse_args()
    if args.status:
        record_status(args.repo.resolve(),args.status)
    else:
        publish(args.repo.resolve())
