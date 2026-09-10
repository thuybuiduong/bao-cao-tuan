from flask import Flask, Response
from pathlib import Path
import hashlib, json

app = Flask(__name__)
BASE=Path('/app')
def listing():
    out=[]
    for p in sorted(BASE.rglob('*')):
        if p.is_file():
            rel=str(p.relative_to(BASE))
            if p.stat().st_size < 2_000_000:
                h=hashlib.sha256(p.read_bytes()).hexdigest()[:16]
            else:h='large'
            out.append({'path':rel,'size':p.stat().st_size,'sha256':h})
    return out
print('FILELIST='+json.dumps(listing(),ensure_ascii=False),flush=True)

@app.get('/')
def index():
    return Response(json.dumps(listing(),ensure_ascii=False,indent=2),mimetype='application/json; charset=utf-8')

@app.get('/health')
def health(): return {'ok':True}
