from flask import Flask, Response
from pathlib import Path
import generator

p = Path(generator.__file__)
src = p.read_text(encoding='utf-8')
for i in range(0, len(src), 1500):
    print(f'GENSRC[{i//1500:04d}]:' + src[i:i+1500].replace('\n','\\n'), flush=True)
print(f'GENSRC_END:{len(src)}', flush=True)

app = Flask(__name__)

@app.get('/')
def index():
    return Response(src, mimetype='text/plain; charset=utf-8')

@app.get('/health')
def health():
    return {'ok': True, 'generator': str(generator.__file__), 'length': len(src)}
