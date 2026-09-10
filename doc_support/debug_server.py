from flask import Flask, Response
from pathlib import Path
import generator

app = Flask(__name__)

@app.get('/')
def index():
    p = Path(generator.__file__)
    return Response(p.read_text(encoding='utf-8'), mimetype='text/plain; charset=utf-8')

@app.get('/health')
def health():
    return {'ok': True, 'generator': str(generator.__file__)}
