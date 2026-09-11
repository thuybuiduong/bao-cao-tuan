import os, base64, hashlib, zipfile
from pathlib import Path

OUT = Path('/app/report_template.docx')
EXPECTED_LEN = 77823
EXPECTED_SHA256 = '2035af2ca321117f0373d8de5d9b84bd0074b6cfa0e7a2bb85f169601ba6bff3'
PARTS = 12

vals = []
for i in range(PARTS):
    key = f'VKS_TEMPLATE_{i:02d}'
    val = os.environ.get(key)
    if not val:
        raise RuntimeError(f'Missing template environment part: {key}')
    vals.append(val)
raw = base64.b64decode(''.join(vals), validate=True)
sha = hashlib.sha256(raw).hexdigest()
if len(raw) != EXPECTED_LEN or sha != EXPECTED_SHA256:
    raise RuntimeError(f'VKS template integrity mismatch: len={len(raw)} sha256={sha}')
OUT.write_bytes(raw)
with zipfile.ZipFile(OUT) as z:
    names=set(z.namelist())
    for req in ('word/document.xml','word/styles.xml','word/footnotes.xml','[Content_Types].xml'):
        if req not in names:
            raise RuntimeError(f'VKS template missing {req}')
print(f'VKS_TEMPLATE_OK len={len(raw)} sha256={sha}')
