import os, base64, hashlib, zipfile, binascii
from pathlib import Path

OUT = Path('/app/report_template.docx')
EXPECTED_LEN = 77823
EXPECTED_SHA256 = '2035af2ca321117f0373d8de5d9b84bd0074b6cfa0e7a2bb85f169601ba6bff3'
EXPECTED_B64_LEN = 103764
PARTS = 26

vals = []
lengths = []
for i in range(PARTS):
    key = f'VKS_TEMPLATE_{i:02d}'
    val = os.environ.get(key)
    if not val:
        raise RuntimeError(f'Missing template environment part: {key}')
    val = val.strip()
    vals.append(val)
    lengths.append(f'{i:02d}:{len(val)}')
joined = ''.join(vals)
print('VKS_TEMPLATE_PART_LENGTHS ' + ','.join(lengths), flush=True)
print(f'VKS_TEMPLATE_JOINED_LEN {len(joined)} expected={EXPECTED_B64_LEN}', flush=True)
try:
    raw = base64.b64decode(joined, validate=True)
except binascii.Error as exc:
    raise RuntimeError(f'VKS template base64 invalid: b64_len={len(joined)} parts=' + ','.join(lengths)) from exc
sha = hashlib.sha256(raw).hexdigest()
if len(raw) != EXPECTED_LEN or sha != EXPECTED_SHA256:
    raise RuntimeError(f'VKS template integrity mismatch: b64_len={len(joined)} len={len(raw)} sha256={sha} parts=' + ','.join(lengths))
OUT.write_bytes(raw)
with zipfile.ZipFile(OUT) as z:
    names=set(z.namelist())
    for req in ('word/document.xml','word/styles.xml','word/footnotes.xml','[Content_Types].xml'):
        if req not in names:
            raise RuntimeError(f'VKS template missing {req}')
print(f'VKS_TEMPLATE_OK b64_len={len(joined)} len={len(raw)} sha256={sha}', flush=True)
