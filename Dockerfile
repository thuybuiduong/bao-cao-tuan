FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-core libreoffice-common antiword fonts-liberation2 fontconfig gzip unzip && rm -rf /var/lib/apt/lists/*
WORKDIR /app
RUN pip install --no-cache-dir flask==3.1.2 werkzeug==3.1.3 python-docx==1.2.0 openpyxl==3.1.5 lxml==6.1.3 waitress==3.0.2

# Mẫu nền duy nhất: Báo cáo tuần VKSND tỉnh Ninh Bình 09/9/2026.
COPY province_template /tmp/province_template
RUN cat /tmp/province_template/p*.txt | base64 -d > /app/report_template.docx && \
    python -c "from pathlib import Path; import hashlib,zipfile; from docx import Document; p=Path('/app/report_template.docx'); raw=p.read_bytes(); sha=hashlib.sha256(raw).hexdigest(); assert len(raw)==77823,(len(raw),sha); assert sha=='2035af2ca321117f0373d8de5d9b84bd0074b6cfa0e7a2bb85f169601ba6bff3',sha; z=zipfile.ZipFile(p); n=set(z.namelist()); assert 'word/document.xml' in n and 'word/styles.xml' in n and 'word/footnotes.xml' in n and '[Content_Types].xml' in n; d=Document(p); t=' '.join(x.text for x in d.paragraphs); assert 'TỈNH NINH BÌNH' in t and 'Tình hình tội phạm và kết quả công tác tuần' in t; print('PROVINCE_TEMPLATE_OK',len(raw),sha,len(d.paragraphs),len(d.tables))"

COPY complete_new/core_parts /tmp/core_parts
RUN cat /tmp/core_parts/p*.txt | base64 -d | gzip -dc > /app/core_overlay.py && python -m py_compile /app/core_overlay.py
COPY complete_new/office_overlay.py.gz.b64 /tmp/office_overlay.b64
RUN cat /tmp/office_overlay.b64 | base64 -d | gzip -dc > /app/office_overlay.py && python -m py_compile /app/office_overlay.py
COPY three_file/excel_overlay.py /app/excel_overlay.py
COPY doc_support/footnote_fix.py /app/footnote_fix.py
COPY complete_new/app.py /app/app.py
COPY patch_app_provincial_guard.py /tmp/patch_app_provincial_guard.py
RUN python /tmp/patch_app_provincial_guard.py && python -m py_compile /app/app.py /app/excel_overlay.py /app/footnote_fix.py
EXPOSE 10000
CMD ["sh","-c","waitress-serve --listen=0.0.0.0:${PORT:-10000} --threads=4 app:app"]
