FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-core libreoffice-common antiword fonts-liberation2 fontconfig && rm -rf /var/lib/apt/lists/*
WORKDIR /app
RUN pip install --no-cache-dir flask==3.1.2 werkzeug==3.1.3 python-docx==1.2.0 openpyxl==3.1.5 lxml==6.1.3 waitress==3.0.2
COPY gen_v3 /tmp/gen_v3
RUN cat /tmp/gen_v3/p*.txt | base64 -d | gzip -dc > /app/generator.py && python -m py_compile /app/generator.py
COPY template_b64 /tmp/template_b64
RUN cat /tmp/template_b64/p*.txt | base64 -d > /app/report_template.docx && python -c "import zipfile; z=zipfile.ZipFile('/app/report_template.docx'); assert 'word/document.xml' in z.namelist()"
COPY three_file/app.py /app/app.py
COPY doc_support/footnote_fix.py /app/footnote_fix.py
EXPOSE 10000
CMD ["sh","-c","waitress-serve --listen=0.0.0.0:${PORT:-10000} --threads=4 app:app"]
