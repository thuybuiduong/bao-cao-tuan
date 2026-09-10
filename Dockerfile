FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer libreoffice-core libreoffice-common \
    antiword fonts-liberation2 fontconfig curl unzip \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

RUN curl -fsSL https://raw.githubusercontent.com/thuybuiduong/bao-cao-tuan/3421662dc337bdf93d27b2814f116dcb5a3cb836/patch/c00.txt -o /tmp/c00 \
    && curl -fsSL https://raw.githubusercontent.com/thuybuiduong/bao-cao-tuan/3421662dc337bdf93d27b2814f116dcb5a3cb836/patch/c01.txt -o /tmp/c01 \
    && curl -fsSL https://raw.githubusercontent.com/thuybuiduong/bao-cao-tuan/3421662dc337bdf93d27b2814f116dcb5a3cb836/patch/c02.txt -o /tmp/c02 \
    && cat /tmp/c00 /tmp/c01 /tmp/c02 | base64 -d > /tmp/patch.zip \
    && unzip -o /tmp/patch.zip -d /app \
    && rm -f /tmp/c00 /tmp/c01 /tmp/c02 /tmp/patch.zip

RUN pip install --no-cache-dir -r requirements.txt waitress==3.0.2 lxml==6.1.3

COPY doc_support/debug_server.py /app/debug_server.py

EXPOSE 10000

CMD ["sh","-c","waitress-serve --listen=0.0.0.0:${PORT:-10000} --threads=4 debug_server:app"]
