FROM python:3.9-slim

# 필수 시스템 패키지 설치
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

# 파이썬 라이브러리 설치
RUN pip install --no-cache-dir -r requirements.txt

# Playwright 브라우저 설치
RUN playwright install chromium
RUN playwright install-deps

# 포트 설정 및 실행 (파일명이 다르면 마지막 부분을 수정하세요!)
ENV PORT=8080
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 event_legalcheck:app
