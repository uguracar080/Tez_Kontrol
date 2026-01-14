FROM python:3.12-slim

# --- Build-time metadata (GitHub Actions build-args ile dolar) ---
ARG APP_VERSION=dev
ARG GIT_SHA=unknown
ARG BUILD_TIME=""
ENV APP_VERSION=${APP_VERSION}
ENV GIT_SHA=${GIT_SHA}
ENV BUILD_TIME=${BUILD_TIME}

# (ReportLab gibi paketler için sık kullanılan runtime kütüphaneleri)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfreetype6 \
    libjpeg62-turbo \
    zlib1g \
    fonts-dejavu \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Önce bağımlılıklar (cache için)
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Uygulama dosyaları
COPY app.py /app/app.py
COPY Tez_Kontrol.py /app/Tez_Kontrol.py
COPY rules.yaml /app/rules.yaml
COPY report.yaml /app/report.yaml
COPY static /app/static
COPY fonts /app/fonts


# app.py içinde uploads_tmp ve reports_tmp zaten oluşturuluyor
EXPOSE 8000

CMD ["uvicorn", "app:app", "--host=0.0.0.0", "--port=8000"]
