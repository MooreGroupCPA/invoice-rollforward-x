FROM python:3.12-slim

# Install LibreOffice for headless conversion
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render expects you to listen on $PORT
ENV PORT=10000
EXPOSE 10000

CMD gunicorn -b 0.0.0.0:${PORT} app:app