FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=9000 \
    DATA_DIR=/app/data \
    DATABASE_PATH=/app/data/database.db \
    UPLOADS_PRODUCTOS_DIR=/app/static/uploads/productos

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends build-essential libjpeg62-turbo-dev zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./

RUN pip install --upgrade pip \
    && pip install -r requirements.txt

COPY . .

RUN mkdir -p /app/data /app/static/uploads/productos

EXPOSE 9000

# "exec" reemplaza al shell por gunicorn: asi recibe la senal de "docker stop" y
# cierra limpio (sin exec, Docker espera 10 s y lo mata: codigo 137).
CMD ["sh", "-c", "exec gunicorn --preload --bind 0.0.0.0:${PORT:-9000} --workers ${GUNICORN_WORKERS:-1} --threads ${GUNICORN_THREADS:-4} --timeout ${GUNICORN_TIMEOUT:-120} wsgi:application"]
