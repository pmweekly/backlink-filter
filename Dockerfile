FROM node:22-alpine AS frontend-build

WORKDIR /app
COPY package*.json ./
RUN if [ -f package-lock.json ]; then npm ci; else npm install; fi
COPY index.html ./
COPY src ./src
COPY web ./web
RUN npm run build

FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    BACKLINK_STORAGE_DIR=/app/storage \
    BACKLINK_MAX_UPLOAD_MB=100 \
    CHECK_BLOGS_DELAY_SECONDS=1.5

WORKDIR /app

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY check_blogs ./check_blogs
COPY processor ./processor
COPY web ./web
COPY --from=frontend-build /app/web/static ./web/static

RUN mkdir -p /app/storage/jobs

EXPOSE 8000

CMD ["uvicorn", "web.app:app", "--host", "0.0.0.0", "--port", "8000"]
