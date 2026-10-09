FROM node:22-bookworm-slim AS frontend-build

WORKDIR /app/web
COPY web/package.json web/package-lock.json ./
RUN npm ci
COPY web/ ./
RUN npm run build

FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    DRAFTER_BROWSER_CHANNEL=chromium

WORKDIR /app
COPY requirements.txt ./
RUN python -m pip install --no-cache-dir -r requirements.txt \
    && python -m playwright install --with-deps chromium \
    && apt-get update \
    && apt-get install -y --no-install-recommends xvfb fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY api.py drafter_to_excel.py ./
COPY --from=frontend-build /app/web/dist ./web/dist

RUN useradd --create-home --uid 10001 appuser \
    && chown -R appuser:appuser /app /ms-playwright

USER appuser
EXPOSE 8000

CMD ["sh", "-c", "exec xvfb-run -a python -m uvicorn api:app --host 0.0.0.0 --port ${PORT:-8000} --proxy-headers --forwarded-allow-ips='*'"]