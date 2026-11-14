# Локальный образ API. Исходники при запуске перекрываются томом с диска.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY api.py crm.py models.py ./

EXPOSE 8000

# Watchdog сам поднимает uvicorn заново, когда на диске меняется .py.
CMD ["watchmedo", "auto-restart", "--directory=/app", "--pattern=**/*.py", "--recursive", "--debug-force-polling", "--interval=1", "--debounce-interval=0.5", "--signal=SIGTERM", "--", "python", "-m", "uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]
