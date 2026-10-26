# Этот контейнер нужен только для разработки и тестирования backend.
FROM python:3.12-slim

WORKDIR /app/backend

COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY backend .
COPY electron/build/index.html electron/build/style.css /app/electron/build/

EXPOSE 5000

CMD ["python", "app.py"]
