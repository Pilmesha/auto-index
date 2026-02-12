# Dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Run your script
CMD ["gunicorn", "well2:app", "--workers", "1", "--threads", "4", "--timeout", "120"]
