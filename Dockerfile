# Dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Expose a port for web service (Render / Fly.io requires it)
EXPOSE 8080

# Run your script
CMD ["python", "well2.py"]
