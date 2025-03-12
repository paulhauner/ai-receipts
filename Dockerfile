FROM python:3.10-slim

WORKDIR /app

# Install required system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first to leverage Docker cache
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY main.py .

# Create directory for config files
RUN mkdir -p /app/config

# Set environment variable for config path
ENV CONFIG_PATH=/app/config/config.yaml

# Create volume mount points
VOLUME ["/app/config"]

CMD ["python", "main.py"]