# Use Python 3.12 slim image for smaller size
FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Set environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Install system dependencies if needed (for OAuth flow)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better layer caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY config.py .
COPY src/ ./src/

# Create directories for credentials, logs, and state (will be mounted as volumes)
# These directories will be created if they don't exist when volumes are mounted
RUN mkdir -p /app/credentials /app/logs && \
    chmod 755 /app/credentials /app/logs

# Set default environment variables (can be overridden)
ENV SPREADSHEET_ID="" \
    SHEET_NAME="Emails" \
    LOG_LEVEL="INFO" \
    SUBJECT_FILTER="" \
    EXCLUDE_NO_REPLY="false"

# Run the application
CMD ["python", "-m", "src.main"]
