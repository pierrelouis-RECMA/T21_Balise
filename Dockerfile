FROM python:3.12-slim

# System dependencies (lxml needs libxml2)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2-dev libxslt-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first (layer cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code + template
COPY app.py fill_template.py generate_pptx_v2.py ./
COPY T21_HK_Agencies_Glass_v13.pptx ./

# Render injects $PORT at runtime (default 10000)
ENV PORT=10000

EXPOSE $PORT

CMD gunicorn app:app \
    --bind 0.0.0.0:$PORT \
    --workers 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile -
