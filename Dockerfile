FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2-dev libxslt-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py fill_template.py generate_pptx_v3.py generate_html.py generate_compitches.py ./
COPY T21_HK_Agencies_Glass_v13.pptx ./

ENV PORT=10000
EXPOSE $PORT

CMD gunicorn app:app \
    --bind 0.0.0.0:$PORT \
    --workers 2 \
    --timeout 120 \
    --access-logfile - \
    --error-logfile -
