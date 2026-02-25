FROM python:3.12-slim

# Install only runtime dependencies (no build-essential/cmake)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libopenblas0 \
    liblapack3 \
    libx11-6 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install dlib-bin (pre-compiled wheel, no compilation needed)
# Then install remaining dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir dlib-bin && \
    pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "300", "--workers", "1", "app:app"]
