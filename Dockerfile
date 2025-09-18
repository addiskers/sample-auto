FROM python:3.11-slim

# install system deps needed for python-pptx and for building wheels quickly
RUN apt-get update && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
    build-essential \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# copy only the files required for pip install first
COPY requirements.txt /app/requirements.txt

# install pip deps
RUN python -m pip install --upgrade pip setuptools wheel \
    && pip install --no-cache-dir -r /app/requirements.txt

# copy the rest of the app
COPY . /app

# ensure upload folder exists
RUN mkdir -p /app/generated_ppts

EXPOSE 5000

# run with gunicorn (production)
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app", "--workers", "2", "--threads", "2", "--timeout", "120"]
