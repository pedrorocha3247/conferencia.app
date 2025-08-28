FROM python:3.11-slim
RUN apt-get update && apt-get install -y \
    build-essential libglib2.0-0 libxrender1 libxext6 libsm6 libx11-6 \
    libfreetype6 libjpeg62-turbo zlib1g && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PYTHONUNBUFFERED=1
CMD ["gunicorn","-w","2","-k","gthread","--threads","4","-t","180","-b","0.0.0.0:${PORT}","app:app"]
