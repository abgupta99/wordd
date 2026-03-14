FROM python:3.11-slim

WORKDIR /app

# Hugging Face Spaces runs containers as a non-root user by default in their examples.
RUN useradd -m -u 1000 user

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN chown -R user:user /app

USER user
ENV PATH="/home/user/.local/bin:$PATH"

ENV HOST=0.0.0.0
ENV MAX_UPLOAD_MB=50
ENV REQUEST_TIMEOUT=600
ENV WEB_CONCURRENCY=1
ENV GUNICORN_THREADS=4

EXPOSE 10000

CMD ["sh", "-c", "exec gunicorn webapp:app --bind 0.0.0.0:${PORT:-10000} --workers ${WEB_CONCURRENCY} --threads ${GUNICORN_THREADS} --timeout ${REQUEST_TIMEOUT}"]
