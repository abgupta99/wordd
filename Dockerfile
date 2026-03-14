FROM python:3.11-slim

WORKDIR /app

# Enable .doc -> .docx conversion via LibreOffice.
RUN apt-get update \
  && apt-get install -y --no-install-recommends libreoffice \
  && rm -rf /var/lib/apt/lists/*

# Hugging Face Spaces runs containers as a non-root user by default in their examples.
RUN useradd -m -u 1000 user

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN chown -R user:user /app

USER user
ENV PATH="/home/user/.local/bin:$PATH"

ENV HOST=0.0.0.0
ENV PORT=10000
ENV MAX_UPLOAD_MB=50

EXPOSE 10000

CMD ["sh", "-c", "gunicorn -b 0.0.0.0:${PORT} webapp:app --threads 4 --timeout 180"]
