# Batch Word Formatter (local website)

This repo includes a tiny local website to batch-format Word documents:

- Body font: Times New Roman, 9pt
- Header + footer + page numbering (based on `format_papers.py`)
- Upload multiple files → download a ZIP of formatted `.docx`
- Header fields: Volume (default 36), Issue, Year, Start page

## Setup

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python3 webapp.py
```

Then open:

- `http://127.0.0.1:8000`

## Deploy (so others can use it)

### Option A: share your computer (fastest)

Use a tunneling tool to expose `http://127.0.0.1:8000` to the internet.

**Cloudflare Tunnel (quick tunnel)**

1. Install `cloudflared` (Cloudflare Tunnel client).
2. Start a public URL:

```bash
./run_public_tunnel.sh
```

Cloudflare will print a public `https://…trycloudflare.com` URL you can share.

### Option B: deploy as a public site (recommended)

This repo includes a `Dockerfile`, so you can deploy to any Docker host (VPS, Render, Fly.io, etc).

Build and run locally:

```bash
docker build -t wordd .
docker run --rm -p 8000:8000 wordd
```

On a server, run the same `docker run` behind a reverse proxy (Nginx/Caddy) with HTTPS.

**Hugging Face Spaces (Docker)**

Spaces requires your container to listen on port `7860`. The included `Dockerfile` is configured for that.

**Environment variables**

- `PORT` (default `8000`)
- `HOST` (default `0.0.0.0` in Docker)
- `MAX_UPLOAD_MB` (default `50`)

## Notes

- `.docx` is supported directly.
- `.doc` uploads require LibreOffice (`soffice`) installed on your machine.
