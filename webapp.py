import io
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

try:
    from flask import Flask, Response, abort, request, send_from_directory
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing dependency: Flask.\n\n"
        "Install dependencies, then rerun:\n"
        "  python3 -m venv .venv\n"
        "  . .venv/bin/activate\n"
        "  pip install -r requirements.txt\n"
        "  python3 webapp.py\n"
    ) from exc

from format_papers import FormatConfig, format_docx_files


ROOT_DIR = Path(__file__).resolve().parent
FRONTEND_DIST_DIR = ROOT_DIR / "frontend" / "dist"

MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "50"))


def _safe_filename(name: str) -> str:
    name = os.path.basename(name or "document.docx")
    name = name.replace("\x00", "")
    name = "".join(ch if ch.isalnum() or ch in (" ", ".", "_", "-", "(", ")") else "_" for ch in name)
    return name.strip() or "document.docx"


def _soffice_cmd():
    return shutil.which("soffice") or shutil.which("libreoffice")


def _convert_doc_to_docx(input_path: Path, out_dir: Path) -> Path:
    cmd = _soffice_cmd()
    if not cmd:
        raise RuntimeError(
            "To accept .doc files, install LibreOffice (soffice) or upload .docx instead."
        )

    out_dir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [cmd, "--headless", "--convert-to", "docx", "--outdir", str(out_dir), str(input_path)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed for {input_path.name}.\n\n{proc.stdout}")

    converted = out_dir / (input_path.stem + ".docx")
    if not converted.exists():
        raise RuntimeError(f"LibreOffice reported success but output is missing: {converted.name}")
    return converted


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.get("/")
def index():
    index_path = FRONTEND_DIST_DIR / "index.html"
    if index_path.exists():
        return send_from_directory(FRONTEND_DIST_DIR, "index.html")
    return Response(
        "Frontend not built.\n\n"
        "To run the React UI in dev mode:\n"
        "  (1) python3 webapp.py\n"
        "  (2) cd frontend && npm install && npm run dev\n\n"
        "To build the React UI for Flask to serve:\n"
        "  cd frontend && npm install && npm run build\n",
        mimetype="text/plain",
    )


@app.get("/<path:path>")
def static_or_spa(path: str):
    if path.startswith("api/"):
        abort(404)

    if not FRONTEND_DIST_DIR.exists():
        abort(404)

    file_path = FRONTEND_DIST_DIR / path
    if file_path.exists() and file_path.is_file():
        return send_from_directory(FRONTEND_DIST_DIR, path)

    index_path = FRONTEND_DIST_DIR / "index.html"
    if index_path.exists():
        return send_from_directory(FRONTEND_DIST_DIR, "index.html")

    abort(404)


def _format_uploaded_files():
    try:
        start_page = int(request.form.get("start_page", "2966"))
    except Exception:
        return Response("Invalid start page.", status=400)

    volume = (request.form.get("volume", "36") or "36").strip()
    paper_year = (request.form.get("paper_year", "2025") or "2025").strip()
    issue = (request.form.get("issue", "2") or "2").strip()

    files = request.files.getlist("files")
    if not files:
        return Response("No files uploaded.", status=400)

    with tempfile.TemporaryDirectory(prefix="wordd_upload_") as tmp:
        tmp_path = Path(tmp)
        input_dir = tmp_path / "in"
        output_dir = tmp_path / "out"
        conv_dir = tmp_path / "converted"
        input_dir.mkdir(parents=True, exist_ok=True)

        input_paths = []
        for f in files:
            if not getattr(f, "filename", ""):
                continue

            name = _safe_filename(f.filename)
            dest = input_dir / name
            f.save(dest)

            suffix = dest.suffix.lower()
            if suffix == ".doc":
                dest = _convert_doc_to_docx(dest, conv_dir)
            elif suffix != ".docx":
                return Response(f"Unsupported file type: {name}", status=400)

            input_paths.append(dest)

        if not input_paths:
            return Response("No valid files uploaded.", status=400)

        config = FormatConfig(volume=volume, paper_year=paper_year, issue=issue, start_page=start_page)
        formatted_paths = format_docx_files(
            input_paths,
            output_dir,
            config=config,
            reference_dir=ROOT_DIR,
        )

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for out_path in formatted_paths:
                out_path = Path(out_path)
                zf.write(out_path, arcname=out_path.name)
        zip_buf.seek(0)

        return Response(
            zip_buf.getvalue(),
            mimetype="application/zip",
            headers={"Content-Disposition": 'attachment; filename="formatted_docs.zip"'},
        )


@app.post("/api/format")
def api_format_endpoint():
    return _format_uploaded_files()


@app.post("/format")
def legacy_format_endpoint():
    return _format_uploaded_files()


def main():
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    app.run(host=host, port=port)


if __name__ == "__main__":
    main()
