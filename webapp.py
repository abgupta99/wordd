import io
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

try:
    from flask import Flask, Response, request
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


FORM_HTML = """<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Batch Word Formatter</title>
    <style>
      body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 40px; max-width: 860px; }
      h1 { margin: 0 0 10px; }
      .card { border: 1px solid #ddd; border-radius: 12px; padding: 18px; }
      label { display: block; margin: 12px 0 6px; font-weight: 600; }
      input[type="text"], input[type="number"], select { width: 280px; padding: 8px 10px; border: 1px solid #ccc; border-radius: 10px; }
      input[type="file"] { width: 100%; }
      .row { display: flex; gap: 16px; flex-wrap: wrap; }
      .hint { color: #555; font-size: 14px; }
      button { margin-top: 14px; padding: 10px 14px; border: 0; border-radius: 10px; background: #111; color: #fff; font-weight: 700; cursor: pointer; }
      button:hover { background: #000; }
      code { background: #f6f6f6; padding: 2px 6px; border-radius: 6px; }
    </style>
  </head>
  <body>
    <h1>Batch Word Formatter</h1>
    <p class="hint">Upload multiple <code>.docx</code> files, apply the selected journal format, then download a ZIP.</p>
    <div class="card">
      <form action="/format" method="post" enctype="multipart/form-data">
        <label>Format</label>
        <select name="template" id="template" required>
          <option value="msw" selected>MSW (existing)</option>
          <option value="ijrss">IJRSS (new)</option>
          <option value="ijmie">IJMIE (new)</option>
        </select>

        <label>Documents</label>
        <input type="file" name="files" multiple required accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" />
        <div class="row">
          <div>
            <label>Start page</label>
            <input type="number" name="start_page" id="start_page" value="2966" min="1" step="1" />
          </div>
          <div>
            <label>Volume</label>
            <input type="text" name="volume" id="volume" value="36" />
          </div>
          <div>
            <label>Paper year</label>
            <input type="text" name="paper_year" id="paper_year" value="2025" />
          </div>
          <div>
            <label>Issue</label>
            <input type="text" name="issue" id="issue" value="2" />
          </div>
          <div>
            <label>Month (IJRSS/IJMIE)</label>
            <input type="text" name="paper_month" id="paper_month" value="March" />
          </div>
        </div>
        <button type="submit">Format and Download ZIP</button>
      </form>
    </div>
    <p class="hint">Note: <code>.doc</code> uploads require LibreOffice (<code>soffice</code>) installed.</p>
    <script>
      (function () {
        const defaultsByTemplate = {
          msw: { start_page: "2966", volume: "36", issue: "2", paper_year: "2025", paper_month: "March" },
          ijrss: { start_page: "57", volume: "16", issue: "03", paper_year: "2026", paper_month: "March" },
          ijmie: { start_page: "66", volume: "16", issue: "03", paper_year: "2026", paper_month: "March" },
        };

        function setDefaults(template) {
          const defaults = defaultsByTemplate[template] || defaultsByTemplate.msw;
          for (const key of Object.keys(defaults)) {
            const el = document.getElementById(key);
            if (el) el.value = defaults[key];
          }
        }

        const templateSelect = document.getElementById("template");
        if (templateSelect) {
          templateSelect.addEventListener("change", () => setDefaults(templateSelect.value));
        }
      })();
    </script>
  </body>
</html>
"""


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.get("/")
def index():
    return Response(FORM_HTML, mimetype="text/html")


@app.get("/healthz")
def healthz():
    return Response("ok", mimetype="text/plain")


def _format_uploaded_files():
    try:
        start_page = int(request.form.get("start_page", "2966"))
    except Exception:
        return Response("Invalid start page.", status=400)

    template = (request.form.get("template", "msw") or "msw").strip().lower()
    volume = (request.form.get("volume", "36") or "36").strip()
    paper_year = (request.form.get("paper_year", "2025") or "2025").strip()
    issue = (request.form.get("issue", "2") or "2").strip()
    paper_month = (request.form.get("paper_month", "March") or "March").strip()

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

        config = FormatConfig(
            template=template,
            volume=volume,
            paper_year=paper_year,
            issue=issue,
            start_page=start_page,
            paper_month=paper_month,
        )
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


@app.post("/format")
def legacy_format_endpoint():
    return _format_uploaded_files()


def main():
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    app.run(host=host, port=port)


if __name__ == "__main__":
    main()
