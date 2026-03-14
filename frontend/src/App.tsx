import React, { useMemo, useState } from "react";

type FormState = {
  startPage: number;
  volume: string;
  paperYear: string;
  issue: string;
};

const DEFAULTS: FormState = {
  startPage: 2966,
  volume: "36",
  paperYear: "2025",
  issue: "2",
};

function bytesToMiB(bytes: number) {
  return bytes / (1024 * 1024);
}

export default function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [form, setForm] = useState<FormState>(DEFAULTS);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const totalSizeMiB = useMemo(
    () => files.reduce((acc, f) => acc + bytesToMiB(f.size), 0),
    [files],
  );

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);

    if (files.length === 0) {
      setError("Please choose at least one .docx (or .doc) file.");
      return;
    }

    const body = new FormData();
    for (const f of files) body.append("files", f, f.name);
    body.set("start_page", String(form.startPage));
    body.set("volume", form.volume);
    body.set("paper_year", form.paperYear);
    body.set("issue", form.issue);

    setBusy(true);
    try {
      const resp = await fetch("/api/format", { method: "POST", body });
      if (!resp.ok) {
        const text = await resp.text();
        throw new Error(text || `Request failed (${resp.status})`);
      }

      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "formatted_docs.zip";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="container">
      <h1>Batch Word Formatter</h1>
      <p>
        Upload multiple <code>.docx</code> files, apply Times New Roman 9pt +
        journal header/footer, then download a ZIP.
      </p>
      <div className="card">
        <form onSubmit={onSubmit}>
          <label>Documents</label>
          <input
            type="file"
            multiple
            required
            accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            onChange={(e) => setFiles(Array.from(e.target.files ?? []))}
          />

          {files.length > 0 && (
            <ul className="files">
              {files.map((f) => (
                <li key={`${f.name}-${f.size}-${f.lastModified}`}>
                  {f.name} ({bytesToMiB(f.size).toFixed(1)} MiB)
                </li>
              ))}
              <li>Total: {totalSizeMiB.toFixed(1)} MiB</li>
            </ul>
          )}

          <div className="row">
            <div>
              <label>Start page</label>
              <input
                type="number"
                min={1}
                step={1}
                value={form.startPage}
                onChange={(e) =>
                  setForm((s) => ({ ...s, startPage: Number(e.target.value) }))
                }
              />
            </div>
            <div>
              <label>Volume</label>
              <input
                type="text"
                value={form.volume}
                onChange={(e) => setForm((s) => ({ ...s, volume: e.target.value }))}
              />
            </div>
            <div>
              <label>Paper year</label>
              <input
                type="text"
                value={form.paperYear}
                onChange={(e) =>
                  setForm((s) => ({ ...s, paperYear: e.target.value }))
                }
              />
            </div>
            <div>
              <label>Issue</label>
              <input
                type="text"
                value={form.issue}
                onChange={(e) => setForm((s) => ({ ...s, issue: e.target.value }))}
              />
            </div>
          </div>

          <div className="actions">
            <button type="submit" disabled={busy}>
              {busy ? "Formatting…" : "Format and Download ZIP"}
            </button>
            <span style={{ color: "var(--muted)" }}>
              <span>Note: </span>
              <code>.doc</code> requires LibreOffice (<code>soffice</code>).
            </span>
          </div>

          {error && <div className="error">{error}</div>}
        </form>
      </div>
    </div>
  );
}

