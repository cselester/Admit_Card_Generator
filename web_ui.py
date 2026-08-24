"""
Local web UI for the admit card generator — run this instead of the CLI.

    python3 web_ui.py

Then open http://localhost:5050 in your browser. Upload your student Excel
file (or leave blank to use the bundled sample), pick the sheet/tab, set a
limit if you just want to test, and click Generate. When it finishes you get
a results page with a zip of every PDF plus the manifest, ready to download
or drop straight into your Google Drive–synced folder.

This runs entirely on your own machine — nothing leaves your computer.
"""

import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from flask import Flask, request, render_template_string, send_file, url_for, redirect

import generate_admit_cards as gen

APP_DIR = Path(__file__).parent
DEFAULT_XLSX = APP_DIR / "Warrior_Admit_card_Maker_Tool____MV.xlsx"
RUNS_DIR = APP_DIR / "web_runs"
RUNS_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

PAGE_STYLE = """
<style>
  :root { --pw-red:#9e1b24; --ink:#1c1614; --paper:#faf6f2; --line:#e7dcd4; --muted:#8a7a72; }
  * { box-sizing: border-box; }
  body { margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
         background:var(--paper); color:var(--ink); display:flex; justify-content:center; }
  main { width:100%; max-width:560px; padding:2.5rem 1.5rem; }
  h1 { font-size:1.4rem; margin:0 0 0.25rem; }
  .sub { color:var(--muted); font-size:0.9rem; margin:0 0 1.75rem; }
  .card { background:#fff; border:1px solid var(--line); border-radius:14px;
          box-shadow:0 12px 30px -14px rgba(30,12,12,0.2); padding:1.75rem; }
  label { display:block; font-size:0.82rem; font-weight:600; margin:1rem 0 0.35rem; }
  label:first-child { margin-top:0; }
  input, select { width:100%; padding:0.6rem 0.75rem; border:1px solid var(--line);
                  border-radius:8px; font-size:0.92rem; background:var(--paper); }
  .hint { color:var(--muted); font-size:0.78rem; margin-top:0.3rem; }
  .row { display:flex; gap:1rem; }
  .row > div { flex:1; }
  button { margin-top:1.5rem; width:100%; padding:0.8rem; border:none; border-radius:8px;
           background:var(--pw-red); color:#fff; font-size:0.95rem; font-weight:700; cursor:pointer; }
  button:hover { background:#7a1119; }
  .result-item { display:flex; align-items:center; justify-content:space-between;
                 padding:0.9rem 1rem; border:1px solid var(--line); border-radius:10px; margin-bottom:0.75rem; }
  .result-item a { background:#1c6b3a; color:#fff; text-decoration:none; font-weight:700;
                    padding:0.5rem 0.9rem; border-radius:8px; font-size:0.85rem; }
  .stats { color:var(--muted); font-size:0.85rem; line-height:1.6; margin:1rem 0 1.5rem; }
  .back { display:inline-block; margin-top:1.25rem; color:var(--pw-red); text-decoration:none;
          font-weight:600; font-size:0.88rem; }
</style>
"""

FORM_PAGE = """
<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Admit Card Generator</title>""" + PAGE_STYLE + """</head>
<body><main>
  <h1>Admit Card Generator</h1>
  <p class="sub">Generate admit cards from your student sheet — no command line needed.</p>
  <div class="card">
    <form action="/generate" method="post" enctype="multipart/form-data">
      <label>Student Excel file</label>
      <input type="file" name="xlsx_file" accept=".xlsx">
      <div class="hint">Leave blank to use the bundled sample workbook.</div>

      <label>Sheet / tab name</label>
      <input type="text" name="sheet" value="CBT" required>
      <div class="hint">Must match a tab name in the workbook exactly, e.g. "CBT", "18th Jan 2026".</div>

      <div class="row">
        <div>
          <label>Limit (0 = all rows)</label>
          <input type="number" name="limit" value="25" min="0">
        </div>
        <div>
          <label>Parallel workers</label>
          <input type="number" name="workers" value="4" min="1" max="32">
        </div>
      </div>

      <label>Output folder (optional)</label>
      <input type="text" name="output_dir" placeholder="Leave blank for default, or paste your Drive-synced folder path">
      <div class="hint">Point this at your Google Drive Desktop synced folder to skip a manual copy step.</div>

      <button type="submit">Generate admit cards</button>
    </form>
  </div>
</main></body></html>
"""

RESULT_PAGE = """
<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Admit Cards Generated</title>""" + PAGE_STYLE + """</head>
<body><main>
  <h1>Done</h1>
  <p class="sub">{{ count }} admit card(s) generated in {{ elapsed }}s ({{ rate }}/sec).</p>
  <div class="card">
    <div class="stats">
      Projected at this rate — 30,000 records: {{ proj_30k }} min · 100,000 records: {{ proj_100k }} min
      {% if output_dir %}<br>Written to: {{ output_dir }}{% endif %}
    </div>
    <div class="result-item">
      <span>All PDFs (.zip)</span>
      <a href="{{ zip_url }}">Download</a>
    </div>
    <div class="result-item">
      <span>Manifest (.xlsx)</span>
      <a href="{{ manifest_url }}">Download</a>
    </div>
  </div>
  <a class="back" href="/">&larr; Run another batch</a>
</main></body></html>
"""

ERROR_PAGE = """
<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Error</title>""" + PAGE_STYLE + """</head>
<body><main>
  <h1>Something went wrong</h1>
  <div class="card"><p>{{ message }}</p></div>
  <a class="back" href="/">&larr; Back</a>
</main></body></html>
"""


@app.route("/")
def index():
    return render_template_string(FORM_PAGE)


@app.route("/generate", methods=["POST"])
def generate():
    sheet = request.form.get("sheet", "CBT").strip()
    limit = int(request.form.get("limit", 0) or 0)
    workers = max(1, int(request.form.get("workers", 4) or 4))
    output_dir_input = request.form.get("output_dir", "").strip()

    run_id = uuid.uuid4().hex[:8]
    run_dir = RUNS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    xlsx_file = request.files.get("xlsx_file")
    if xlsx_file and xlsx_file.filename:
        xlsx_path = run_dir / "input.xlsx"
        xlsx_file.save(xlsx_path)
    else:
        if not DEFAULT_XLSX.exists():
            return render_template_string(ERROR_PAGE,
                message="No Excel file uploaded and no bundled sample found. Please upload one."), 400
        xlsx_path = DEFAULT_XLSX

    out_dir = Path(output_dir_input) if output_dir_input else (run_dir / "output_cards")
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return render_template_string(ERROR_PAGE, message=f"Couldn't create output folder: {e}"), 400

    try:
        records = gen.load_records(xlsx_path, sheet, limit or None)
    except KeyError:
        return render_template_string(ERROR_PAGE,
            message=f'Sheet "{sheet}" not found in that workbook. Check the tab name and try again.'), 400

    if not records:
        return render_template_string(ERROR_PAGE, message="No rows found in that sheet."), 400

    start = time.time()
    from concurrent.futures import ProcessPoolExecutor, as_completed
    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(gen.render_one, rec, out_dir) for rec in records]
        for f in as_completed(futures):
            f.result()
    elapsed = time.time() - start

    gen.write_manifest(records, out_dir, sheet)

    zip_base = run_dir / f"admit_cards_{sheet}"
    zip_path = shutil.make_archive(str(zip_base), "zip", root_dir=out_dir)

    rate = len(records) / elapsed if elapsed > 0 else 0

    return render_template_string(
        RESULT_PAGE,
        count=len(records),
        elapsed=round(elapsed, 2),
        rate=round(rate, 1),
        proj_30k=round(30000 / rate / 60, 1) if rate else "-",
        proj_100k=round(100000 / rate / 60, 1) if rate else "-",
        output_dir=output_dir_input or None,
        zip_url=url_for("download", run_id=run_id, filename=Path(zip_path).name),
        manifest_url=url_for("download", run_id=run_id, filename="output_cards/batch_manifest.xlsx"
                              if not output_dir_input else "batch_manifest.xlsx",
                              _external=False) if not output_dir_input
                     else f"/download_external?path={out_dir / 'batch_manifest.xlsx'}",
    )


@app.route("/download/<run_id>/<path:filename>")
def download(run_id, filename):
    file_path = RUNS_DIR / run_id / filename
    if not file_path.exists():
        return "File not found", 404
    return send_file(file_path, as_attachment=True)


@app.route("/download_external")
def download_external():
    path = Path(request.args.get("path", ""))
    if not path.exists():
        return "File not found", 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    print("Admit Card Generator running at http://localhost:5050")
    app.run(host="127.0.0.1", port=5050, debug=False)
