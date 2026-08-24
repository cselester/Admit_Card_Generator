# Admit Card Generator — Phase 1 Prototype

Three pieces, all free-tier and matching your existing template:

1. `web_ui.py` — a simple local web page to run generation (no command line)
2. `generate_admit_cards.py` — the underlying fast local PDF generation (also runnable directly from the CLI)
3. `portal/` — the student-facing lookup page (GitHub Pages) + its backend (Apps Script), and `drive_pipeline/` — Drive sync, indexing, and export report

## 1. Web UI (recommended — no command line)

One-time setup:

```bash
pip install flask openpyxl pypdf reportlab
```

Run it:

```bash
python3 web_ui.py
```

Then open **http://localhost:5050** in your browser. You'll see a form:

- **Student Excel file** — upload your own, or leave blank to use the bundled sample.
- **Sheet / tab name** — must match a tab in that workbook exactly (e.g. `CBT`, `18th Jan 2026`).
- **Limit** — cap rows for a quick test; `0` processes every row.
- **Parallel workers** — how many run at once; 4 is safe on most laptops.
- **Output folder** (optional) — paste your Google Drive Desktop synced folder path here to skip a manual copy step. Leave blank and it just gives you a zip to download.

Click **Generate admit cards** and you'll land on a results page with the count, speed, and two downloads: a zip of every PDF, and the `batch_manifest.xlsx` (needed by the Drive indexing step in Phase 3 below). The page just wraps the same generation code described next — nothing about the underlying speed or output changes.

Leave the terminal window open while you use it; closing it stops the local server. To stop deliberately, go back to that window and press Ctrl+C.

## 2. Generation script (CLI, if you prefer it)

```bash
pip install openpyxl pypdf reportlab
python3 generate_admit_cards.py --sheet CBT --limit 0 --workers 8
```

- `--sheet` — which sheet in the workbook to process (matches your batch tabs, e.g. `CBT`, `18th Jan 2026`).
- `--limit` — cap the number of rows for testing; use `0` to process the whole sheet.
- `--workers` — parallel processes; set to roughly your CPU core count.

Output PDFs land in `output_cards/`, one file per student named by registration number.

**Tested performance** (this environment, single core): ~50–110 cards/sec. On a normal multi-core laptop this scales up further. At that rate, 30,000 records finish in single-digit minutes rather than an hour, and 100,000 in well under 30.

**How it works:** `template_bg.pdf` is your existing template rendered once. `FIELD_MAP` in the script records where each placeholder sits (found once via the template, reused for every student). For each student, the script draws a small overlay — whiteout the placeholder text, draw the real value — and merges it onto a fresh copy of the background page. No Google Docs API calls, no per-student network round trip.

**When the template changes:** re-render the new template to PDF, re-locate the field positions (a short one-time step, e.g. with `pdfplumber`'s `extract_words()`), and update `FIELD_MAP`. Everything else stays the same.

**Next step to wire into your Drive backup flow:** after generation, batch-upload the `output_cards/*.pdf` files to a Drive folder via the Drive API (parallel calls, not Apps Script's serial ones), then write each student's registration number + Drive link into an `Index` sheet — that sheet is what the portal backend reads from, and it's also your exportable backup record.

## 3. Student portal

### Frontend — `portal/index.html`

A single static page: student enters registration number + last 4 mobile digits, gets back their Drive link. No build step — deploy as-is.

**To host on GitHub Pages:**

1. Push this repo (or just the `portal/` contents) to a GitHub repository.
2. Repo Settings → Pages → deploy from the branch/folder containing `index.html`.
3. Add teammates as collaborators on the repo so others can maintain it.

### Backend — `portal/backend/Code.gs`

An Apps Script Web App reading from an `Index` sheet (Registration Number, Student Name, Mobile Last4, Drive Link, Batch). This is the same index the generation step should write to.

**To deploy:**

1. Open the spreadsheet that holds your `Index` sheet → Extensions → Apps Script.
2. Paste in `Code.gs`.
3. Deploy → New deployment → Web app → execute as "Me", access "Anyone".
4. Copy the deployed URL into `BACKEND_URL` at the top of the `<script>` block in `portal/index.html`.

Both pieces are free — GitHub Pages hosting and Apps Script Web Apps have no cost at this scale.

## 4. Drive sync, index, and export report

Once you're happy with local generation, this phase gets the PDFs into Drive, builds the link index, and produces a downloadable Excel report — all without any new cloud account or API credentials.

### How it works

1. **Google Drive for Desktop** (free — [drive.google.com/drive/download](https://www.google.com/drive/download/)) syncs a local folder to a folder in your Drive automatically. Point the generator's output at that synced folder:

   ```bash
   python3 generate_admit_cards.py --sheet CBT --limit 0 --workers 8 --output-dir "/path/to/your/GoogleDrive/AdmitCards/CBT_batch"
   ```

   This writes every PDF _and_ a `batch_manifest.xlsx` (registration number, student details, filename) into that folder. Drive for Desktop uploads them in the background — no code needed for the upload step itself.

2. **`drive_pipeline/DriveSync.gs`** — paste this into the same Apps Script project as the portal backend (`portal/backend/Code.gs`). It adds an "Admit Cards" menu to your spreadsheet with two actions:
   - **Build/update Index from Drive folder** — shares the synced _folder_ once (every PDF inside inherits that sharing automatically, which is what keeps this fast — no per-file API calls), scans it, reads `batch_manifest.xlsx` for student details, and writes/updates the `Index` sheet (the same one the student portal reads from) in a single batched write.
   - **Export link report (.xlsx)** — copies the Index sheet into a fresh spreadsheet, exports it as a `.xlsx` file into a "Admit Card Backups" Drive folder, and shows you the link. Run this once per batch for your permanent, exportable record.

### One-time setup

1. In the Apps Script editor: **Services** (+ icon) → add **Drive API** (this is the "Advanced Drive Service," built into Apps Script — no external console project or credentials needed, just a toggle).
2. In `DriveSync.gs`, set `DRIVE_FOLDER_ID` to your synced output folder's ID (right-click the folder in Drive → "Share" or check the URL for the ID after `/folders/`).
3. Reload the spreadsheet — the "Admit Cards" menu appears. Run the two menu items in order after each batch.

### Note on "anyone with link" sharing

If your Google Workspace has external sharing restrictions, `folder.setSharing(ANYONE_WITH_LINK)` may be blocked by org policy — `buildIndexFromFolder()` will show an alert if this happens rather than failing silently. In that case, switch it to share with your domain instead (`DriveApp.Access.DOMAIN_WITH_LINK`) — one-line change — so only people signed into your org's Google account can open the links, while everything else stays the same.

### Note on speed

Sharing is set once on the _folder_, not per file — this is what keeps `buildIndexFromFolder()` fast at scale. An earlier version of this script called `setSharing()` on every individual PDF, which is a separate network round-trip per file (~1-1.5 sec each) and made even 400 files take ~10 minutes; at 30,000-100,000 files that approach would take hours and exceed Apps Script's execution time limit entirely. If you're updating from that version, just pulling the latest `DriveSync.gs` fixes it — no other changes needed.

## 5. Sample output

`sample_output/` in this prototype has a small batch of generated cards from your `CBT` sheet — open a couple to see the visual match against your original template.
