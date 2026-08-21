# LinkedIn Analysis PPT Generator

Upload an Excel **workbook** (one sheet per executive) → get back one
`Sample.pptx`-based PowerPoint with every sheet's slides included, in
order, each with its own LinkedIn URL / Connections header and its own
"(1/2)" page counter.

This guide assumes **zero coding experience**. Follow it top to bottom.

---

## What changed from the last version

- ❌ Removed: manual "LinkedIn ID", "Executive name", and "Photo upload" fields.
- ✅ The LinkedIn URL and Connections count are now read **straight from
  each Excel sheet** (no typing needed).
- ✅ One upload can contain **many sheets** (e.g. `LA1`...`LA10`) — the app
  builds slides for all of them and combines everything into a single
  downloadable PowerPoint.
- The headshot-photo box stays on every slide (just as decorative
  branding). Its name/country text is blanked out since there's no data
  source for it anymore.
- A **section-divider slide** is now inserted before each executive's run
  of slides, showing the sheet name (e.g. `LA2`) and their LinkedIn handle
  under the photo, so it's obvious in the deck where one executive ends
  and the next begins.

---

## What's in this folder

| File | What it does | Do you edit it? |
|---|---|---|
| `app.py` | The web page itself (upload button, Generate button, download button) | No |
| `ppt_generator.py` | The "brain" — reads the Excel, fills the PowerPoint | No |
| `requirements.txt` | List of Python packages the app needs | No |
| `Sample.pptx` | Your template — **you must add this file yourself** | You add it once |
| `README.md` | This guide | — |

---

## Part 1 — Update the project on GitHub

If this is your **first time** setting this up, follow all 6 steps below.
If you already have the repo from before, you only need steps 4-6 (replace
the two Python files) — you can skip straight to those.

1. Go to https://github.com and log in.
2. Click the **+** icon (top right) → **New repository** (skip if your repo
   already exists).
3. Name it something like `linkedin-ppt-generator`. Keep it **Public**
   (needed for the free deployment in Part 3). Click **Create repository**.
4. Open your repo's page. Click **Add file → Upload files**.
5. Drag in these files (download them from this chat first):
   - `app.py` *(replaces the old one — this removes the manual fields)*
   - `ppt_generator.py` *(replaces the old one — this adds multi-sheet support)*
   - `requirements.txt`
   - `Sample.pptx` (your PowerPoint template — rename your file to exactly
     `Sample.pptx`, or edit the `TEMPLATE_PATH` line in `app.py` to match
     your filename). Skip this if it's already in the repo from before.
   - GitHub will warn you that `app.py` and `ppt_generator.py` already
     exist and ask if you want to replace them — say **yes**.
6. Scroll down, click **Commit changes**.

Your updated code is now on GitHub.

---

## Part 2 — Run it on your own computer (optional, to test first)

1. Install Python from https://python.org (any 3.10+ version) if you don't
   have it. During install, tick **"Add Python to PATH"**.
2. Download all 4 files from this repo into one folder on your computer
   (e.g. `Desktop/linkedin-ppt`).
3. Open a terminal / command prompt in that folder:
   - **Windows**: open the folder in File Explorer, click the address bar,
     type `cmd`, press Enter.
   - **Mac**: open **Terminal**, type `cd ` (with a space), drag the folder
     into the terminal window, press Enter.
4. Install the required packages:
   ```
   pip install -r requirements.txt
   ```
5. Start the app:
   ```
   streamlit run app.py
   ```
6. Your browser opens automatically to `http://localhost:8501`. Upload
   your workbook (the one with all the `LA1`...`LA10`-style sheets), click
   **Generate PowerPoint**, then **Download PPTX**.

---

## Part 3 — Put it online for free (Streamlit Community Cloud)

*Skip this whole section if you already deployed the app before — since
Part 1 already pushed the updated files to the same GitHub repo, your
live app link updates itself automatically within a minute or two. Just
refresh the page.*

1. Go to https://share.streamlit.io and sign in with your GitHub account.
2. Click **Create app** (or **New app**).
3. Choose **"Deploy a public app from GitHub"**.
4. Pick:
   - **Repository**: your `linkedin-ppt-generator` repo
   - **Branch**: `main`
   - **Main file path**: `app.py`
5. Click **Deploy**. Wait 1–2 minutes.
6. You'll get a public link like `https://your-app-name.streamlit.app`.
   Bookmark it — anyone with the link can upload a workbook and download
   the finished PPTX, no coding needed.

---

## What each Excel sheet needs to look like

Based on your `LA1`...`LA10` sheets, each sheet should have:

| Cell | Content |
|---|---|
| A1 | `Linkedin URL` (label) |
| B1 | the executive's full LinkedIn profile URL |
| A2 | `Connection` (label) |
| B2 | the connections text, e.g. `500+ connections` |
| *(header row)* | `Poster`, `PostType`, `TimeAgo`, `Content`, `PostLink`, `Source`, `Headline` |
| *(rows below)* | one LinkedIn post per row |

The app automatically finds the header row by looking for the `PostType`
column, so it doesn't matter if it's row 3 or a different row — as long as
the LinkedIn URL / Connections labels are somewhere in column A above it.

- **Source** column: this cell should contain an actual **hyperlink**
  (Excel: select the cell → Insert → Link → paste the URL). That's the
  link the "Source" word in the PPT will point to.
- If a row has no hyperlink in the `Source` cell, the tool falls back to
  the plain URL in the `PostLink` column.
- Sheets with no `PostType` column anywhere in their first rows are
  skipped automatically (so a stray blank sheet won't break the run).

## What gets filled into the PowerPoint

| Excel source | Goes into PPT |
|---|---|
| Sheet's `Linkedin URL` cell | "LINKEDIN ID" box at the top (shown as the short profile handle, e.g. `jane-doe-12345/`, and clickable through to the full URL) |
| Sheet's `Connection` cell | "CONNECTIONS" box at the top |
| `PostType` | Post Type column |
| `TimeAgo` | Timeline column |
| `Headline` | Key Highlight column |
| `Content` | Details column (+ a hyperlinked **"Source"** word appended at the end, linking to the URL from the `Source` column) |

Each sheet becomes its own run of slides, in the order the sheets appear
in the workbook, with its own "(1/2)"-style counter that restarts for
every new sheet. The template's table only has 3 data rows, so the tool
automatically duplicates the slide for every extra batch of rows within a
sheet (e.g. 42 rows → 14 slides for that executive).

## Changing how many rows go on each slide

In the app, there's a **"Rows per slide"** box (defaults to 3, matching
the template). Change it there — no code editing needed. If you set it
higher than what the template table can show, you'd need to also
widen/duplicate a row in `Sample.pptx` itself; ask if you'd like help
with that.

## Troubleshooting

- **"missing expected column headers"** error → open the sheet named in
  the error and check it has the header names listed above somewhere in
  its first ~15 rows (spelling/case must match).
- **"No usable LinkedIn-activity data found in any sheet"** → none of the
  sheets had a `PostType` column — double-check the workbook structure.
- **App won't deploy on Streamlit Cloud** → double check `Sample.pptx` was
  uploaded to the GitHub repo and the filename in `app.py`'s
  `TEMPLATE_PATH` matches exactly.
- **Hyperlink missing on "Source"** → the Excel `Source` cell needs a real
  hyperlink (not just text saying "Source"), or the `PostLink` column
  needs a plain URL.