# Pick List & Delivery Note Generator

Streamlit app. Upload the daily produce export (same layout as the example file)
and download two ZIPs:

- **pick_list_<date>.zip** — one Word file **per vehicle** (each truck's shops broken down inside)
- **delivery_notes_<date>.zip** — one Word file **per store**, grouped into per-vehicle folders (`<vehicle>/<store>.docx`); big shop name/address, small vehicle line

## Input file
The app reads 5 columns **by position** (so header text can change freely):

| Column | Field | Used for |
|---|---|---|
| A | Vehicle | pick list grouping + delivery note vehicle line |
| C | Store | delivery note / shop breakdown |
| E | Barcode | shown on both |
| F | Item (ნომენკლატურა) | shown on both |
| I | Boxes (ყუთი) | quantity shown on both |

Row 1 is treated as the header and skipped. Boxes are summed across stores for
each vehicle. Rows with a blank vehicle/store/barcode/item or non-positive boxes
are dropped and reported in the summary.

## Run (Windows)
Double-click **run_app.bat** (first run installs dependencies), or manually:

```bat
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

The app opens in your browser. Upload → check the summary → **დოკუმენტების გენერაცია**
→ download the two ZIPs.

### Optional (clean environment with conda)
```bat
conda create -n warehouse python=3.11 -y
conda activate warehouse
pip install -r requirements.txt
streamlit run app.py
```

## Options (sidebar)
- **თარიღი** — date printed on the documents (default: today).
- **ხელმოწერის ველი** — add a name/signature/date line to each delivery note.
- **მანქანის ფოლდერებში** — nest each delivery note in a folder named after its
  vehicle (`<vehicle>/<store>.docx`); uncheck for a flat list (`<vehicle> - <store>.docx`).

## Files
- `app.py` — Streamlit UI
- `docgen.py` — data aggregation + Word generation (python-docx) + zipping
- `requirements.txt`
- `run_app.bat`

## Deploy: GitHub → Streamlit Community Cloud

### 1. Put the code on GitHub
**Option A — web upload (no git needed):**
1. At github.com → **New repository** (e.g. `warehouse-picklist`). Choose **Private** for an internal tool. Create it (don't tick "Add a README" — this folder already has one).
2. On the empty repo page, click **"uploading an existing file"**, then drag in every file from this folder — `app.py`, `docgen.py`, `requirements.txt`, `README.md`, `.gitignore`, `run_app.bat` — so they sit at the repo **root**. Commit to `main`.

**Option B — git command line** (needs Git installed + a GitHub sign-in/token):
```bat
git init
git add .
git commit -m "Initial commit: warehouse pick/delivery generator"
git branch -M main
git remote add origin https://github.com/<username>/warehouse-picklist.git
git push -u origin main
```

### 2. Deploy on Streamlit Community Cloud
1. Go to **share.streamlit.io** → sign in with GitHub → authorize Streamlit (grant **private-repo** access when prompted if your repo is private).
2. Click **Create app** (upper-right) → **"Yup, I have an app."**
3. Set: **Repository** = `<username>/warehouse-picklist`, **Branch** = `main`, **Main file path** = `app.py`.
4. (Optional) choose a custom subdomain, then **Deploy**. First build takes a few minutes (it installs from `requirements.txt`).
5. Later, just push changes to `main` — the app redeploys automatically.

### 3. Restrict who can open it (internal data)
The app URL is reachable by anyone who has it. In your Streamlit workspace, open the app's **Settings → Sharing** and limit viewers to your team's email addresses. A **Private** GitHub repo hides the *source code* but does **not** restrict who can *open* the running app — the viewer setting does that.

> Deployment needs only `requirements.txt` (already included: streamlit, pandas, openpyxl, python-docx). If a build fails, open the app **logs** (right-hand panel) — it's almost always a missing package or a wrong main-file path.
