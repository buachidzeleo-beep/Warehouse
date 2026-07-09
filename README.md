# Pick List & Delivery Note Generator

Streamlit app. Upload the daily produce export (same layout as the example file)
and download two ZIPs:

- **pick_list_<date>.zip** — one Word file **per vehicle** (each truck's shops broken down inside)
- **delivery_notes_<date>.zip** — one Word file **per store** (big shop name/address, small vehicle line)

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
- **მანქანის პრეფიქსით** — name delivery files `<vehicle> - <store>.docx` so a truck's
  notes group together in the ZIP; uncheck for plain `<store>.docx`.

## Files
- `app.py` — Streamlit UI
- `docgen.py` — data aggregation + Word generation (python-docx) + zipping
- `requirements.txt`
- `run_app.bat`
