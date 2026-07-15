"""
Streamlit UI — Pick List & Delivery Note generator.

Run:
    streamlit run app.py

Upload the daily produce export (same layout as the example file). Produces:
  • pick_list_<date>.zip      — one Word file per VEHICLE  (shops inside)
  • delivery_notes_<date>.zip — one Word file per STORE
"""

import io
import hashlib
from datetime import date

import pandas as pd
import streamlit as st

import docgen

st.set_page_config(page_title="Pick & Delivery Generator", page_icon="📦", layout="centered")

st.title("📦 შესაგროვებელი სია & მიღების ფურცლები")
st.caption(
    "ატვირთე დღის Excel ფაილი (ნიმუშის სტრუქტურით) → მიიღე ორი ZIP: "
    "**შესაგროვებელი სია** მანქანების მიხედვით და **მიღების ფურცლები** მაღაზიების მიხედვით. "
    "თითოეული მანქანა / მაღაზია — ცალკე Word ფაილი."
)

# ---------------------------------------------------------------- sidebar opts
with st.sidebar:
    st.header("პარამეტრები")
    run_date = st.date_input("თარიღი", value=date.today())
    sig = st.checkbox("მიღების ფურცელზე ხელმოწერის ველი", value=False)
    veh_folders = st.checkbox(
        "მიღების ფაილები დაჯგუფდეს მანქანის ფოლდერებში",
        value=True,
        help="ჩართული: თითო მანქანა → ცალკე ფოლდერი მაღაზიების ფაილებით "
             "(<მანქანა>/<მაღაზია>.docx). გამორთული: ბრტყელი სია "
             "(<მანქანა> - <მაღაზია>.docx).",
    )

date_str = run_date.strftime("%d.%m.%Y")

# ---------------------------------------------------------------- upload
uploaded = st.file_uploader("Excel ფაილი (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("ატვირთე ფაილი დასაწყებად.")
    st.stop()

file_bytes = uploaded.getvalue()
try:
    df, stats = docgen.load_and_prepare(io.BytesIO(file_bytes))
except Exception as e:  # noqa: BLE001
    st.error(f"ფაილის წაკითხვა ვერ მოხერხდა: {e}")
    st.stop()

# ---------------------------------------------------------------- summary
c1, c2, c3, c4 = st.columns(4)
c1.metric("მანქანა", stats["vehicles"])
c2.metric("მაღაზია", stats["stores"])
c3.metric("SKU", stats["skus"])
c4.metric("სულ ყუთი", f"{stats['total_boxes']:,}")

if stats["rows_dropped"]:
    st.warning(
        f"გამოტოვდა {stats['rows_dropped']} სტრიქონი (ცარიელი/არასწორი). "
        f"გამოყენებულია {stats['rows_used']} / {stats['rows_read']}."
    )
else:
    st.caption(f"წაკითხულია {stats['rows_read']} სტრიქონი — ყველა ვალიდურია.")

with st.expander("მანქანების მიმოხილვა"):
    summ = (
        df.groupby("Vehicle")
        .agg(მაღაზია=("Store", "nunique"), ყუთი=("Boxes", "sum"))
        .reset_index()
        .sort_values("Vehicle")
        .rename(columns={"Vehicle": "მანქანა"})
    )
    st.dataframe(summ, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------- generate
token = hashlib.md5(file_bytes).hexdigest() + f"|{date_str}|{sig}|{veh_folders}"

if st.button("📄 დოკუმენტების გენერაცია", type="primary", use_container_width=True):
    with st.spinner("მიმდინარეობს გენერაცია..."):
        pick_zip, npick = docgen.build_picklist_zip(df, date_str)
        del_zip, ndel = docgen.build_delivery_zip(
            df, date_str, signature=sig, group_by_vehicle=veh_folders
        )
    st.session_state["result"] = {
        "token": token, "pick": pick_zip, "del": del_zip, "npick": npick, "ndel": ndel,
    }

res = st.session_state.get("result")
if res and res["token"] == token:
    st.success(f"მზადაა: {res['npick']} შესაგროვებელი ფაილი · {res['ndel']} მიღების ფურცელი")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Pick List (ZIP)", res["pick"],
        file_name=f"pick_list_{run_date.isoformat()}.zip",
        mime="application/zip", use_container_width=True,
    )
    d2.download_button(
        "⬇️ Delivery Notes (ZIP)", res["del"],
        file_name=f"delivery_notes_{run_date.isoformat()}.zip",
        mime="application/zip", use_container_width=True,
    )
elif res:
    st.info("პარამეტრები შეიცვალა — ხელახლა დააჭირე „დოკუმენტების გენერაცია“.")
