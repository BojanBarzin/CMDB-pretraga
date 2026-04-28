import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import date

st.set_page_config(page_title="CMDB Pregled", layout="wide")
st.title("📊 CMDB Pregled")

@st.cache_data
def load_data():
    try:
        return pd.read_excel("data.xlsx", dtype=str).fillna("")
    except:
        return pd.DataFrame()

df = load_data()

if df.empty:
    st.warning("data.xlsx nije pronađen ili je prazan")
    st.stop()

def set_cell(ws, cell, value):
    for merged_range in ws.merged_cells.ranges:
        if cell in merged_range:
            top_left = merged_range.start_cell.coordinate
            ws[top_left] = value
            ws[top_left].alignment = Alignment(horizontal="center", vertical="center")
            return
    ws[cell] = value
    ws[cell].alignment = Alignment(horizontal="center", vertical="center")

def to_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="CMDB")
    return output.getvalue()

st.subheader("🔎 Pretraga")

c1, c2, c3, c4 = st.columns(4)
with c1:
    f_name = st.text_input("Name")
    f_vendor = st.text_input("Vendor")
with c2:
    f_model = st.text_input("Model")
    f_type = st.text_input("Type")
with c3:
    f_sp = st.text_input("SPInventoryNumber")
    f_inv = st.text_input("InventoryNumber")
with c4:
    f_serial = st.text_input("SerialNumber")

filtered_df = df.copy()

filters = {
    "Name": f_name,
    "Vendor": f_vendor,
    "Model": f_model,
    "Type": f_type,
    "SPInventoryNumber": f_sp,
    "InventoryNumber": f_inv,
    "SerialNumber": f_serial,
}

for col, value in filters.items():
    if value and col in filtered_df.columns:
        filtered_df = filtered_df[
            filtered_df[col].astype(str).str.contains(value, case=False, na=False)
        ]

st.subheader(f"📦 Rezultati: {len(filtered_df)}")

selected_rows = []

if not filtered_df.empty:
    display_cols = [
        "Name", "Vendor", "Model", "Type",
        "SPInventoryNumber", "InventoryNumber", "SerialNumber"
    ]
    available_cols = [c for c in display_cols if c in filtered_df.columns]

    for idx, row in filtered_df.iterrows():
        label = " | ".join([f"{col}: {row.get(col, '')}" for col in available_cols])

        if st.checkbox(label, key=f"select_{idx}"):
            selected_rows.append(row)

else:
    st.info("Nema rezultata za prikaz.")

st.download_button(
    "📥 Preuzmi filtrirani CMDB",
    data=to_excel(filtered_df),
    file_name="cmdb_pregled.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

def generate_internal_transfer(selected_rows, transfer_type):
    if not selected_rows:
        st.error("Nisi izabrao nijedan uređaj.")
        st.stop()

    if transfer_type == "BG_NS":
        broj_prenosa = "BG-NS"
        iz_magacina = "FSBG"
        uredjaj_zaduzio = "FSNS"
        file_name = "interni_prenos_BG_NS.xlsx"
    else:
        broj_prenosa = "FSNIS-FSNS"
        iz_magacina = "FSNIŠ"
        uredjaj_zaduzio = "FSNS"
        file_name = "interni_prenos_NIS_NS.xlsx"

    try:
        wb = load_workbook("otpremnica_template.xlsx")
        ws = wb.active
    except:
        st.error("Nije pronađen fajl: otpremnica_template.xlsx")
        st.stop()

    set_cell(ws, "F4", broj_prenosa)
    set_cell(ws, "G5", date.today().strftime("%d.%m.%Y"))

    set_cell(ws, "B8", iz_magacina)
    set_cell(ws, "G8", uredjaj_zaduzio)

    set_cell(ws, "G9", "")
    set_cell(ws, "G10", "")
    set_cell(ws, "G11", "")

    start_row = 14

    for i, row in enumerate(selected_rows):
        r = start_row + i

        set_cell(ws, f"B{r}", i + 1)
        set_cell(ws, f"C{r}", row.get("Name", ""))
        set_cell(ws, f"D{r}", row.get("Model", ""))
        set_cell(ws, f"E{r}", row.get("InventoryNumber", ""))
        set_cell(ws, f"F{r}", row.get("SerialNumber", ""))
        set_cell(ws, f"G{r}", row.get("SPInventoryNumber", ""))

    out = BytesIO()
    wb.save(out)

    st.download_button(
        "Preuzmi internu otpremnicu",
        data=out.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.subheader("🔁 Interni prenos")

col_bg, col_nis = st.columns(2)

with col_bg:
    if st.button("BG → NS"):
        generate_internal_transfer(selected_rows, "BG_NS")

with col_nis:
    if st.button("NIŠ → NS"):
        generate_internal_transfer(selected_rows, "NIS_NS")