import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import date

st.set_page_config(page_title="CMDB Pregled", layout="wide")
st.title("📊 CMDB Pregled")

# =========================
# SESSION STATE
# =========================
if "transfer_list" not in st.session_state:
    st.session_state.transfer_list = []

# =========================
# LOAD DATA
# =========================
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

# =========================
# HELPERS
# =========================
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


def generate_internal_transfer(selected_rows, transfer_type):
    if not selected_rows:
        st.error("Lista za interni prenos je prazna.")
        st.stop()

    if transfer_type == "BG_NS":
        broj_prenosa = "BG-NS"
        iz_magacina = "FSBG"
        uredjaj_zaduzio = "FSNS"
        file_name = "interni_prenos_BG_NS.xlsx"
        iz_magacina_cell = "B8"
    else:
        broj_prenosa = "FSNIS-FSNS"
        iz_magacina = "FSNIŠ"
        uredjaj_zaduzio = "FSNS"
        file_name = "interni_prenos_NIS_NS.xlsx"
        iz_magacina_cell = "C8"

    try:
        wb = load_workbook("otpremnica_template.xlsx")
        ws = wb.active
    except:
        st.error("Nije pronađen fajl: otpremnica_template.xlsx")
        st.stop()

    set_cell(ws, "F4", broj_prenosa)
    set_cell(ws, "G5", date.today().strftime("%d.%m.%Y"))

    set_cell(ws, iz_magacina_cell, iz_magacina)
    set_cell(ws, "G8", uredjaj_zaduzio)

    set_cell(ws, "G9", "")
    set_cell(ws, "G10", "")
    set_cell(ws, "G11", "")

    start_row = 14

    for i, row in enumerate(selected_rows, start=1):
        r = start_row + i - 1

        set_cell(ws, f"B{r}", i)
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

# =========================
# PRETRAGA
# =========================
st.subheader("🔎 Pretraga")

SEARCH_COLUMNS = [
    "Name",
    "Vendor",
    "Model",
    "Type",
    "SPInventoryNumber",
    "InventoryNumber",
    "SerialNumber"
]

available_search_columns = [c for c in SEARCH_COLUMNS if c in df.columns]

col_search_1, col_search_2 = st.columns([1, 2])

with col_search_1:
    search_col = st.selectbox("Parametar", available_search_columns)

with col_search_2:
    search_value = st.text_input("Vrednost za pretragu")

if search_value:
    filtered_df = df[
        df[search_col].astype(str).str.contains(search_value, case=False, na=False)
    ].copy()

    st.subheader(f"📦 Rezultati: {len(filtered_df)}")

    if filtered_df.empty:
        st.info("Nema rezultata.")
    else:
        display_cols = [
            "Name",
            "Vendor",
            "Model",
            "Type",
            "SPInventoryNumber",
            "InventoryNumber",
            "SerialNumber"
        ]

        available_cols = [c for c in display_cols if c in filtered_df.columns]

        view_df = filtered_df[available_cols].copy()
        view_df.insert(0, "Izaberi", False)

        edited_df = st.data_editor(
            view_df,
            use_container_width=True,
            hide_index=True,
            height=350,
            key="cmdb_selection_editor",
            column_config={
                "Izaberi": st.column_config.CheckboxColumn(
                    "Izaberi",
                    default=False
                )
            },
            disabled=available_cols
        )

        selected_rows = edited_df[edited_df["Izaberi"] == True].drop(columns=["Izaberi"])

        if st.button("➕ Dodaj uređaj za interni prenos"):
            if selected_rows.empty:
                st.error("Nisi štiklirao nijedan uređaj.")
            else:
                added = 0

                existing_keys = {
                    row.get("SPInventoryNumber", "")
                    for row in st.session_state.transfer_list
                }

                for _, row in selected_rows.iterrows():
                    sp = row.get("SPInventoryNumber", "")

                    if sp and sp not in existing_keys:
                        st.session_state.transfer_list.append(row.to_dict())
                        existing_keys.add(sp)
                        added += 1

                st.success(f"Dodato uređaja: {added}")

        st.download_button(
            "📥 Preuzmi filtrirani CMDB",
            data=to_excel(edited_df.drop(columns=["Izaberi"])),
            file_name="cmdb_pregled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Unesi parametar za pretragu da bi se prikazali rezultati.")

# =========================
# INTERNI PRENOS
# =========================
st.markdown("---")
st.subheader("🔁 Lista za interni prenos")

if st.session_state.transfer_list:
    transfer_df = pd.DataFrame(st.session_state.transfer_list)

    st.dataframe(
        transfer_df,
        use_container_width=True,
        hide_index=True
    )

    st.info(f"Ukupno uređaja za prenos: {len(st.session_state.transfer_list)}")
else:
    st.info("Lista je prazna.")

col_bg, col_nis, col_clear = st.columns(3)

with col_bg:
    if st.button("BG → NS"):
        generate_internal_transfer(st.session_state.transfer_list, "BG_NS")

with col_nis:
    if st.button("NIŠ → NS"):
        generate_internal_transfer(st.session_state.transfer_list, "NIS_NS")

with col_clear:
    if st.button("Obriši izbor"):
        st.session_state.transfer_list = []
        st.success("Lista obrisana")