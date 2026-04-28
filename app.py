import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import date

st.set_page_config(page_title="CMDB Pregled", layout="wide")
st.title("📊 CMDB Pregled")

# =========================
# LOAD DATA
# =========================
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("data.xlsx", dtype=str)
        return df.fillna("")
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

# =========================
# SEARCH
# =========================
st.subheader("🔎 Pretraga")

search = st.text_input(
    "Pretraga po Name / Model / SP / Inventory / Serial",
    placeholder="Unesi vrednost..."
)

filtered_df = df.copy()

if search:
    s = search.lower()
    mask = filtered_df.apply(
        lambda row: row.astype(str).str.lower().str.contains(s).any(),
        axis=1
    )
    filtered_df = filtered_df[mask]

# =========================
# FILTERS
# =========================
st.subheader("🎛️ Filteri")

c1, c2, c3 = st.columns(3)

with c1:
    type_filter = st.multiselect(
        "Type",
        sorted(df["Type"].unique()) if "Type" in df.columns else []
    )

with c2:
    project_filter = st.multiselect(
        "Project",
        sorted(df["Project"].unique()) if "Project" in df.columns else []
    )

with c3:
    state_filter = st.multiselect(
        "Deployment State",
        sorted(df["Deployment State"].unique()) if "Deployment State" in df.columns else []
    )

if type_filter and "Type" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["Type"].isin(type_filter)]

if project_filter and "Project" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["Project"].isin(project_filter)]

if state_filter and "Deployment State" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["Deployment State"].isin(state_filter)]

# =========================
# TABLE
# =========================
st.subheader(f"📦 Rezultati: {len(filtered_df)}")

st.dataframe(
    filtered_df,
    use_container_width=True,
    height=500
)

st.download_button(
    "📥 Preuzmi filtrirani CMDB",
    data=to_excel(filtered_df),
    file_name="cmdb_pregled.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =========================
# INTERNI PRENOS
# =========================
st.markdown("---")
st.subheader("🔁 Interni prenos")

if "internal_transfer" not in st.session_state:
    st.session_state.internal_transfer = False

if st.button("Interni prenos"):
    st.session_state.internal_transfer = True

if st.session_state.internal_transfer:

    transfer_type = st.selectbox(
        "Tip internog prenosa",
        ["BG → NS", "FSNIŠ → FSNS"]
    )

    if transfer_type == "BG → NS":
        broj_prenosa = "BG-NS"
        iz_magacina = "FSBG"
        uredjaj_zaduzio = "FSNS"
        file_suffix = "BG_NS"
    else:
        broj_prenosa = "FSNIS-FSNS"
        iz_magacina = "FSNIŠ"
        uredjaj_zaduzio = "FSNS"
        file_suffix = "FSNIS_FSNS"

    transfer_count = st.number_input(
        "Broj uređaja za interni prenos",
        min_value=1,
        max_value=50,
        value=1,
        key="transfer_count"
    )

    transfer_devices = []

    st.info("Za svaki uređaj unesi SP broj, inventarni broj ili serijski broj.")

    for i in range(int(transfer_count)):
        st.markdown("---")
        st.subheader(f"Uređaj za prenos {i+1}")

        search_value = st.text_input(
            "SP / Inventory / Serial",
            key=f"transfer_search_{i}"
        )

        if search_value:
            value = search_value.strip().upper()
            found = None

            for col in ["SPInventoryNumber", "InventoryNumber", "SerialNumber"]:
                if col in df.columns:
                    match = df[
                        df[col].astype(str).str.strip().str.upper() == value
                    ]

                    if not match.empty:
                        found = match.iloc[0]
                        break

            if found is None:
                st.error("Uređaj nije pronađen")
            else:
                name = found.get("Name", "")
                model = found.get("Model", "")
                sp = found.get("SPInventoryNumber", "")
                inventory = found.get("InventoryNumber", "")
                serial = found.get("SerialNumber", "")

                st.success("Uređaj pronađen")

                col_a, col_b, col_c = st.columns(3)

                with col_a:
                    st.text_input("Name", value=name, disabled=True, key=f"tr_name_{i}")
                    st.text_input("Model", value=model, disabled=True, key=f"tr_model_{i}")

                with col_b:
                    st.text_input("SPInventoryNumber", value=sp, disabled=True, key=f"tr_sp_{i}")
                    st.text_input("InventoryNumber", value=inventory, disabled=True, key=f"tr_inv_{i}")

                with col_c:
                    st.text_input("SerialNumber", value=serial, disabled=True, key=f"tr_serial_{i}")

                transfer_devices.append({
                    "Name": name,
                    "Model": model,
                    "SPInventoryNumber": sp,
                    "InventoryNumber": inventory,
                    "SerialNumber": serial
                })

    if st.button("Preuzmi otpremnicu za interni prenos"):

        if not transfer_devices:
            st.error("Nema pronađenih uređaja za prenos.")
            st.stop()

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

        for i, d in enumerate(transfer_devices):
            r = start_row + i

            set_cell(ws, f"B{r}", i + 1)
            set_cell(ws, f"C{r}", d["Name"])
            set_cell(ws, f"D{r}", d["Model"])
            set_cell(ws, f"E{r}", d["InventoryNumber"])
            set_cell(ws, f"F{r}", d["SerialNumber"])
            set_cell(ws, f"G{r}", d["SPInventoryNumber"])

        out = BytesIO()
        wb.save(out)

        st.download_button(
            "Preuzmi internu otpremnicu",
            data=out.getvalue(),
            file_name=f"interni_prenos_{file_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )