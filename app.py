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


def print_document_button(selected_rows, transfer_type):
    if transfer_type == "BG_NS":
        naslov = "Interni prenos BG → NS"
        iz_magacina = "FSBG"
        zaduzio = "FSNS"
    else:
        naslov = "Interni prenos NIŠ → NS"
        iz_magacina = "FSNIŠ"
        zaduzio = "FSNS"

    rows_html = ""
    for i, row in enumerate(selected_rows, start=1):
        rows_html += f"""
        <tr>
            <td>{i}</td>
            <td>{row.get("Name", "")}</td>
            <td>{row.get("Model", "")}</td>
            <td>{row.get("InventoryNumber", "")}</td>
            <td>{row.get("SerialNumber", "")}</td>
            <td>{row.get("SPInventoryNumber", "")}</td>
        </tr>
        """

    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial; padding: 30px; }}
            h2 {{ text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            td, th {{ border: 1px solid black; padding: 6px; text-align: center; }}
        </style>
    </head>
    <body>
        <h2>{naslov}</h2>
        <p><b>Datum:</b> {date.today().strftime("%d.%m.%Y")}</p>
        <p><b>Iz magacina:</b> {iz_magacina}</p>
        <p><b>Uređaj zadužio:</b> {zaduzio}</p>

        <table>
            <tr>
                <th>BR</th>
                <th>NAZIV</th>
                <th>MODEL</th>
                <th>INV</th>
                <th>SN</th>
                <th>SP/FS</th>
            </tr>
            {rows_html}
        </table>

        <script>
            window.print();
        </script>
    </body>
    </html>
    """

    st.components.v1.html(html, height=0)


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
    excel_bytes = out.getvalue()

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Preuzmi dokument",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        if st.button("🖨️ Print dokument"):
            print_document_button(selected_rows, transfer_type)


# =========================
# PRETRAGA
# =========================
st.subheader("🔎 Pretraga")

SEARCH_COLUMNS = [
    "SPInventoryNumber",
    "Name",
    "Vendor",
    "Model",
    "Type",
    "InventoryNumber",
    "SerialNumber"
]

available_search_columns = [c for c in SEARCH_COLUMNS if c in df.columns]

col_search_1, col_search_2 = st.columns([1, 2])

with col_search_1:
    search_col = st.selectbox(
        "Parametar",
        available_search_columns,
        index=available_search_columns.index("SPInventoryNumber") if "SPInventoryNumber" in available_search_columns else 0
    )

with col_search_2:
    search_value = st.text_input("Vrednost za pretragu")

if search_value:
    filtered_df = df[
        df[search_col].astype(str).str.contains(search_value, case=False, na=False)
    ].copy()

    st.subheader(f"📦 Rezultati: {len(filtered_df)}")

    if not filtered_df.empty:
        display_cols = [
            "Name","Vendor","Model","Type",
            "SPInventoryNumber","InventoryNumber","SerialNumber"
        ]

        view_df = filtered_df[display_cols].copy()
        view_df.insert(0, "Izaberi", False)

        edited_df = st.data_editor(
            view_df,
            use_container_width=True,
            hide_index=True,
            height=350,
            key="editor",
            column_config={"Izaberi": st.column_config.CheckboxColumn("Izaberi")},
            disabled=display_cols
        )

        selected_rows = edited_df[edited_df["Izaberi"]].drop(columns=["Izaberi"])

        if st.button("➕ Dodaj uređaj"):
            for _, row in selected_rows.iterrows():
                if row["SPInventoryNumber"] not in [x["SPInventoryNumber"] for x in st.session_state.transfer_list]:
                    st.session_state.transfer_list.append(row.to_dict())

# =========================
# LISTA
# =========================
st.markdown("---")
st.subheader("🔁 Lista za interni prenos")

if st.session_state.transfer_list:
    transfer_df = pd.DataFrame(st.session_state.transfer_list)
    st.dataframe(transfer_df, use_container_width=True, hide_index=True)

    remove_index = st.selectbox(
        "Ukloni uređaj",
        [""] + list(range(1, len(st.session_state.transfer_list)+1))
    )

    if st.button("Ukloni"):
        if remove_index:
            st.session_state.transfer_list.pop(remove_index-1)
            st.rerun()

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("BG → NS"):
        generate_internal_transfer(st.session_state.transfer_list, "BG_NS")

with col2:
    if st.button("NIŠ → NS"):
        generate_internal_transfer(st.session_state.transfer_list, "NIS_NS")

with col3:
    if st.button("Obriši listu"):
        st.session_state.transfer_list = []
        st.rerun()