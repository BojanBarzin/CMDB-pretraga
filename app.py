import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import date
import streamlit.components.v1 as components

st.set_page_config(page_title="CMDB Pregled", layout="wide")
st.title("📊 CMDB Pregled")

if "transfer_list" not in st.session_state:
    st.session_state.transfer_list = []

if "generated_excel" not in st.session_state:
    st.session_state.generated_excel = None

if "generated_file_name" not in st.session_state:
    st.session_state.generated_file_name = ""

if "print_html" not in st.session_state:
    st.session_state.print_html = ""

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

def build_print_html(selected_rows, transfer_type):
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

    return f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                padding: 30px;
                color: #000;
            }}
            h2 {{
                text-align: center;
                margin-bottom: 30px;
            }}
            .info {{
                margin-bottom: 20px;
                font-size: 14px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 20px;
            }}
            th, td {{
                border: 1px solid #000;
                padding: 6px;
                text-align: center;
                font-size: 13px;
            }}
            th {{
                font-weight: bold;
            }}
            .signatures {{
                margin-top: 60px;
                display: flex;
                justify-content: space-between;
                font-size: 13px;
            }}
            .sig {{
                width: 35%;
                text-align: center;
                border-top: 1px solid #000;
                padding-top: 8px;
            }}
            @media print {{
                button {{
                    display: none;
                }}
            }}
        </style>
    </head>
    <body>
        <button onclick="window.print()" style="padding:10px 18px; margin-bottom:20px;">
            🖨️ Print
        </button>

        <h2>{naslov}</h2>

        <div class="info">
            <p><b>Datum:</b> {date.today().strftime("%d.%m.%Y")}</p>
            <p><b>Iz magacina:</b> {iz_magacina}</p>
            <p><b>Uređaj zadužio:</b> {zaduzio}</p>
        </div>

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

        <div class="signatures">
            <div class="sig">Izdao</div>
            <div class="sig">Primio</div>
        </div>
    </body>
    </html>
    """

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

    st.session_state.generated_excel = out.getvalue()
    st.session_state.generated_file_name = file_name
    st.session_state.print_html = build_print_html(selected_rows, transfer_type)

def add_selected_to_transfer(selected_rows):
    if selected_rows.empty:
        st.error("Nisi štiklirao nijedan uređaj.")
        return

    added = 0
    existing_sp = [
        x.get("SPInventoryNumber", "")
        for x in st.session_state.transfer_list
    ]

    for _, row in selected_rows.iterrows():
        sp = row.get("SPInventoryNumber", "")

        if sp and sp not in existing_sp:
            st.session_state.transfer_list.append(row.to_dict())
            existing_sp.append(sp)
            added += 1

    if added > 0:
        st.success(f"Dodato uređaja: {added}")
    else:
        st.warning("Nema novih uređaja za dodavanje.")

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

    if filtered_df.empty:
        st.info("Nema rezultata.")
    else:
        display_cols = [
            "Name", "Vendor", "Model", "Type",
            "SPInventoryNumber", "InventoryNumber", "SerialNumber"
        ]

        available_cols = [c for c in display_cols if c in filtered_df.columns]

        view_df = filtered_df[available_cols].copy()
        view_df.insert(0, "Izaberi", False)

        edited_df = st.data_editor(
            view_df,
            use_container_width=True,
            hide_index=True,
            height=350,
            key="editor",
            column_config={
                "Izaberi": st.column_config.CheckboxColumn("Izaberi")
            },
            disabled=available_cols
        )

        selected_rows = edited_df[edited_df["Izaberi"] == True].drop(columns=["Izaberi"])

        if st.button("➕ Dodaj uređaj"):
            add_selected_to_transfer(selected_rows)

        st.download_button(
            "📥 Preuzmi filtrirani CMDB",
            data=to_excel(edited_df.drop(columns=["Izaberi"])),
            file_name="cmdb_pregled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Unesi parametar za pretragu da bi se prikazali rezultati.")

# =========================
# POSLEDNJA ŠANSA
# =========================
st.markdown("---")
st.subheader("🔍 Poslednja šansa")
st.caption("Ako nisi našao uređaj pokušaj još jednom ovde")

last_chance_value = st.text_input(
    "Pretraga po svim kolonama",
    key="last_chance_search"
)

if last_chance_value:
    search_text = last_chance_value.lower()

    last_chance_df = df[
        df.apply(
            lambda row: row.astype(str).str.lower().str.contains(search_text).any(),
            axis=1
        )
    ].copy()

    st.subheader(f"📦 Rezultati poslednje šanse: {len(last_chance_df)}")

    if last_chance_df.empty:
        st.info("Nema rezultata.")
    else:
        display_cols = [
            "Name", "Vendor", "Model", "Type",
            "SPInventoryNumber", "InventoryNumber", "SerialNumber"
        ]

        available_cols = [c for c in display_cols if c in last_chance_df.columns]

        view_df = last_chance_df[available_cols].copy()
        view_df.insert(0, "Izaberi", False)

        edited_last_df = st.data_editor(
            view_df,
            use_container_width=True,
            hide_index=True,
            height=350,
            key="last_chance_editor",
            column_config={
                "Izaberi": st.column_config.CheckboxColumn("Izaberi")
            },
            disabled=available_cols
        )

        selected_last_rows = edited_last_df[
            edited_last_df["Izaberi"] == True
        ].drop(columns=["Izaberi"])

        if st.button("➕ Dodaj uređaj iz poslednje šanse"):
            add_selected_to_transfer(selected_last_rows)

# =========================
# LISTA ZA INTERNI PRENOS
# =========================
st.markdown("---")
st.subheader("🔁 Lista za interni prenos")

if st.session_state.transfer_list:
    header_cols = st.columns([2, 2, 2, 2, 2, 1])
    header_cols[0].markdown("**Name**")
    header_cols[1].markdown("**Model**")
    header_cols[2].markdown("**SP**")
    header_cols[3].markdown("**Inventory**")
    header_cols[4].markdown("**Serial**")
    header_cols[5].markdown("**Ukloni**")

    for i, row in enumerate(st.session_state.transfer_list):
        c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 2, 2, 2, 1])

        with c1:
            st.write(row.get("Name", ""))
        with c2:
            st.write(row.get("Model", ""))
        with c3:
            st.write(row.get("SPInventoryNumber", ""))
        with c4:
            st.write(row.get("InventoryNumber", ""))
        with c5:
            st.write(row.get("SerialNumber", ""))
        with c6:
            if st.button("🗑️", key=f"delete_{i}"):
                st.session_state.transfer_list.pop(i)
                st.rerun()

    st.info(f"Ukupno uređaja za prenos: {len(st.session_state.transfer_list)}")
else:
    st.info("Lista je prazna.")

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
        st.session_state.generated_excel = None
        st.session_state.generated_file_name = ""
        st.session_state.print_html = ""
        st.rerun()

# =========================
# DOWNLOAD + PRINT
# =========================
if st.session_state.generated_excel:
    st.markdown("---")
    st.subheader("📄 Dokument")

    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            "📥 Preuzmi Excel",
            data=st.session_state.generated_excel,
            file_name=st.session_state.generated_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with d2:
        if st.button("🖨️ Prikaži dokument za štampu"):
            components.html(
                st.session_state.print_html,
                height=900,
                scrolling=True
            )