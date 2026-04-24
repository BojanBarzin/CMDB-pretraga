import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="CMDB System", layout="wide")

st.title("🔎 CMDB System")

# =========================
# LOAD DATA
# =========================
@st.cache_data
def load_data():
    if os.path.exists("data.csv"):
        return pd.read_csv("data.csv")
    elif os.path.exists("data.xlsx"):
        return pd.read_excel("data.xlsx")
    else:
        st.error("❌ Nedostaje data.csv ili data.xlsx")
        return pd.DataFrame()

df = load_data()

# =========================
# NOVI UREĐAJI FILE
# =========================
file_path = "novi_uredjaji.csv"

if os.path.exists(file_path):
    novi_df = pd.read_csv(file_path)
else:
    novi_df = pd.DataFrame(columns=[
        "Name", "Model", "SerialNumber",
        "InventoryNumber", "SPInventoryNumber",
        "Status"
    ])

# =========================
# KOMBINOVANA BAZA
# =========================
full_df = pd.concat([df, novi_df], ignore_index=True)

# =========================
# FILTERI
# =========================
st.sidebar.header("📊 Filteri")

filtered_df = full_df.copy()

for col in full_df.columns:
    unique_values = full_df[col].dropna().astype(str).unique()
    if len(unique_values) < 50:
        selected = st.sidebar.multiselect(f"{col}", unique_values)
        if selected:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected)]

# =========================
# INSTANT SEARCH (ISPRAVLJEN)
# =========================
st.subheader("⚡ Pretraga")

search = st.text_input("Kucaj za pretragu")

search_cols = [
    "Name",
    "Model",
    "SerialNumber",
    "InventoryNumber",
    "SPInventoryNumber"
]

available_cols = [c for c in search_cols if c in filtered_df.columns]

if search:
    combined = filtered_df[available_cols].fillna("").astype(str).apply(
        lambda x: " ".join(x),
        axis=1
    )

    mask = combined.str.contains(search, case=False, na=False)
    result = filtered_df[mask]
else:
    result = filtered_df

st.write(f"📄 Rezultati: {len(result)}")
st.dataframe(result, use_container_width=True)

# =========================
# EXPORT
# =========================
st.download_button(
    "📥 Export CSV",
    result.to_csv(index=False).encode("utf-8"),
    "cmdb_export.csv",
    "text/csv"
)

# =========================
# ➕ ADD DEVICE
# =========================
st.subheader("➕ Novi uređaj")

with st.form("add_device"):
    c1, c2 = st.columns(2)

    with c1:
        name = st.text_input("Name")
        model = st.text_input("Model")
        serial = st.text_input("SerialNumber")

    with c2:
        inv = st.text_input("InventoryNumber")
        sp = st.text_input("SPInventoryNumber")
        status = st.selectbox("Status", ["Aktivan", "Na servisu", "Otpisan"])

    submit = st.form_submit_button("💾 Sačuvaj")

    if submit:
        if not inv or not sp:
            st.error("❌ Inventory i SP su obavezni!")
        elif inv in full_df["InventoryNumber"].astype(str).values:
            st.error("❌ Inventory već postoji!")
        elif sp in full_df["SPInventoryNumber"].astype(str).values:
            st.error("❌ SP već postoji!")
        else:
            new_row = pd.DataFrame([{
                "Name": name,
                "Model": model,
                "SerialNumber": serial,
                "InventoryNumber": inv,
                "SPInventoryNumber": sp,
                "Status": status
            }])

            if os.path.exists(file_path):
                new_row.to_csv(file_path, mode="a", header=False, index=False)
            else:
                new_row.to_csv(file_path, index=False)

            st.success("✅ Uređaj dodat!")

# =========================
# EDIT DEVICE
# =========================
st.subheader("✏️ Edit uređaja")

if not novi_df.empty:

    selected = st.selectbox("Izaberi Inventory", novi_df["InventoryNumber"])

    row = novi_df[novi_df["InventoryNumber"] == selected].iloc[0]

    with st.form("edit_device"):
        c1, c2 = st.columns(2)

        with c1:
            name_e = st.text_input("Name", row.get("Name", ""))
            model_e = st.text_input("Model", row.get("Model", ""))
            serial_e = st.text_input("SerialNumber", row.get("SerialNumber", ""))

        with c2:
            inv_e = st.text_input("InventoryNumber", row["InventoryNumber"])
            sp_e = st.text_input("SPInventoryNumber", row["SPInventoryNumber"])
            status_e = st.selectbox(
                "Status",
                ["Aktivan", "Na servisu", "Otpisan"],
                index=["Aktivan", "Na servisu", "Otpisan"].index(row.get("Status", "Aktivan"))
            )

        save = st.form_submit_button("💾 Sačuvaj izmene")

        if save:
            temp = novi_df[novi_df["InventoryNumber"] != selected]

            if inv_e in temp["InventoryNumber"].astype(str).values:
                st.error("❌ Inventory već postoji!")
            elif sp_e in temp["SPInventoryNumber"].astype(str).values:
                st.error("❌ SP već postoji!")
            else:
                novi_df.loc[novi_df["InventoryNumber"] == selected] = [
                    name_e, model_e, serial_e,
                    inv_e, sp_e, status_e
                ]

                novi_df.to_csv(file_path, index=False)
                st.success("✅ Izmenjeno!")

else:
    st.info("Nema novih uređaja za edit.")