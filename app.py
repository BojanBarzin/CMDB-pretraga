import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="CMDB Pretraga", layout="wide")

st.title("🔎 CMDB Pretraga")

# =========================
# UČITAVANJE GLAVNE BAZE
# =========================
@st.cache_data
def load_data():
    try:
        return pd.read_csv("data.csv")
    except:
        return pd.read_excel("data.xlsx")

df = load_data()

# =========================
# UČITAVANJE NOVIH UREĐAJA
# =========================
file_path = "novi_uredjaji.csv"

if os.path.exists(file_path):
    novi_df = pd.read_csv(file_path)
else:
    novi_df = pd.DataFrame(columns=[
        "InventoryNumber", "SP_Broj", "Korisnik",
        "Lokacija", "Tip", "Status"
    ])

# =========================
# SPAJANJE PODATAKA
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
# PRETRAGA
# =========================
st.subheader("🔍 Pretraga")

search = st.text_input("Pretraga po svim parametrima")
column = st.selectbox("Kolona", ["Sve"] + list(full_df.columns))

if search:
    if column == "Sve":
        mask = filtered_df.astype(str).apply(
            lambda row: row.str.contains(search, case=False, na=False).any(),
            axis=1
        )
    else:
        mask = filtered_df[column].astype(str).str.contains(search, case=False, na=False)

    result = filtered_df[mask]
else:
    result = filtered_df

st.write(f"📄 Pronađeno: {len(result)} rezultata")
st.dataframe(result, use_container_width=True)

# =========================
# UNOS NOVOG UREĐAJA
# =========================
st.subheader("➕ Novi uređaj")

with st.form("unos_forma"):
    col1, col2 = st.columns(2)

    with col1:
        inventory = st.text_input("Inventory broj")
        sp_broj = st.text_input("SP broj")
        korisnik = st.text_input("Korisnik")

    with col2:
        lokacija = st.text_input("Lokacija")
        tip = st.text_input("Tip uređaja")
        status = st.selectbox("Status", ["Aktivan", "Na servisu", "Otpisan"])

    submit = st.form_submit_button("💾 Sačuvaj")

    if submit:
        # VALIDACIJA
        if not inventory or not sp_broj:
            st.error("❌ Inventory i SP broj su obavezni!")
        elif inventory in full_df["InventoryNumber"].astype(str).values:
            st.error("❌ Inventory već postoji!")
        elif sp_broj in full_df["SP_Broj"].astype(str).values:
            st.error("❌ SP broj već postoji!")
        else:
            new_data = pd.DataFrame([{
                "InventoryNumber": inventory,
                "SP_Broj": sp_broj,
                "Korisnik": korisnik,
                "Lokacija": lokacija,
                "Tip": tip,
                "Status": status
            }])

            if os.path.exists(file_path):
                new_data.to_csv(file_path, mode='a', header=False, index=False)
            else:
                new_data.to_csv(file_path, index=False)

            st.success("✅ Uređaj dodat! Refresh stranice.")

# =========================
# EDIT POSTOJEĆIH UNOSA
# =========================
st.subheader("✏️ Izmena unosa (novi uređaji)")

if not novi_df.empty:
    selected_inv = st.selectbox(
        "Izaberi Inventory za izmenu",
        novi_df["InventoryNumber"]
    )

    edit_row = novi_df[novi_df["InventoryNumber"] == selected_inv].iloc[0]

    with st.form("edit_forma"):
        col1, col2 = st.columns(2)

        with col1:
            new_inventory = st.text_input("Inventory", edit_row["InventoryNumber"])
            new_sp = st.text_input("SP broj", edit_row["SP_Broj"])
            new_korisnik = st.text_input("Korisnik", edit_row["Korisnik"])

        with col2:
            new_lokacija = st.text_input("Lokacija", edit_row["Lokacija"])
            new_tip = st.text_input("Tip", edit_row["Tip"])
            new_status = st.selectbox(
                "Status",
                ["Aktivan", "Na servisu", "Otpisan"],
                index=["Aktivan", "Na servisu", "Otpisan"].index(edit_row["Status"])
            )

        save_edit = st.form_submit_button("💾 Sačuvaj izmene")

        if save_edit:
            # VALIDACIJA (osim trenutnog reda)
            temp_df = novi_df[novi_df["InventoryNumber"] != selected_inv]

            if new_inventory in temp_df["InventoryNumber"].astype(str).values:
                st.error("❌ Inventory već postoji!")
            elif new_sp in temp_df["SP_Broj"].astype(str).values:
                st.error("❌ SP broj već postoji!")
            else:
                novi_df.loc[novi_df["InventoryNumber"] == selected_inv] = [
                    new_inventory,
                    new_sp,
                    new_korisnik,
                    new_lokacija,
                    new_tip,
                    new_status
                ]

                novi_df.to_csv(file_path, index=False)
                st.success("✅ Izmena sačuvana! Refresh stranice.")
else:
    st.info("Nema novih uređaja za izmenu.")