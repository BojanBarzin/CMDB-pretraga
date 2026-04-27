import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="CMDB System", layout="wide")

st.title("🔎 CMDB Pregled")

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
# FILTERI
# =========================
st.sidebar.header("📊 Filteri")

filtered_df = df.copy()

for col in df.columns:
    unique_values = df[col].dropna().astype(str).unique()
    if len(unique_values) < 50:
        selected = st.sidebar.multiselect(f"{col}", unique_values)
        if selected:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected)]

# =========================
# INSTANT SEARCH
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

# =========================
# PRIKAZ
# =========================
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