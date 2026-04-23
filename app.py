import streamlit as st
import pandas as pd


# =========================
# CONFIG
# =========================
client = OpenAI()

st.set_page_config(
    page_title="CMDB Pretraga",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================
# LOAD DATA (FAST CACHE)
# =========================
@st.cache_data
def load_data():
    df = pd.read_excel("SPTS CMDB.xlsx")
    df = df.astype(str)
    return df

df = load_data()

# =========================
# MOBILE UI STYLE (FULL SCREEN FEEL)
# =========================
st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        header {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("📦 CMDB Pretraga + AI Chat")

# =========================
# TABS (Search + AI)
# =========================
tab1, tab2 = st.tabs(["🔎 Pretraga", "🧠 AI Chat"])

# ==========================================================
# 🔎 TAB 1 - INSTANT SEARCH (TYPE & SEARCH LIVE)
# ==========================================================
with tab1:

    query = st.text_input(
        "Unesi inventarni / serijski / SP / model",
        placeholder="Kucaj bilo šta..."
    )

    if query:

        result = df[
            df["InventoryNumber"].str.contains(query, case=False, na=False) |
            df["SerialNumber"].str.contains(query, case=False, na=False) |
            df["Model"].str.contains(query, case=False, na=False) |
            df["SPInventoryNumber"].str.contains(query, case=False, na=False)
        ]

        st.success(f"Pronađeno: {len(result)}")

        for _, row in result.iterrows():

            with st.container():
                st.markdown("### 📦 Uređaj")

                st.write("🏷️ Inventarni:", row["InventoryNumber"])
                st.write("🔢 Serijski:", row["SerialNumber"])
                st.write("📦 Model:", row["Model"])
                st.write("🧾 SP:", row["SPInventoryNumber"])

                st.divider()

# ==========================================================
# 🧠 TAB 2 - AI CHAT NAD CSV
# ==========================================================
with tab2:

    st.subheader("Postavi pitanje o opremi")

    user_question = st.text_input("Npr: Koji uređaji imaju isti model?")

    if user_question:

        # uzmi mali sample za kontekst (brže)
        sample_data = df.head(50).to_string()

        prompt = f"""
Ti si IT CMDB asistent.

Podaci (primer):
{sample_data}

Korisnik pitanje:
{user_question}

Odgovori jasno i konkretno na osnovu podataka.
"""

        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[
                {"role": "system", "content": "Ti si analitičar za IT opremu."},
                {"role": "user", "content": prompt}
            ]
        )

        st.write("🧠 Odgovor:")
        st.success(response.choices[0].message.content)