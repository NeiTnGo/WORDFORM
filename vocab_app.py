import streamlit as st
import pandas as pd
import requests
import io
import os

st.set_page_config(
    page_title="Vocabulary Manager",
    page_icon="📚",
    layout="wide",
)

LOCAL_FILE = "vocab.xlsx"
GITHUB_URL = ""  # User can paste their own GitHub raw URL

EXPECTED_COLUMNS = ["Word", "IPA", "Type", "Meaning"]
TYPE_OPTIONS = ["All", "N", "V", "Adj", "Adv"]

TYPE_COLORS = {
    "N": "#4A90D9",
    "V": "#27AE60",
    "Adj": "#E67E22",
    "Adv": "#8E44AD",
}

st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    .vocab-header {
        font-size: 2rem;
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 0.2rem;
    }
    .vocab-subheader {
        color: #7f8c8d;
        font-size: 0.95rem;
        margin-bottom: 1.5rem;
    }
    .stat-card {
        background: white;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.07);
        text-align: center;
    }
    .stat-number { font-size: 1.8rem; font-weight: 700; color: #2c3e50; }
    .stat-label { font-size: 0.8rem; color: #95a5a6; text-transform: uppercase; }
    .type-badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 600;
        color: white;
    }
    div[data-testid="stSidebarContent"] { background-color: #ffffff; }
    .sidebar-section { margin-bottom: 1.2rem; }
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_from_url(url: str) -> pd.DataFrame:
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    return pd.read_excel(io.BytesIO(resp.content))


def load_local() -> pd.DataFrame | None:
    if os.path.exists(LOCAL_FILE):
        return pd.read_excel(LOCAL_FILE)
    return None


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip() for c in df.columns]
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[EXPECTED_COLUMNS].copy()
    df["Word"] = df["Word"].astype(str).str.strip()
    df["IPA"] = df["IPA"].astype(str).str.strip().replace("nan", "")
    df["Type"] = df["Type"].astype(str).str.strip()
    df["Meaning"] = df["Meaning"].astype(str).str.strip()
    df = df[df["Word"].notna() & (df["Word"] != "") & (df["Word"] != "nan")]
    return df.reset_index(drop=True)


def type_badge(t: str) -> str:
    color = TYPE_COLORS.get(t, "#95a5a6")
    return f'<span class="type-badge" style="background:{color}">{t}</span>'


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📚 Vocab Manager")
    st.markdown("---")

    st.markdown("### 📂 Load Data")

    github_url = st.text_input(
        "GitHub Raw URL",
        value=GITHUB_URL,
        placeholder="https://raw.githubusercontent.com/...",
        help="Paste a raw GitHub link to your Excel file",
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        help="Upload will auto-save as vocab.xlsx",
    )

    st.markdown("---")
    st.markdown("### 🔍 Search & Filter")

    search_query = st.text_input("Search word or meaning", placeholder="Type to search…")
    type_filter = st.selectbox("Filter by Type", TYPE_OPTIONS)

    st.markdown("---")
    if os.path.exists(LOCAL_FILE):
        fsize = os.path.getsize(LOCAL_FILE) / 1024
        st.caption(f"💾 Local file: `vocab.xlsx` ({fsize:.1f} KB)")
    else:
        st.caption("💾 No local file saved yet")


# ── Data Loading ──────────────────────────────────────────────────────────────
df_raw = None
source_label = ""

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        with open(LOCAL_FILE, "wb") as f:
            uploaded_file.seek(0)
            f.write(uploaded_file.read())
        source_label = f"✅ Uploaded & saved: **{uploaded_file.name}**"
    except Exception as e:
        st.error(f"Failed to read uploaded file: {e}")

elif github_url.strip():
    try:
        with st.spinner("Loading from GitHub…"):
            df_raw = load_from_url(github_url.strip())
        source_label = "✅ Loaded from GitHub URL"
    except Exception as e:
        st.error(f"Failed to load from GitHub: {e}")

else:
    df_raw = load_local()
    if df_raw is not None:
        source_label = "✅ Loaded from local `vocab.xlsx`"

# ── Main Content ──────────────────────────────────────────────────────────────
st.markdown('<div class="vocab-header">📚 Vocabulary Manager</div>', unsafe_allow_html=True)
st.markdown('<div class="vocab-subheader">Browse, search, and filter your vocabulary collection</div>', unsafe_allow_html=True)

if df_raw is None:
    st.info("👈 Load a vocabulary file using the sidebar — paste a GitHub URL or upload an Excel file.")
    st.markdown("""
    **Expected Excel columns:**
    | Word | IPA | Type | Meaning |
    |------|-----|------|---------|
    | ephemeral | /ɪˈfem.ər.əl/ | Adj | Lasting for a very short time |
    | perceive | /pəˈsiːv/ | V | To become aware of something |
    """)
    st.stop()

# Normalize
df = normalize_df(df_raw)

if source_label:
    st.success(source_label)

# Stats row
total = len(df)
counts = df["Type"].value_counts()

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.markdown(f'<div class="stat-card"><div class="stat-number">{total:,}</div><div class="stat-label">Total Words</div></div>', unsafe_allow_html=True)
for i, t in enumerate(["N", "V", "Adj", "Adv"]):
    c = counts.get(t, 0)
    color = TYPE_COLORS.get(t, "#95a5a6")
    with [col2, col3, col4, col5][i]:
        st.markdown(
            f'<div class="stat-card"><div class="stat-number" style="color:{color}">{c:,}</div>'
            f'<div class="stat-label">{t}</div></div>',
            unsafe_allow_html=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# ── Filtering ─────────────────────────────────────────────────────────────────
filtered = df.copy()

if type_filter != "All":
    filtered = filtered[filtered["Type"] == type_filter]

if search_query.strip():
    q = search_query.strip().lower()
    mask = (
        filtered["Word"].str.lower().str.contains(q, na=False)
        | filtered["Meaning"].str.lower().str.contains(q, na=False)
    )
    filtered = filtered[mask]

st.markdown(f"**Showing {len(filtered):,} of {total:,} words**")

# ── Table Display ─────────────────────────────────────────────────────────────
if filtered.empty:
    st.warning("No words match your search/filter criteria.")
else:
    # Render as HTML table for badge colors & performance
    rows_html = ""
    for _, row in filtered.iterrows():
        badge = type_badge(row["Type"]) if row["Type"] in TYPE_COLORS else row["Type"]
        word = str(row["Word"])
        ipa = f'<span style="color:#7f8c8d;font-size:0.85rem">{row["IPA"]}</span>' if row["IPA"] else ""
        meaning = str(row["Meaning"])
        rows_html += f"""
        <tr>
            <td style="font-weight:600;color:#2c3e50">{word}</td>
            <td>{ipa}</td>
            <td style="text-align:center">{badge}</td>
            <td style="color:#34495e">{meaning}</td>
        </tr>
        """

    table_html = f"""
    <div style="background:white;border-radius:12px;box-shadow:0 1px 6px rgba(0,0,0,0.08);overflow:hidden;margin-top:0.5rem">
        <div style="overflow-x:auto;max-height:65vh;overflow-y:auto">
            <table style="width:100%;border-collapse:collapse;font-size:0.92rem;font-family:sans-serif">
                <thead>
                    <tr style="background:#2c3e50;color:white;position:sticky;top:0;z-index:10">
                        <th style="padding:12px 16px;text-align:left;width:18%">Word</th>
                        <th style="padding:12px 16px;text-align:left;width:18%">IPA</th>
                        <th style="padding:12px 16px;text-align:center;width:10%">Type</th>
                        <th style="padding:12px 16px;text-align:left">Meaning</th>
                    </tr>
                </thead>
                <tbody>
                    {"".join(f'<tr style="border-bottom:1px solid #f0f0f0;transition:background 0.15s">' + row + '</tr>' for row in rows_html.split('</tr>') if '<tr' in row)}
                </tbody>
            </table>
        </div>
    </div>
    """

    # Simpler, reliable approach for large tables
    display_df = filtered.copy()
    display_df.index = range(1, len(display_df) + 1)

    st.dataframe(
        display_df,
        use_container_width=True,
        height=min(600, 56 + len(filtered) * 35),
        column_config={
            "Word": st.column_config.TextColumn("Word", width="medium"),
            "IPA": st.column_config.TextColumn("IPA", width="medium"),
            "Type": st.column_config.TextColumn("Type", width="small"),
            "Meaning": st.column_config.TextColumn("Meaning", width="large"),
        },
    )

    # Download filtered results
    st.markdown("<br>", unsafe_allow_html=True)
    buf = io.BytesIO()
    filtered.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        label="⬇️ Download filtered list as Excel",
        data=buf.getvalue(),
        file_name="vocab_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
