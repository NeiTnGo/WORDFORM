import streamlit as st
import pandas as pd
import requests
import io
import os

st.set_page_config(
    page_title="VocabVault",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LOCAL_FILE = "vocab.xlsx"
GITHUB_URL = ""

EXPECTED_COLUMNS = ["Word", "IPA", "Type", "Meaning"]
TYPE_OPTIONS = ["All", "N", "V", "Adj", "Adv"]

TYPE_COLORS = {
    "N":   {"bg": "#1a3a5c", "border": "#2e6da4", "text": "#60a8e0"},
    "V":   {"bg": "#0d3320", "border": "#1e7a40", "text": "#3dba6a"},
    "Adj": {"bg": "#3d2000", "border": "#a05a00", "text": "#f5a623"},
    "Adv": {"bg": "#2a1040", "border": "#6b3a9e", "text": "#b57bee"},
}

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Reset & Base ── */
*, *::before, *::after { box-sizing: border-box; }

html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
    background-color: #0b0d12 !important;
    color: #e2e8f0 !important;
    font-family: 'DM Sans', sans-serif !important;
}

[data-testid="stHeader"] { background: transparent !important; display: none; }
[data-testid="collapsedControl"] { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
footer { display: none !important; }
#MainMenu { display: none !important; }

/* ── Main Content Centering ── */
.main .block-container {
    max-width: 1280px !important;
    padding: 0 2rem 4rem 2rem !important;
    margin: 0 auto !important;
}

/* ── Navbar ── */
.vv-nav {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 1.1rem 0 1.4rem 0;
    border-bottom: 1px solid #1e2433;
    margin-bottom: 2rem;
    gap: 1rem;
}
.vv-logo {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-size: 1.3rem;
    font-weight: 700;
    color: #fff;
    white-space: nowrap;
    flex-shrink: 0;
}
.vv-logo-icon {
    width: 32px; height: 32px;
    background: linear-gradient(135deg, #3b82f6, #6366f1);
    border-radius: 8px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 0.95rem;
}
.vv-logo-dot { color: #3b82f6; }

/* ── Stat Cards ── */
.vv-stats {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 0.9rem;
    margin-bottom: 1.8rem;
}
.vv-stat {
    background: #13161f;
    border: 1px solid #1e2433;
    border-radius: 14px;
    padding: 1rem 1.2rem;
    text-align: center;
    transition: border-color 0.2s, transform 0.15s;
}
.vv-stat:hover { border-color: #3b82f6; transform: translateY(-2px); }
.vv-stat-num {
    font-size: 2rem;
    font-weight: 700;
    line-height: 1.1;
    letter-spacing: -0.03em;
}
.vv-stat-label {
    font-size: 0.72rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #4a5568;
    margin-top: 0.3rem;
}

/* ── Toolbar ── */
.vv-toolbar {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 1.2rem;
    flex-wrap: wrap;
}
.vv-count {
    font-size: 0.82rem;
    color: #4a5568;
    margin-bottom: 0.6rem;
    font-weight: 500;
    letter-spacing: 0.02em;
}
.vv-count span { color: #e2e8f0; font-weight: 600; }

/* ── Streamlit widget overrides ── */
div[data-testid="stTextInput"] input,
div[data-testid="stSelectbox"] select,
div[data-testid="stSelectbox"] > div > div {
    background-color: #13161f !important;
    color: #e2e8f0 !important;
    border: 1px solid #1e2433 !important;
    border-radius: 10px !important;
    font-family: 'DM Sans', sans-serif !important;
}
div[data-testid="stTextInput"] input:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,0.15) !important;
    outline: none !important;
}
div[data-testid="stTextInput"] label,
div[data-testid="stSelectbox"] label,
div[data-testid="stFileUploader"] label {
    color: #4a5568 !important;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
}
.stDownloadButton > button {
    background: #13161f !important;
    border: 1px solid #1e2433 !important;
    color: #e2e8f0 !important;
    border-radius: 10px !important;
    font-family: 'DM Sans', sans-serif !important;
    transition: border-color 0.2s, background 0.2s !important;
}
.stDownloadButton > button:hover {
    border-color: #3b82f6 !important;
    background: #1a1d2e !important;
}
.stAlert {
    background: #13161f !important;
    border: 1px solid #1e2433 !important;
    border-radius: 12px !important;
    color: #e2e8f0 !important;
}
.stSuccess {
    background: #0d2215 !important;
    border-color: #1e7a40 !important;
}
.stInfo {
    background: #0d1a33 !important;
    border-color: #2e6da4 !important;
}
div[data-testid="stFileUploader"] {
    background: #13161f !important;
    border: 1px dashed #1e2433 !important;
    border-radius: 12px !important;
}
div[data-testid="stFileUploader"]:hover {
    border-color: #3b82f6 !important;
}

/* ── Table ── */
.vv-table-wrap {
    background: #13161f;
    border: 1px solid #1e2433;
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 4px 24px rgba(0,0,0,0.35);
}
.vv-table-scroll {
    overflow-x: auto;
    max-height: 62vh;
    overflow-y: auto;
}
.vv-table-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
.vv-table-scroll::-webkit-scrollbar-track { background: #0b0d12; }
.vv-table-scroll::-webkit-scrollbar-thumb { background: #1e2433; border-radius: 3px; }
.vv-table-scroll::-webkit-scrollbar-thumb:hover { background: #2e3a50; }

table.vv-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.9rem;
}
table.vv-table thead tr {
    background: #0b0d12;
    position: sticky;
    top: 0;
    z-index: 10;
}
table.vv-table thead th {
    padding: 0.85rem 1.2rem;
    text-align: left;
    font-size: 0.7rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: #4a5568;
    border-bottom: 1px solid #1e2433;
    white-space: nowrap;
}
table.vv-table thead th.center { text-align: center; }
table.vv-table tbody tr {
    border-bottom: 1px solid #111318;
    transition: background 0.12s;
}
table.vv-table tbody tr:nth-child(even) { background: #0f1219; }
table.vv-table tbody tr:hover { background: #1a1e2d !important; }
table.vv-table tbody td {
    padding: 0.8rem 1.2rem;
    vertical-align: middle;
    line-height: 1.5;
}
.vv-word {
    font-weight: 600;
    font-size: 0.96rem;
    color: #f1f5f9;
    font-family: 'DM Sans', sans-serif;
}
.vv-ipa {
    color: #4a5568;
    font-size: 0.82rem;
    font-family: 'JetBrains Mono', monospace;
    font-style: italic;
}
.vv-meaning { color: #94a3b8; }
.vv-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 6px;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    border: 1px solid;
}
.vv-highlight {
    background: rgba(59,130,246,0.25);
    color: #93c5fd;
    border-radius: 3px;
    padding: 0 2px;
}
.vv-empty {
    text-align: center;
    padding: 4rem 2rem;
    color: #4a5568;
    font-size: 0.95rem;
}
.vv-empty-icon { font-size: 2.5rem; margin-bottom: 0.75rem; }

/* ── Info box ── */
.vv-info {
    background: #13161f;
    border: 1px dashed #1e2433;
    border-radius: 16px;
    padding: 2.5rem 2rem;
    text-align: center;
    color: #4a5568;
}
.vv-info-icon { font-size: 2.5rem; margin-bottom: 0.5rem; }
.vv-info h3 { color: #e2e8f0; font-size: 1.1rem; margin: 0.5rem 0; }
.vv-info p { font-size: 0.88rem; margin: 0.25rem 0; }
.vv-table-sample {
    display: inline-block;
    margin-top: 1.5rem;
    background: #0b0d12;
    border: 1px solid #1e2433;
    border-radius: 10px;
    overflow: hidden;
    font-size: 0.82rem;
    text-align: left;
}
.vv-table-sample table { width: 100%; border-collapse: collapse; }
.vv-table-sample th {
    background: #0b0d12;
    color: #4a5568;
    padding: 0.5rem 1rem;
    font-size: 0.7rem;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    border-bottom: 1px solid #1e2433;
}
.vv-table-sample td {
    padding: 0.45rem 1rem;
    color: #94a3b8;
    border-bottom: 1px solid #111318;
}
.vv-table-sample td:first-child { color: #e2e8f0; font-weight: 600; }

/* ── Section divider ── */
.vv-divider { height: 1px; background: #1e2433; margin: 1.5rem 0; border: none; }

/* ── Upload area label ── */
.vv-section-label {
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.09em;
    color: #4a5568;
    margin-bottom: 0.4rem;
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────────────

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


def badge_html(t: str) -> str:
    c = TYPE_COLORS.get(t)
    if c:
        return (f'<span class="vv-badge" '
                f'style="background:{c["bg"]};border-color:{c["border"]};color:{c["text"]}">'
                f'{t}</span>')
    return f'<span class="vv-badge" style="background:#1a1d24;border-color:#2e3a50;color:#4a5568">{t}</span>'


def highlight(text: str, query: str) -> str:
    if not query:
        return text
    import re
    pattern = re.compile(re.escape(query), re.IGNORECASE)
    return pattern.sub(
        lambda m: f'<mark class="vv-highlight">{m.group(0)}</mark>',
        text,
    )


# ── Top Navbar ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="vv-nav">
  <div class="vv-logo">
    <div class="vv-logo-icon">📖</div>
    Vocab<span class="vv-logo-dot">Vault</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Toolbar Row ───────────────────────────────────────────────────────────────
tc1, tc2, tc3, tc4 = st.columns([3, 1.5, 2, 1.5])

with tc1:
    search_query = st.text_input(
        "SEARCH",
        placeholder="🔍  Search word or meaning…",
        label_visibility="visible",
    )

with tc2:
    type_filter = st.selectbox(
        "FILTER BY TYPE",
        TYPE_OPTIONS,
        label_visibility="visible",
    )

with tc3:
    github_url = st.text_input(
        "GITHUB RAW URL",
        value=GITHUB_URL,
        placeholder="https://raw.githubusercontent.com/…",
        label_visibility="visible",
    )

with tc4:
    uploaded_file = st.file_uploader(
        "UPLOAD EXCEL",
        type=["xlsx", "xls"],
        label_visibility="visible",
    )

st.markdown('<hr class="vv-divider">', unsafe_allow_html=True)

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

# ── No data state ─────────────────────────────────────────────────────────────
if df_raw is None:
    st.markdown("""
    <div class="vv-info">
      <div class="vv-info-icon">📂</div>
      <h3>No vocabulary loaded</h3>
      <p>Paste a GitHub raw URL above, or upload an Excel file to get started.</p>
      <div class="vv-table-sample">
        <table>
          <thead><tr><th>Word</th><th>IPA</th><th>Type</th><th>Meaning</th></tr></thead>
          <tbody>
            <tr><td>ephemeral</td><td>/ɪˈfem.ər.əl/</td><td>Adj</td><td>Lasting for a very short time</td></tr>
            <tr><td>perceive</td><td>/pəˈsiːv/</td><td>V</td><td>To become aware of something</td></tr>
            <tr><td>resilience</td><td>/rɪˈzɪl.i.əns/</td><td>N</td><td>Ability to recover from difficulty</td></tr>
          </tbody>
        </table>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Normalize ─────────────────────────────────────────────────────────────────
df = normalize_df(df_raw)

if source_label:
    st.success(source_label)

# ── Stats Row ─────────────────────────────────────────────────────────────────
total = len(df)
counts = df["Type"].value_counts()

stat_items = [
    ("Total", f"{total:,}", "#3b82f6"),
    ("N", f"{counts.get('N', 0):,}", TYPE_COLORS["N"]["text"]),
    ("V", f"{counts.get('V', 0):,}", TYPE_COLORS["V"]["text"]),
    ("Adj", f"{counts.get('Adj', 0):,}", TYPE_COLORS["Adj"]["text"]),
    ("Adv", f"{counts.get('Adv', 0):,}", TYPE_COLORS["Adv"]["text"]),
]

cards_html = '<div class="vv-stats">'
for label, num, color in stat_items:
    cards_html += f"""
    <div class="vv-stat">
        <div class="vv-stat-num" style="color:{color}">{num}</div>
        <div class="vv-stat-label">{label}</div>
    </div>"""
cards_html += "</div>"

st.markdown(cards_html, unsafe_allow_html=True)

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

q_display = search_query.strip()
st.markdown(
    f'<div class="vv-count">Showing <span>{len(filtered):,}</span> of <span>{total:,}</span> words</div>',
    unsafe_allow_html=True,
)

# ── Table ─────────────────────────────────────────────────────────────────────
if filtered.empty:
    st.markdown("""
    <div class="vv-table-wrap">
      <div class="vv-empty">
        <div class="vv-empty-icon">🔎</div>
        No words match your search or filter.
      </div>
    </div>
    """, unsafe_allow_html=True)
else:
    rows = ""
    for _, row in filtered.iterrows():
        word_hl = highlight(str(row["Word"]), q_display)
        meaning_hl = highlight(str(row["Meaning"]), q_display)
        ipa = f'<div class="vv-ipa">{row["IPA"]}</div>' if row["IPA"] else ""
        b = badge_html(row["Type"])
        rows += f"""
        <tr>
          <td><div class="vv-word">{word_hl}</div></td>
          <td>{ipa}</td>
          <td style="text-align:center">{b}</td>
          <td><div class="vv-meaning">{meaning_hl}</div></td>
        </tr>"""

    table_html = f"""
    <div class="vv-table-wrap">
      <div class="vv-table-scroll">
        <table class="vv-table">
          <thead>
            <tr>
              <th style="width:18%">Word</th>
              <th style="width:18%">IPA</th>
              <th style="width:10%" class="center">Type</th>
              <th>Meaning</th>
            </tr>
          </thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
    </div>"""

    st.markdown(table_html, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    buf = io.BytesIO()
    filtered.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        label="⬇️  Download filtered list as Excel",
        data=buf.getvalue(),
        file_name="vocab_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
