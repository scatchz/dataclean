import json
import warnings
from datetime import datetime
from io import BytesIO

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy import stats

warnings.filterwarnings("ignore")
matplotlib.use("Agg")

try:
    import openpyxl  # noqa
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

# =========================================================
# PAGE CONFIG  — must be first Streamlit call
# =========================================================
st.set_page_config(
    page_title="DataClean",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# GOOGLE OAUTH — imports + functions (after set_page_config)
# =========================================================
try:
    from google_auth_oauthlib.flow import Flow
    from google.oauth2.credentials import Credentials
    import gspread
    GOOGLE_OK = True
except ImportError:
    GOOGLE_OK = False

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]
REDIRECT_URI = "https://dataclean-st.streamlit.app/"

def get_client_secret_file():
    """Build a temp client_secret.json from Streamlit secrets."""
    import tempfile
    secret_data = {
        "web": {
            "client_id":     st.secrets["google_oauth"]["client_id"],
            "client_secret": st.secrets["google_oauth"]["client_secret"],
            "redirect_uris": [st.secrets["google_oauth"]["redirect_uri"]],
            "auth_uri":  "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }
    tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
    json.dump(secret_data, tmp)
    tmp.close()
    return tmp.name

def get_auth_url():
    """Return the Google OAuth authorization URL and the flow object."""
    secret_file = get_client_secret_file()
    flow = Flow.from_client_secrets_file(secret_file, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    auth_url, state = flow.authorization_url(prompt="consent", access_type="offline")
    st.session_state["oauth_state"] = state
    return auth_url, flow

def exchange_code_for_token(code):
    """Exchange the OAuth callback code for a token and store it in session."""
    secret_file = get_client_secret_file()
    flow = Flow.from_client_secrets_file(
        secret_file, scopes=SCOPES, redirect_uri=REDIRECT_URI,
        state=st.session_state.get("oauth_state"),
    )
    flow.fetch_token(code=code)
    creds = flow.credentials
    st.session_state["google_token"] = {
        "token":         creds.token,
        "refresh_token": creds.refresh_token,
        "client_id":     creds.client_id,
        "client_secret": creds.client_secret,
    }

def get_google_credentials():
    """Return a Credentials object from session, or None if not signed in."""
    tok = st.session_state.get("google_token")
    if tok:
        return Credentials(
            token=tok["token"],
            refresh_token=tok["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=tok["client_id"],
            client_secret=tok["client_secret"],
            scopes=SCOPES,
        )
    return None

def load_google_sheet(sheet_url: str, creds) -> pd.DataFrame:
    """Load a Google Sheet into a DataFrame."""
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(sheet_url)
    worksheet = sh.get_active_worksheet()
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

# =========================================================
# STYLES
# =========================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
.main .block-container { padding-top: 1.2rem; max-width: 1300px; }

.app-title {
    font-family: 'JetBrains Mono', monospace;
    font-size: 1.9rem; font-weight: 600;
    background: linear-gradient(120deg, #6366f1 0%, #06b6d4 100%);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    background-clip: text; margin-bottom: .05rem;
}
.app-sub {
    font-family: 'JetBrains Mono', monospace;
    font-size: .62rem; opacity: 0.55;
    letter-spacing: 2.5px; text-transform: uppercase; margin-bottom: 1.2rem;
}
.sec-label {
    font-family: 'JetBrains Mono', monospace; font-size: .58rem;
    letter-spacing: 3px; text-transform: uppercase; color: #6366f1;
    padding-bottom: 5px; border-bottom: 1px solid rgba(128,128,128,0.18);
    margin: 1.4rem 0 .8rem 0;
}
.info-card {
    background: rgba(99,102,241,0.06); border: 1px solid rgba(99,102,241,0.22);
    border-left: 3px solid #6366f1;
    border-radius: 8px; padding: 11px 16px; margin-bottom: 10px;
    font-size: .84rem; line-height: 1.6;
}
.success-card {
    background: rgba(16,185,129,0.06); border: 1px solid rgba(16,185,129,0.22);
    border-left: 3px solid #10b981;
    border-radius: 8px; padding: 11px 16px; margin-bottom: 10px;
    font-size: .84rem; line-height: 1.6;
}
.warn-card {
    background: rgba(245,158,11,0.06); border: 1px solid rgba(245,158,11,0.22);
    border-left: 3px solid #f59e0b;
    border-radius: 8px; padding: 11px 16px; margin-bottom: 10px;
    font-size: .84rem; line-height: 1.6;
}
.error-card {
    background: rgba(239,68,68,0.06); border: 1px solid rgba(239,68,68,0.22);
    border-left: 3px solid #ef4444;
    border-radius: 8px; padding: 11px 16px; margin-bottom: 10px;
    font-size: .84rem; line-height: 1.6;
}
.impact-row {
    display: flex; gap: 0;
    border: 1px solid rgba(128,128,128,0.18);
    border-radius: 8px; overflow: hidden; margin: 8px 0 12px 0;
}
.impact-cell {
    flex: 1; padding: 8px 14px;
    border-right: 1px solid rgba(128,128,128,0.18);
    text-align: center;
}
.impact-cell:last-child { border-right: none; }
.ic-lbl { font-family:'JetBrains Mono',monospace; font-size:.57rem; text-transform:uppercase; letter-spacing:1.5px; opacity:.6; }
.ic-val { font-family:'JetBrains Mono',monospace; font-size:1rem; font-weight:600; }
.ic-arrow { opacity:.4; margin:0 3px; }
.pos { color:#10b981; font-size:.72rem; font-family:'JetBrains Mono',monospace; }
.neg { color:#ef4444; font-size:.72rem; font-family:'JetBrains Mono',monospace; }
.neu { opacity:.45; font-size:.72rem; font-family:'JetBrains Mono',monospace; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# CONSTANTS
# =========================================================
_COLORS = ["#6366f1","#06b6d4","#10b981","#f59e0b","#ef4444","#a78bfa","#fb923c","#e879f9"]
MAX_UNDO = 20

# =========================================================
# SESSION STATE
# =========================================================
def _init():
    defaults = {
        "original_df": None,
        "working_df": None,
        "history": [],
        "log": [],
        "filename": None,
        "last_msg": None,
        "last_msg_type": "info",
        "last_before": None,
        "last_after": None,
        "last_impact": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()

# =========================================================
# HELPERS — I/O
# =========================================================
@st.cache_data(show_spinner=False)
def load_file(name: str, data: bytes) -> pd.DataFrame:
    bio = BytesIO(data)
    ext = name.lower().rsplit(".", 1)[-1]
    if ext == "csv":
        for enc in ["utf-8", "utf-8-sig", "latin1", "cp1252"]:
            for sep in [",", ";", "\t", "|"]:
                try:
                    bio.seek(0)
                    df = pd.read_csv(bio, encoding=enc, sep=sep, low_memory=False)
                    if df.shape[1] > 1:
                        return df
                except Exception:
                    pass
        bio.seek(0)
        return pd.read_csv(bio, low_memory=False)
    if ext in ("xlsx", "xls"):
        return pd.read_excel(bio)
    if ext == "json":
        try:
            return pd.read_json(bio)
        except Exception:
            bio.seek(0)
            return pd.json_normalize(json.load(bio))
    raise ValueError(f"Unsupported file type: .{ext}")

def to_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8")

def to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Cleaned")
    return buf.getvalue()

# =========================================================
# HELPERS — STATE
# =========================================================
def save_undo():
    if st.session_state["working_df"] is not None:
        h = st.session_state["history"]
        if len(h) >= MAX_UNDO:
            h.pop(0)
        h.append(st.session_state["working_df"].copy())

def undo_last():
    if st.session_state["history"]:
        st.session_state["working_df"] = st.session_state["history"].pop()
        set_msg("↩ Last step undone.", "warn")
        st.session_state["last_impact"] = None
    else:
        set_msg("Nothing to undo.", "warn")

def set_msg(text, kind="info"):
    st.session_state["last_msg"] = text
    st.session_state["last_msg_type"] = kind

def log_step(action, params=None, cols=None):
    st.session_state["log"].append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "step": action,
        "params": params or {},
        "affected_columns": cols or [],
    })

def commit(before, after, msg, kind="success"):
    st.session_state["last_before"] = before.head(8)
    st.session_state["last_after"] = after.head(8)
    st.session_state["last_impact"] = {
        "rb": len(before), "ra": len(after),
        "cb": before.shape[1], "ca": after.shape[1],
        "mb": int(before.isna().sum().sum()), "ma": int(after.isna().sum().sum()),
    }
    set_msg(msg, kind)

# =========================================================
# HELPERS — ANALYSIS
# =========================================================
def num_cols(df): return df.select_dtypes(include=np.number).columns.tolist()
def cat_cols(df): return df.select_dtypes(include=["object", "category", "bool"]).columns.tolist()
def dt_cols(df):
    out = []
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            out.append(c)
        elif df[c].dtype == object:
            s = df[c].dropna().head(40)
            if len(s) > 0 and pd.to_datetime(s, errors="coerce").notna().mean() > 0.6:
                out.append(c)
    return list(dict.fromkeys(out))

def profile_df(df):
    rows = []
    for col in df.columns:
        s = df[col]
        miss = int(s.isna().sum())
        rows.append({
            "Column": col,
            "Type": str(s.dtype),
            "Missing": miss,
            "Missing %": f"{miss / max(len(df), 1) * 100:.1f}%",
            "Unique": s.nunique(dropna=True),
            "Sample": ", ".join(str(v) for v in s.dropna().unique()[:3]),
        })
    return pd.DataFrame(rows)

def clean_numeric_str(s):
    c = s.astype(str).str.replace(r"[$€£,%\s]", "", regex=True).str.strip()
    return pd.to_numeric(c.replace({"": "nan", "None": "nan", "null": "nan", "NA": "nan"}), errors="coerce")

def iqr_outliers(s):
    n = pd.to_numeric(s, errors="coerce")
    v = n.dropna()
    if v.empty:
        return pd.Series(False, index=s.index)
    q1, q3 = v.quantile(0.25), v.quantile(0.75)
    iqr = q3 - q1
    return (n < q1 - 1.5 * iqr) | (n > q3 + 1.5 * iqr)

def zscore_outliers(s, thr=3.0):
    n = pd.to_numeric(s, errors="coerce")
    v = n.dropna()
    if v.empty:
        return pd.Series(False, index=s.index)
    z = np.abs(stats.zscore(v))
    mask = pd.Series(False, index=s.index)
    mask[v.index] = z > thr
    return mask

# =========================================================
# UI HELPERS
# =========================================================
def card(text, kind="info"):
    st.markdown(f'<div class="{kind}-card">{text}</div>', unsafe_allow_html=True)

def sec(label):
    st.markdown(f'<div class="sec-label">{label}</div>', unsafe_allow_html=True)

def impact_strip(imp):
    if not imp:
        return
    rd = imp["ra"] - imp["rb"]; rdc = "neg" if rd < 0 else ("pos" if rd > 0 else "neu")
    cd = imp["ca"] - imp["cb"]; cdc = "neg" if cd < 0 else ("pos" if cd > 0 else "neu")
    md = imp["ma"] - imp["mb"]; mdc = "pos" if md <= 0 else "neg"
    st.markdown(f"""
    <div class="impact-row">
      <div class="impact-cell">
        <div class="ic-lbl">Rows</div>
        <div class="ic-val">{imp["rb"]:,}<span class="ic-arrow">→</span>{imp["ra"]:,}
        <span class="{rdc}">{rd:+,}</span></div>
      </div>
      <div class="impact-cell">
        <div class="ic-lbl">Columns</div>
        <div class="ic-val">{imp["cb"]:,}<span class="ic-arrow">→</span>{imp["ca"]:,}
        <span class="{cdc}">{cd:+,}</span></div>
      </div>
      <div class="impact-cell">
        <div class="ic-lbl">Missing Cells</div>
        <div class="ic-val">{imp["mb"]:,}<span class="ic-arrow">→</span>{imp["ma"]:,}
        <span class="{mdc}">{md:+,}</span></div>
      </div>
    </div>""", unsafe_allow_html=True)

def show_result():
    msg = st.session_state.get("last_msg")
    kind = st.session_state.get("last_msg_type", "info")
    if msg:
        card(msg, kind)
        impact_strip(st.session_state.get("last_impact"))

def show_before_after():
    b = st.session_state.get("last_before")
    a = st.session_state.get("last_after")
    if b is not None and a is not None:
        sec("Before / After Preview (first 8 rows)")
        c1, c2 = st.columns(2)
        with c1:
            st.caption("📋 Before")
            st.dataframe(b, use_container_width=True, height=200)
        with c2:
            st.caption("✅ After")
            st.dataframe(a, use_container_width=True, height=200)

def quality_bar(df):
    nc = num_cols(df); cc = cat_cols(df); dc = dt_cols(df)
    miss = int(df.isna().sum().sum())
    dups = int(df.duplicated().sum())
    cols = st.columns(7)
    cols[0].metric("Rows", f"{len(df):,}")
    cols[1].metric("Columns", f"{df.shape[1]:,}")
    cols[2].metric("Numeric", f"{len(nc)}")
    cols[3].metric("Text/Cat", f"{len(cc)}")
    cols[4].metric("Datetime", f"{len(dc)}")
    cols[5].metric("Missing Cells", f"{miss:,}")
    cols[6].metric("Duplicates", f"{dups:,}")

def theme_fig(fig):
    fig.update_layout(
        font=dict(family="Inter,sans-serif"),
        colorway=_COLORS,
        margin=dict(l=40, r=20, t=50, b=40),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    return fig

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:1.2rem;font-weight:600;'
        'background:linear-gradient(120deg,#6366f1,#06b6d4);-webkit-background-clip:text;'
        '-webkit-text-fill-color:transparent;background-clip:text;margin-bottom:.1rem;">'
        '🚀 DataClean</div>', unsafe_allow_html=True)
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.55rem;'
        'opacity:0.5;letter-spacing:2px;margin-bottom:.8rem;">STUDIO</div>', unsafe_allow_html=True)
    st.markdown("---")

    page = st.radio("Go to", [
        "📁 Upload & Overview",
        "🔧 Cleaning Studio",
        "📊 Visualization Builder",
        "📤 Export & Report",
    ], label_visibility="collapsed")

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("↩ Undo", use_container_width=True, help="Undo the last cleaning step"):
            undo_last(); st.rerun()
    with c2:
        if st.button("⟳ Reset", use_container_width=True, help="Reset session completely"):
            st.session_state.clear(); st.rerun()

    fn = st.session_state.get("filename")
    wdf = st.session_state.get("working_df")
    if fn and wdf is not None:
        st.markdown("---")
        st.caption(f"📄 **{fn}**")
        st.caption(f"`{len(wdf):,} rows × {wdf.shape[1]} cols`")
        st.caption(f"Undo history: {len(st.session_state['history'])}/{MAX_UNDO}")


# =========================================================
# PAGE A — UPLOAD & OVERVIEW
# =========================================================
if page == "📁 Upload & Overview":
    st.markdown('<div class="app-title">DataClean Studio</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-sub">Upload · Profile · Clean · Visualize · Export</div>', unsafe_allow_html=True)

    card("Upload a <b>CSV, Excel (.xlsx/.xls), or JSON</b> file to get started. The app will profile your data and guide you through cleaning.", "info")

    # ── Google Sheets Section ──────────────────────────
    st.markdown("---")
    st.markdown("#### 🔗 Connect Google Sheets (optional)")

    if not GOOGLE_OK:
        card("Google Sheets libraries not installed. Run: <code>pip install google-auth google-auth-oauthlib gspread</code>", "warn")
    elif "google_oauth" not in st.secrets:
        card("Google OAuth not configured. Add <code>[google_oauth]</code> credentials to your Streamlit secrets to enable this feature.", "warn")
    else:
        # Handle OAuth callback — Google redirects back here with ?code=...
        query_params = st.query_params
        if "code" in query_params and "google_token" not in st.session_state:
            with st.spinner("Completing Google sign-in…"):
                try:
                    exchange_code_for_token(query_params["code"])
                    st.query_params.clear()
                    st.success("✓ Connected to Google!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Google auth failed: {e}")

        creds = get_google_credentials()

        if creds is None:
            # Not signed in — show Sign In button
            try:
                auth_url, _ = get_auth_url()
                st.link_button("🔐 Sign in with Google", auth_url)
                st.caption("You will be redirected to Google and then back to this app automatically.")
            except Exception as e:
                st.error(f"Could not generate auth URL: {e}")
        else:
            # Signed in — show sheet loader
            st.success("✓ Google account connected")
            if st.button("Disconnect Google account"):
                del st.session_state["google_token"]
                st.rerun()

            sheet_url = st.text_input(
                "Paste your Google Sheet URL:",
                placeholder="https://docs.google.com/spreadsheets/d/...",
            )
            if sheet_url and st.button("📥 Load Google Sheet"):
                with st.spinner("Loading sheet…"):
                    try:
                        gdf = load_google_sheet(sheet_url, creds)
                        st.session_state.update({
                            "original_df":    gdf.copy(),
                            "working_df":     gdf.copy(),
                            "history": [], "log": [],
                            "filename":       "Google Sheet",
                            "last_msg":       f"✓ Loaded Google Sheet — {len(gdf):,} rows × {gdf.shape[1]} columns.",
                            "last_msg_type":  "success",
                            "last_before": None, "last_after": None, "last_impact": None,
                        })
                        log_step("load_google_sheet", {"url": sheet_url}, list(gdf.columns))
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to load sheet: {e}")

    # ── File Upload Section ────────────────────────────
    st.markdown("---")
    uploaded = st.file_uploader(
        "Drop your dataset here",
        type=["csv", "xlsx", "xls", "json"],
        label_visibility="collapsed",
    )

    if uploaded:
        with st.spinner("Loading…"):
            try:
                raw = uploaded.getvalue()
                df = load_file(uploaded.name, raw)
                st.session_state.update({
                    "original_df":   df.copy(),
                    "working_df":    df.copy(),
                    "history": [], "log": [],
                    "filename":      uploaded.name,
                    "last_msg":      f"✓ Loaded '{uploaded.name}' — {len(df):,} rows × {df.shape[1]} columns.",
                    "last_msg_type": "success",
                    "last_before": None, "last_after": None, "last_impact": None,
                })
                log_step("upload_file", {"file": uploaded.name, "rows": len(df), "cols": df.shape[1]}, list(df.columns))
            except Exception as e:
                card(f"Failed to load file: {e}", "error")

    df = st.session_state["working_df"]
    show_result()

    if df is not None:
        sec("Data Quality Snapshot")
        quality_bar(df)

        tab1, tab2, tab3, tab4 = st.tabs(["👁 Preview", "📋 Column Profile", "📈 Summary Stats", "🔍 Missing Values"])

        with tab1:
            n = st.slider("Rows to show", 5, min(500, len(df)), min(20, len(df)))
            st.dataframe(df.head(n), use_container_width=True)

        with tab2:
            st.dataframe(profile_df(df), use_container_width=True, height=420)

        with tab3:
            try:
                st.dataframe(df.describe(include="all").T, use_container_width=True)
            except Exception:
                st.info("Could not compute statistics.")

        with tab4:
            miss = df.isna().sum()
            miss = miss[miss > 0].sort_values(ascending=False)
            if miss.empty:
                card("No missing values found — dataset is complete!", "success")
            else:
                mdf = miss.reset_index()
                mdf.columns = ["Column", "Count"]
                mdf["Pct"] = (mdf["Count"] / len(df) * 100).round(1)
                fig = px.bar(mdf, x="Pct", y="Column", orientation="h",
                             text="Count", color="Pct",
                             color_continuous_scale=["#6366f1", "#ef4444"],
                             title="Missing Values by Column (%)")
                fig = theme_fig(fig)
                fig.update_traces(textposition="outside")
                st.plotly_chart(fig, use_container_width=True, theme="streamlit")
    else:
        card("⬆ Upload a dataset above to begin.", "info")


# =========================================================
# PAGE B — CLEANING STUDIO
# =========================================================
elif page == "🔧 Cleaning Studio":
    st.markdown('<div class="app-title" style="font-size:1.6rem;">Cleaning Studio</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-sub">Apply transformations step by step</div>', unsafe_allow_html=True)

    df = st.session_state["working_df"]
    if df is None:
        st.warning("Upload a dataset first (Upload & Overview page)."); st.stop()

    quality_bar(df)
    st.markdown("---")
    show_result()
    show_before_after()
    st.markdown("---")

    card("Suggested order: ① Missing Values → ② Duplicates → ③ Column Types → ④ Text & Categories → ⑤ Outliers → ⑥ Scale → ⑦ Column Ops → ⑧ Validate", "info")

    tool = st.selectbox("Select a cleaning tool:", [
        "① Handle Missing Values",
        "② Remove Duplicates",
        "③ Convert Column Types",
        "④ Clean Text & Categories",
        "⑤ Handle Outliers",
        "⑥ Scale / Normalize Columns",
        "⑦ Column Operations",
        "⑧ Data Validation",
    ])

    # ──────────────────────────────────────────────
    # ① MISSING VALUES
    # ──────────────────────────────────────────────
    if tool == "① Handle Missing Values":
        mc = df.isna().sum()
        miss_df = pd.DataFrame({
            "Column": df.columns,
            "Missing": mc.values,
            "Missing %": (mc / len(df) * 100).round(1).values,
        })
        has_miss = miss_df[miss_df["Missing"] > 0]
        total_miss = int(mc.sum())

        st.markdown(f"""
        <div class="impact-row">
          <div class="impact-cell"><div class="ic-lbl">Total Rows</div><div class="ic-val">{len(df):,}</div></div>
          <div class="impact-cell"><div class="ic-lbl">Missing Cells</div><div class="ic-val">{total_miss:,}</div></div>
          <div class="impact-cell"><div class="ic-lbl">Columns with Gaps</div><div class="ic-val">{len(has_miss)}</div></div>
        </div>""", unsafe_allow_html=True)

        if has_miss.empty:
            card("No missing values — nothing to do here.", "success")
            st.stop()

        st.dataframe(has_miss.set_index("Column"), use_container_width=True)

        action = st.radio("Action:", [
            "Drop rows with missing values",
            "Drop columns above missing % threshold",
            "Fill / impute a column",
        ])

        if action == "Drop rows with missing values":
            cols_sel = st.multiselect("Check these columns (leave blank = any column with missing)", df.columns.tolist())
            preview_n = int(df[cols_sel].isna().any(axis=1).sum()) if cols_sel else int(df.isna().any(axis=1).sum())
            st.caption(f"Rows to remove: **{preview_n:,}** → leaving **{len(df) - preview_n:,}** rows")
            if st.button("✓ Drop Rows"):
                before = df.copy(); save_undo()
                after = df.dropna(subset=cols_sel or None).reset_index(drop=True)
                st.session_state["working_df"] = after
                log_step("drop_rows_missing", {"cols": cols_sel or "all"}, cols_sel or list(df.columns))
                commit(before, after, f"Removed {len(before) - len(after):,} rows with missing values.")
                st.rerun()

        elif action == "Drop columns above missing % threshold":
            thr = st.slider("Remove column if missing % exceeds:", 0, 100, 50)
            to_drop = df.isna().mean().mul(100).pipe(lambda s: s[s > thr]).index.tolist()
            st.caption(f"Columns that will be dropped ({len(to_drop)}): `{'`, `'.join(to_drop) if to_drop else 'none'}`")
            if st.button("✓ Drop Columns"):
                if not to_drop:
                    card("No columns exceed this threshold.", "warn"); st.stop()
                before = df.copy(); save_undo()
                after = df.drop(columns=to_drop)
                st.session_state["working_df"] = after
                log_step("drop_cols_missing", {"threshold_%": thr}, to_drop)
                commit(before, after, f"Dropped {len(to_drop)} column(s) with >{thr}% missing values.")
                st.rerun()

        elif action == "Fill / impute a column":
            col = st.selectbox("Column to fill:", df.columns.tolist())
            nm = int(df[col].isna().sum())
            st.caption(f"Missing in **{col}**: {nm:,} ({nm / len(df) * 100:.1f}%)")

            is_num = pd.api.types.is_numeric_dtype(df[col])
            if is_num:
                method_opts = ["Mean", "Median", "Constant value", "Forward Fill", "Backward Fill", "Interpolate"]
            else:
                method_opts = ["Mode (most frequent)", "Constant value", "Forward Fill", "Backward Fill"]

            method = st.selectbox("Fill method:", method_opts)
            const_val = ""
            if method == "Constant value":
                const_val = st.text_input("Enter constant:")

            if st.button("✓ Fill Missing Values"):
                before = df.copy(); save_undo(); new_df = df.copy()
                try:
                    s = new_df[col]
                    if method == "Mean":
                        new_df[col] = pd.to_numeric(s, errors="coerce").fillna(pd.to_numeric(s, errors="coerce").mean())
                    elif method == "Median":
                        new_df[col] = pd.to_numeric(s, errors="coerce").fillna(pd.to_numeric(s, errors="coerce").median())
                    elif method == "Mode (most frequent)":
                        m = s.mode(dropna=True)
                        new_df[col] = s.fillna(m.iloc[0] if not m.empty else np.nan)
                    elif method == "Forward Fill":
                        new_df[col] = s.ffill()
                    elif method == "Backward Fill":
                        new_df[col] = s.bfill()
                    elif method == "Interpolate":
                        new_df[col] = pd.to_numeric(s, errors="coerce").interpolate()
                    elif method == "Constant value":
                        try:
                            fv = float(const_val) if const_val.replace(".", "", 1).lstrip("-").isdigit() else const_val
                        except Exception:
                            fv = const_val
                        new_df[col] = s.fillna(fv)
                    filled = int(before[col].isna().sum() - new_df[col].isna().sum())
                    st.session_state["working_df"] = new_df
                    log_step("fill_missing", {"method": method, "col": col}, [col])
                    commit(before, new_df, f"Filled {filled:,} cells in '{col}' using {method}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    # ──────────────────────────────────────────────
    # ② DUPLICATES
    # ──────────────────────────────────────────────
    elif tool == "② Remove Duplicates":
        total_dups = int(df.duplicated().sum())
        st.markdown(f"""
        <div class="impact-row">
          <div class="impact-cell"><div class="ic-lbl">Total Rows</div><div class="ic-val">{len(df):,}</div></div>
          <div class="impact-cell"><div class="ic-lbl">Full Duplicates</div><div class="ic-val">{total_dups:,}</div></div>
          <div class="impact-cell"><div class="ic-lbl">Dup %</div><div class="ic-val">{total_dups / max(len(df), 1) * 100:.1f}%</div></div>
        </div>""", unsafe_allow_html=True)

        subset = st.multiselect("Check duplicate by these columns only (blank = all columns):", df.columns.tolist())
        keep_opt = st.radio("Which copy to keep?", ["First occurrence", "Last occurrence", "Remove ALL copies"], horizontal=True)
        keep_map = {"First occurrence": "first", "Last occurrence": "last", "Remove ALL copies": False}

        dup_preview = df[df.duplicated(subset=subset or None, keep=False)]
        if not dup_preview.empty:
            card(f"<b>{len(dup_preview):,} rows</b> belong to duplicate groups. Preview:", "warn")
            st.dataframe(dup_preview.head(100), use_container_width=True, height=220)
        else:
            card("No duplicates found with this column selection.", "success")

        if st.button("✓ Remove Duplicates"):
            before = df.copy(); save_undo()
            after = df.drop_duplicates(subset=subset or None, keep=keep_map[keep_opt]).reset_index(drop=True)
            st.session_state["working_df"] = after
            log_step("remove_duplicates", {"keep": keep_opt, "subset": subset or "all"}, subset)
            commit(before, after, f"Removed {len(before) - len(after):,} duplicate rows.")
            st.rerun()

    # ──────────────────────────────────────────────
    # ③ CONVERT COLUMN TYPES
    # ──────────────────────────────────────────────
    elif tool == "③ Convert Column Types":
        col = st.selectbox("Column to convert:", df.columns.tolist())
        cur_type = str(df[col].dtype)
        nm = int(df[col].isna().sum())
        sample = ", ".join(str(v) for v in df[col].dropna().unique()[:5])
        st.caption(f"Current type: `{cur_type}` | Missing: {nm:,} | Sample: {sample}")

        target = st.selectbox("Convert to:", [
            "Numeric (float)", "Integer", "Text (string)", "Category", "Date / Time", "Boolean"
        ])
        dt_fmt = ""
        if target == "Date / Time":
            dt_fmt = st.text_input("Date format (leave blank = auto-detect)", placeholder="%Y-%m-%d")

        def try_convert(s, tgt, fmt=""):
            if tgt == "Numeric (float)":  return clean_numeric_str(s)
            if tgt == "Integer":          return pd.to_numeric(clean_numeric_str(s), errors="coerce").round().astype("Int64")
            if tgt == "Category":         return s.astype("category")
            if tgt == "Text (string)":    return s.astype(str)
            if tgt == "Date / Time":
                kw = {"format": fmt} if fmt.strip() else {}
                return pd.to_datetime(s, errors="coerce", **kw)
            if tgt == "Boolean":
                return s.astype(str).str.lower().str.strip().isin({"true", "1", "yes", "y"})

        preview = try_convert(df[col], target, dt_fmt)
        if preview is not None:
            new_nulls = int(preview.isna().sum()) - nm
            pct = new_nulls / max(len(df), 1) * 100
            if new_nulls == 0:
                card("✓ Conversion preview looks clean — no new null values.", "success")
            elif pct < 10:
                card(f"⚠ {new_nulls:,} value(s) ({pct:.1f}%) cannot convert and will become null.", "warn")
            else:
                card(f"✗ {new_nulls:,} value(s) ({pct:.1f}%) cannot convert — high failure rate. Clean the column first.", "error")

        if st.button("✓ Convert"):
            before = df.copy(); save_undo(); new_df = df.copy()
            result = try_convert(new_df[col], target, dt_fmt)
            if result is None:
                st.error("Conversion failed."); st.stop()
            new_df[col] = result
            st.session_state["working_df"] = new_df
            log_step("convert_type", {"to": target, "col": col}, [col])
            commit(before, new_df, f"Converted '{col}' from {cur_type} to {target}.")
            st.rerun()

    # ──────────────────────────────────────────────
    # ④ CLEAN TEXT & CATEGORIES
    # ──────────────────────────────────────────────
    elif tool == "④ Clean Text & Categories":
        cc = cat_cols(df)
        if not cc:
            card("No text or category columns found.", "warn"); st.stop()

        col = st.selectbox("Column to clean:", cc)
        vc = df[col].value_counts(dropna=False)
        n_unique = df[col].nunique()
        st.caption(f"Unique values: {n_unique:,} | Most common: '{vc.index[0] if len(vc) > 0 else '—'}' ({vc.iloc[0]:,}×)")
        st.dataframe(vc.head(15).reset_index().rename(columns={"index": col, col: "Count"}), use_container_width=True, height=180)

        action = st.radio("Action:", [
            "Standardize case & trim whitespace",
            "Replace / remap values",
            "Group rare categories → 'Other'",
            "One-hot encode this column",
        ])

        if action == "Standardize case & trim whitespace":
            style = st.selectbox("Case style:", ["lowercase", "UPPERCASE", "Title Case", "Sentence case"])
            if st.button("✓ Standardize"):
                before = df.copy(); save_undo(); new_df = df.copy()
                s = new_df[col].astype(str).str.strip()
                s = {"lowercase": s.str.lower, "UPPERCASE": s.str.upper,
                     "Title Case": s.str.title, "Sentence case": s.str.capitalize}[style]()
                new_df[col] = s
                st.session_state["working_df"] = new_df
                log_step("standardize_text", {"style": style}, [col])
                commit(before, new_df, f"Standardized '{col}' to {style}. Unique: {before[col].nunique()} → {new_df[col].nunique()}")
                st.rerun()

        elif action == "Replace / remap values":
            top = vc.head(20).index.tolist()
            card("Edit the 'New Value' column to remap entries. Leave unchanged rows as-is.", "info")
            edited = st.data_editor(
                pd.DataFrame([{"Old Value": str(v), "New Value": str(v)} for v in top]),
                num_rows="dynamic", use_container_width=True,
                column_config={
                    "Old Value": st.column_config.TextColumn("Old Value (current)", disabled=True),
                    "New Value": st.column_config.TextColumn("New Value (replacement)"),
                },
            )
            unmatched = st.checkbox("Set values NOT in this list → 'Other'")
            if st.button("✓ Apply Mapping"):
                mapping = {r["Old Value"]: r["New Value"] for _, r in edited.iterrows()
                           if pd.notna(r["Old Value"]) and str(r["Old Value"]) != str(r["New Value"])}
                before = df.copy(); save_undo(); new_df = df.copy()
                if unmatched:
                    new_df[col] = new_df[col].map(lambda x: mapping.get(str(x), "Other") if pd.notna(x) else x)
                else:
                    new_df[col] = new_df[col].replace(mapping)
                st.session_state["working_df"] = new_df
                log_step("remap_values", {"n_mappings": len(mapping), "unmatched_to_other": unmatched}, [col])
                commit(before, new_df, f"Applied {len(mapping)} remapping(s) in '{col}'.")
                st.rerun()

        elif action == "Group rare categories → 'Other'":
            thr = st.slider("Minimum count to keep (below this → 'Other'):", 1, 500, 10)
            rare = vc[vc < thr].index.tolist()
            st.caption(f"Values that will become 'Other': **{len(rare)}**")
            if rare:
                with st.expander(f"Show {min(len(rare), 50)} values"):
                    st.write(", ".join(str(r) for r in rare[:50]) + ("…" if len(rare) > 50 else ""))
            if st.button("✓ Group Rare Values"):
                before = df.copy(); save_undo(); new_df = df.copy()
                new_df[col] = new_df[col].replace(rare, "Other")
                st.session_state["working_df"] = new_df
                log_step("group_rare", {"threshold": thr}, [col])
                commit(before, new_df, f"Grouped {len(rare)} rare values into 'Other' in '{col}'.")
                st.rerun()

        elif action == "One-hot encode this column":
            n_unique_enc = df[col].nunique()
            if n_unique_enc > 30:
                card(f"⚠ This column has {n_unique_enc} unique values — one-hot encoding will add {n_unique_enc} new columns. Consider grouping rare values first.", "warn")
            drop_orig = st.checkbox("Drop original column after encoding", value=True)
            drop_first = st.checkbox("Drop first dummy column (avoids multicollinearity)", value=False)
            st.caption(f"Will create {n_unique_enc} new binary columns.")
            if st.button("✓ One-Hot Encode"):
                before = df.copy(); save_undo(); new_df = df.copy()
                try:
                    dummies = pd.get_dummies(new_df[col].astype(str), prefix=col, drop_first=drop_first, dtype=int)
                    if drop_orig:
                        new_df = new_df.drop(columns=[col])
                    new_df = pd.concat([new_df, dummies], axis=1)
                    st.session_state["working_df"] = new_df
                    log_step("one_hot_encode", {"col": col, "drop_orig": drop_orig}, list(dummies.columns))
                    commit(before, new_df, f"One-hot encoded '{col}' → created {len(dummies.columns)} binary columns.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Encoding failed: {e}")

    # ──────────────────────────────────────────────
    # ⑤ HANDLE OUTLIERS
    # ──────────────────────────────────────────────
    elif tool == "⑤ Handle Outliers":
        nc = num_cols(df)
        if not nc:
            card("No numeric columns found.", "warn"); st.stop()

        col = st.selectbox("Numeric column:", nc)
        method = st.radio("Detection method:", ["IQR (recommended)", "Z-Score"], horizontal=True)
        thr = 3.0
        if method == "Z-Score":
            thr = st.slider("Z-Score threshold:", 1.5, 5.0, 3.0, 0.1)

        s = pd.to_numeric(df[col], errors="coerce")
        mask = iqr_outliers(df[col]) if method == "IQR (recommended)" else zscore_outliers(df[col], thr)
        n_out = int(mask.sum())

        st.markdown(f"""
        <div class="impact-row">
          <div class="impact-cell"><div class="ic-lbl">Outliers Found</div><div class="ic-val">{n_out:,}</div></div>
          <div class="impact-cell"><div class="ic-lbl">Outlier %</div><div class="ic-val">{n_out / max(len(df), 1) * 100:.1f}%</div></div>
          <div class="impact-cell"><div class="ic-lbl">Range</div><div class="ic-val" style="font-size:.85rem">{s.min():.3g} → {s.max():.3g}</div></div>
        </div>""", unsafe_allow_html=True)

        fig_m, ax = plt.subplots(figsize=(8, 2.5))
        clean = s.dropna()
        ax.boxplot(clean, vert=False, patch_artist=True,
                   boxprops=dict(facecolor="#6366f120", color="#6366f1"),
                   medianprops=dict(color="#06b6d4", linewidth=2),
                   whiskerprops=dict(color="#6366f1"),
                   capprops=dict(color="#6366f1"),
                   flierprops=dict(marker="o", color="#ef4444", alpha=0.6, markersize=4))
        ax.set_xlabel(col)
        ax.set_title(f"Box Plot — {col}", fontsize=11, fontweight="bold")
        ax.spines[["top", "right", "left"]].set_visible(False)
        ax.tick_params(left=False)
        plt.tight_layout()
        st.pyplot(fig_m, use_container_width=True)
        plt.close(fig_m)

        action = st.selectbox("What to do with outliers:", [
            "Inspect only (no changes)",
            "Remove outlier rows",
            "Cap at quantiles (Winsorize)",
            "Replace outliers with NaN",
        ])

        lower_q, upper_q = 0.01, 0.99
        if action == "Cap at quantiles (Winsorize)":
            c1, c2 = st.columns(2)
            lower_q = c1.slider("Lower cap (quantile)", 0.0, 0.2, 0.01, 0.005)
            upper_q = c2.slider("Upper cap (quantile)", 0.8, 1.0, 0.99, 0.005)

        if action != "Inspect only (no changes)" and st.button("✓ Apply"):
            before = df.copy(); save_undo(); new_df = df.copy()
            sv = pd.to_numeric(new_df[col], errors="coerce")
            if action == "Remove outlier rows":
                new_df = new_df[~mask].reset_index(drop=True)
            elif action == "Cap at quantiles (Winsorize)":
                lo, hi = sv.quantile(lower_q), sv.quantile(upper_q)
                new_df[col] = sv.clip(lower=lo, upper=hi)
            elif action == "Replace outliers with NaN":
                new_df.loc[mask, col] = np.nan
            st.session_state["working_df"] = new_df
            log_step("handle_outliers", {"method": method, "action": action, "col": col}, [col])
            commit(before, new_df, f"Handled {n_out:,} outliers in '{col}' — {action}.")
            st.rerun()

    # ──────────────────────────────────────────────
    # ⑥ SCALE / NORMALIZE
    # ──────────────────────────────────────────────
    elif tool == "⑥ Scale / Normalize Columns":
        nc = num_cols(df)
        if not nc:
            card("No numeric columns found.", "warn"); st.stop()

        cols_sel = st.multiselect("Columns to scale:", nc)
        method = st.selectbox("Scaling method:", [
            "Min-Max (0 to 1)",
            "Z-Score (mean=0, std=1)",
            "Robust (median-IQR)",
        ])
        suffix = st.text_input("Add suffix to scaled column name (blank = overwrite original):", value="")

        def scale(s, m):
            n = pd.to_numeric(s, errors="coerce")
            if m == "Min-Max (0 to 1)":
                mn, mx = n.min(), n.max()
                return n if (pd.isna(mn) or mn == mx) else (n - mn) / (mx - mn)
            if m == "Z-Score (mean=0, std=1)":
                mean, std = n.mean(), n.std()
                return n if (pd.isna(std) or std == 0) else (n - mean) / std
            if m == "Robust (median-IQR)":
                med = n.median()
                q1, q3 = n.quantile(0.25), n.quantile(0.75)
                iqr = q3 - q1
                return n if iqr == 0 else (n - med) / iqr

        if cols_sel:
            preview_data = []
            for c in cols_sel:
                orig_s = pd.to_numeric(df[c], errors="coerce")
                scaled_s = scale(df[c], method)
                preview_data.append({
                    "Column": c,
                    "Before Min": round(float(orig_s.min()), 3),
                    "Before Max": round(float(orig_s.max()), 3),
                    "Before Mean": round(float(orig_s.mean()), 3),
                    "After Min": round(float(scaled_s.min()), 3),
                    "After Max": round(float(scaled_s.max()), 3),
                    "After Mean": round(float(scaled_s.mean()), 3),
                })
            st.caption("Preview of scaling effect:")
            st.dataframe(pd.DataFrame(preview_data), use_container_width=True)

        if st.button("✓ Scale Columns"):
            if not cols_sel:
                st.warning("Select at least one column."); st.stop()
            before = df.copy(); save_undo(); new_df = df.copy()
            for c in cols_sel:
                dest = f"{c}{suffix}" if suffix else c
                new_df[dest] = scale(new_df[c], method)
            st.session_state["working_df"] = new_df
            log_step("scale", {"method": method, "suffix": suffix}, cols_sel)
            commit(before, new_df, f"Scaled {len(cols_sel)} column(s) using {method}.")
            st.rerun()

    # ──────────────────────────────────────────────
    # ⑦ COLUMN OPERATIONS
    # ──────────────────────────────────────────────
    elif tool == "⑦ Column Operations":
        action = st.radio("Action:", [
            "Rename a column",
            "Drop columns",
            "Create computed column",
            "Bin numeric column into categories",
        ])

        if action == "Rename a column":
            old = st.selectbox("Column to rename:", df.columns.tolist())
            new_name = st.text_input("New name:")
            if st.button("✓ Rename"):
                if not new_name.strip():
                    st.warning("Enter a new name."); st.stop()
                if new_name.strip() in df.columns:
                    st.warning("That name already exists."); st.stop()
                before = df.copy(); save_undo()
                after = df.rename(columns={old: new_name.strip()})
                st.session_state["working_df"] = after
                log_step("rename_column", {"from": old, "to": new_name.strip()}, [old])
                commit(before, after, f"Renamed '{old}' → '{new_name.strip()}'.")
                st.rerun()

        elif action == "Drop columns":
            drop = st.multiselect("Columns to remove:", df.columns.tolist())
            st.caption(f"Columns remaining after drop: {df.shape[1] - len(drop)}")
            if st.button("✓ Drop"):
                if not drop:
                    st.warning("Select at least one column."); st.stop()
                before = df.copy(); save_undo()
                after = df.drop(columns=drop)
                st.session_state["working_df"] = after
                log_step("drop_columns", {"dropped": drop}, drop)
                commit(before, after, f"Dropped {len(drop)} column(s).")
                st.rerun()

        elif action == "Create computed column":
            nc = num_cols(df)
            if not nc:
                card("No numeric columns.", "warn"); st.stop()
            new_col = st.text_input("New column name:")
            formula = st.selectbox("Formula:", [
                "A + B", "A - B", "A × B", "A / B",
                "log(A)", "sqrt(A)", "A²",
                "A - mean(A)", "A - median(A)",
            ])
            col_a = st.selectbox("Column A:", nc, key="ca")
            col_b = None
            if "B" in formula:
                col_b = st.selectbox("Column B:", nc, key="cb")

            if st.button("✓ Create"):
                if not new_col.strip():
                    st.warning("Enter a column name."); st.stop()
                before = df.copy(); save_undo(); new_df = df.copy()
                try:
                    a = pd.to_numeric(new_df[col_a], errors="coerce")
                    b = pd.to_numeric(new_df[col_b], errors="coerce") if col_b else None
                    ops = {
                        "A + B": lambda: a + b,
                        "A - B": lambda: a - b,
                        "A × B": lambda: a * b,
                        "A / B": lambda: a / b.replace(0, np.nan),
                        "log(A)": lambda: np.log(a.where(a > 0)),
                        "sqrt(A)": lambda: np.sqrt(a.where(a >= 0)),
                        "A²": lambda: a ** 2,
                        "A - mean(A)": lambda: a - a.mean(),
                        "A - median(A)": lambda: a - a.median(),
                    }
                    new_df[new_col.strip()] = ops[formula]()
                    st.session_state["working_df"] = new_df
                    log_step("computed_column", {"formula": formula, "col_a": col_a, "col_b": str(col_b)}, [new_col.strip()])
                    commit(before, new_df, f"Created column '{new_col.strip()}' = {formula}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

        elif action == "Bin numeric column into categories":
            nc = num_cols(df)
            if not nc:
                card("No numeric columns.", "warn"); st.stop()
            col = st.selectbox("Column to bin:", nc)
            bins = st.slider("Number of bins:", 2, 20, 4)
            bin_method = st.radio("Binning method:", ["Equal Width", "Equal Frequency (Quantiles)"], horizontal=True)
            new_col = st.text_input("New column name:", value=f"{col}_bin")
            if st.button("✓ Bin"):
                before = df.copy(); save_undo(); new_df = df.copy()
                try:
                    if bin_method == "Equal Width":
                        new_df[new_col] = pd.cut(new_df[col], bins=bins)
                    else:
                        new_df[new_col] = pd.qcut(new_df[col], q=bins, duplicates="drop")
                    st.session_state["working_df"] = new_df
                    log_step("bin_column", {"bins": bins, "method": bin_method}, [col, new_col])
                    commit(before, new_df, f"Binned '{col}' into {bins} categories → '{new_col}'.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Binning failed: {e}")

    # ──────────────────────────────────────────────
    # ⑧ DATA VALIDATION
    # ──────────────────────────────────────────────
    elif tool == "⑧ Data Validation":
        card("Define a rule to find rows that violate it. Violations can be exported or removed.", "info")

        rule = st.selectbox("Rule type:", [
            "Numeric range check",
            "Allowed category values",
            "Required non-null columns",
        ])

        violations = pd.DataFrame()
        checked = False

        if rule == "Numeric range check":
            nc = num_cols(df)
            if not nc:
                card("No numeric columns.", "warn"); st.stop()
            col = st.selectbox("Column:", nc)
            sv = pd.to_numeric(df[col], errors="coerce")
            c1, c2 = st.columns(2)
            mn = c1.number_input("Minimum allowed:", value=float(sv.min()) if sv.notna().any() else 0.0)
            mx = c2.number_input("Maximum allowed:", value=float(sv.max()) if sv.notna().any() else 100.0)
            if st.button("✓ Find Violations"):
                violations = df[(sv < mn) | (sv > mx) | sv.isna()]
                checked = True

        elif rule == "Allowed category values":
            all_cols = cat_cols(df) or df.columns.tolist()
            col = st.selectbox("Column:", all_cols)
            all_vals = sorted(df[col].dropna().astype(str).unique().tolist())
            allowed = st.multiselect("Allowed values:", all_vals, default=all_vals[:min(10, len(all_vals))])
            if st.button("✓ Find Violations"):
                violations = df[~df[col].astype(str).isin(allowed)]
                checked = True

        elif rule == "Required non-null columns":
            req_cols = st.multiselect("Columns that must not be empty:", df.columns.tolist())
            if st.button("✓ Find Violations"):
                if req_cols:
                    violations = df[df[req_cols].isna().any(axis=1)]
                checked = True

        if checked:
            if not violations.empty:
                card(f"Found <b>{len(violations):,} violating rows</b> (showing first 200).", "warn")
                st.dataframe(violations.head(200), use_container_width=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("⬇ Export Violations CSV", violations.to_csv(index=False).encode(), "violations.csv", "text/csv", use_container_width=True)
                with c2:
                    if st.button("✗ Remove Violating Rows", use_container_width=True):
                        before = df.copy(); save_undo()
                        new_df = df.drop(index=violations.index).reset_index(drop=True)
                        st.session_state["working_df"] = new_df
                        log_step("remove_violations", {"rule": rule}, [])
                        commit(before, new_df, f"Removed {len(violations):,} violating rows.")
                        st.rerun()
            else:
                card("No violations found — data passes this rule!", "success")


# =========================================================
# PAGE C — VISUALIZATION BUILDER
# =========================================================
elif page == "📊 Visualization Builder":
    st.markdown('<div class="app-title" style="font-size:1.6rem;">Visualization Builder</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-sub">Build charts from your cleaned data</div>', unsafe_allow_html=True)

    df = st.session_state["working_df"]
    if df is None:
        st.warning("Upload a dataset first."); st.stop()

    nc = num_cols(df)
    cc = cat_cols(df)
    dc = dt_cols(df)

    CHART_TYPES = {
        "📊 Histogram": "Distribution of one numeric column.",
        "📦 Box Plot": "Spread and outliers, optionally grouped by category.",
        "🔵 Scatter Plot": "Relationship between two numeric columns.",
        "📈 Line Chart": "Trend over time or an ordered sequence.",
        "📊 Bar Chart": "Compare totals or averages across categories.",
        "🌡 Correlation Heatmap": "Correlation matrix of numeric columns.",
    }

    chart = st.selectbox("Chart type:", list(CHART_TYPES.keys()))
    card(CHART_TYPES[chart], "info")

    with st.expander("🔽 Filter rows (optional)", expanded=False):
        fm = st.selectbox("Filter by:", ["No filter", "Category values", "Numeric range"])
        fdf = df.copy()
        if fm == "Category values" and cc:
            fc = st.selectbox("Column:", cc, key="fcat")
            vals = sorted(fdf[fc].dropna().astype(str).unique().tolist())
            sel = st.multiselect("Keep values:", vals, default=vals[:min(10, len(vals))])
            if sel:
                fdf = fdf[fdf[fc].astype(str).isin(sel)]
        elif fm == "Numeric range" and nc:
            fc = st.selectbox("Column:", nc, key="fnum")
            sv = pd.to_numeric(fdf[fc], errors="coerce").dropna()
            if not sv.empty:
                lo, hi = float(sv.min()), float(sv.max())
                rng = st.slider("Range:", lo, hi, (lo, hi))
                fdf = fdf[pd.to_numeric(fdf[fc], errors="coerce").between(rng[0], rng[1])]
        st.caption(f"Rows after filtering: **{len(fdf):,}** of {len(df):,}")

    if fdf.empty:
        st.warning("No data after filtering."); st.stop()

    st.markdown("---")
    fig = None

    if chart == "📊 Histogram":
        if not nc:
            card("Need at least one numeric column.", "warn"); st.stop()
        col_x = st.selectbox("Column:", nc)
        c1, c2 = st.columns(2)
        bins = c1.slider("Number of bins:", 5, 100, 30)
        col_c = c2.selectbox("Color by (optional):", ["None"] + cc)
        plot_df = fdf.copy()
        plot_df[col_x] = pd.to_numeric(plot_df[col_x], errors="coerce")
        fig = px.histogram(plot_df, x=col_x, nbins=bins,
                           color=None if col_c == "None" else col_c,
                           marginal="box", opacity=0.8,
                           title=f"Distribution of {col_x}")
        fig = theme_fig(fig)

    elif chart == "📦 Box Plot":
        if not nc:
            card("Need at least one numeric column.", "warn"); st.stop()
        col_y = st.selectbox("Numeric column:", nc)
        col_x = st.selectbox("Group by (optional):", ["None"] + cc)
        x = None if col_x == "None" else col_x

        fig_m, ax = plt.subplots(figsize=(10, 4))
        if x is None:
            data_to_plot = [pd.to_numeric(fdf[col_y], errors="coerce").dropna().tolist()]
            labels = [col_y]
        else:
            groups = sorted(fdf[x].dropna().unique())[:15]
            data_to_plot = [pd.to_numeric(fdf[fdf[x] == g][col_y], errors="coerce").dropna().tolist() for g in groups]
            labels = [str(g) for g in groups]
        ax.boxplot(data_to_plot, labels=labels, patch_artist=True,
                   boxprops=dict(facecolor="#6366f120", color="#6366f1"),
                   medianprops=dict(color="#06b6d4", linewidth=2),
                   whiskerprops=dict(color="#6366f1"),
                   capprops=dict(color="#6366f1"),
                   flierprops=dict(marker="o", color="#ef4444", alpha=0.5, markersize=4))
        ax.set_title(f"Box Plot — {col_y}" + (f" by {x}" if x else ""), fontweight="bold")
        ax.set_ylabel(col_y)
        if x:
            ax.set_xlabel(x)
            plt.xticks(rotation=30, ha="right")
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        st.pyplot(fig_m, use_container_width=True)
        plt.close(fig_m)

        fig = px.box(fdf, x=x, y=col_y, title=f"Box Plot — {col_y}" + (f" by {x}" if x else ""),
                     color=x, points="outliers")
        fig = theme_fig(fig)

    elif chart == "🔵 Scatter Plot":
        if len(nc) < 2:
            card("Need at least 2 numeric columns.", "warn"); st.stop()
        c1, c2 = st.columns(2)
        col_x = c1.selectbox("X axis:", nc, key="sx")
        col_y = c2.selectbox("Y axis:", [c for c in nc if c != col_x] or nc, key="sy")
        col_c = st.selectbox("Color by (optional):", ["None"] + cc + nc)
        trend = st.checkbox("Add trendline (OLS)")
        plot_df = fdf.copy()
        for c in [col_x, col_y]:
            plot_df[c] = pd.to_numeric(plot_df[c], errors="coerce")
        plot_df = plot_df.dropna(subset=[col_x, col_y])
        fig = px.scatter(plot_df, x=col_x, y=col_y,
                         color=None if col_c == "None" else col_c,
                         trendline="ols" if trend else None,
                         opacity=0.7, title=f"{col_y} vs {col_x}")
        fig = theme_fig(fig)

    elif chart == "📈 Line Chart":
        if not nc:
            card("Need at least one numeric column.", "warn"); st.stop()
        col_x = st.selectbox("X axis (time/sequence):", dc + df.columns.tolist())
        col_y = st.multiselect("Y column(s):", nc, default=[nc[0]])
        c1, c2 = st.columns(2)
        agg = c1.selectbox("Aggregate by X:", ["None", "Sum", "Mean", "Count"])
        smooth = c2.slider("Rolling average window (1 = off):", 1, 50, 1)
        if not col_y:
            card("Select at least one Y column.", "warn")
        else:
            plot_df = fdf[[col_x] + col_y].copy()
            plot_df[col_x] = pd.to_datetime(plot_df[col_x], errors="coerce")
            for c in col_y:
                plot_df[c] = pd.to_numeric(plot_df[c], errors="coerce")
            plot_df = plot_df.dropna(subset=[col_x]).sort_values(col_x)
            if agg != "None":
                am = {"Sum": "sum", "Mean": "mean", "Count": "count"}
                plot_df = plot_df.groupby(col_x)[col_y].agg(am[agg]).reset_index()
            if smooth > 1:
                for c in col_y:
                    plot_df[c] = plot_df[c].rolling(smooth, center=True).mean()
            fig = px.line(plot_df, x=col_x, y=col_y, title=f"Trend: {', '.join(col_y)}",
                          markers=len(plot_df) < 100)
            fig = theme_fig(fig)

    elif chart == "📊 Bar Chart":
        if not cc:
            card("Need at least one category column.", "warn"); st.stop()
        col_x = st.selectbox("Category column:", cc)
        c1, c2 = st.columns(2)
        agg = c1.selectbox("Aggregation:", ["Count", "Sum", "Mean", "Median"])
        top_n = c2.slider("Top N categories:", 3, 50, 15)
        orient = st.radio("Orientation:", ["Vertical", "Horizontal"], horizontal=True)

        if agg == "Count":
            grouped = fdf[col_x].astype(str).value_counts().head(top_n).reset_index()
            grouped.columns = [col_x, "Count"]
            y_col = "Count"
        else:
            if not nc:
                card("Need a numeric column for this aggregation.", "warn"); st.stop()
            col_y = st.selectbox("Numeric column:", nc)
            am = {"Sum": "sum", "Mean": "mean", "Median": "median"}
            tmp = fdf[[col_x, col_y]].copy()
            tmp[col_y] = pd.to_numeric(tmp[col_y], errors="coerce")
            grouped = tmp.groupby(col_x)[col_y].agg(am[agg]).sort_values(ascending=False).head(top_n).reset_index()
            y_col = col_y

        if orient == "Horizontal":
            fig = px.bar(grouped, x=y_col, y=col_x, orientation="h",
                         title=f"{agg} by {col_x} (top {top_n})", text_auto=True)
        else:
            fig = px.bar(grouped, x=col_x, y=y_col,
                         title=f"{agg} by {col_x} (top {top_n})", text_auto=True)
        fig = theme_fig(fig)

    elif chart == "🌡 Correlation Heatmap":
        if len(nc) < 2:
            card("Need at least 2 numeric columns.", "warn"); st.stop()
        sel = st.multiselect("Columns to include:", nc, default=nc[:min(10, len(nc))])
        method = st.selectbox("Method:", ["pearson", "spearman"])
        if len(sel) < 2:
            card("Select at least 2 columns.", "warn")
        else:
            corr = fdf[sel].apply(pd.to_numeric, errors="coerce").corr(method=method)
            fig = px.imshow(corr, color_continuous_scale="RdBu_r", zmin=-1, zmax=1,
                            text_auto=".2f", aspect="auto",
                            title=f"Correlation Matrix ({method.title()})")
            fig = theme_fig(fig)

            fig_m, ax = plt.subplots(figsize=(max(6, len(sel)), max(5, len(sel) - 1)))
            im = ax.imshow(corr.values, cmap="RdBu_r", vmin=-1, vmax=1, aspect="auto")
            plt.colorbar(im, ax=ax, shrink=0.8)
            ax.set_xticks(range(len(sel))); ax.set_xticklabels(sel, rotation=45, ha="right", fontsize=9)
            ax.set_yticks(range(len(sel))); ax.set_yticklabels(sel, fontsize=9)
            for i in range(len(sel)):
                for j in range(len(sel)):
                    ax.text(j, i, f"{corr.values[i, j]:.2f}", ha="center", va="center", fontsize=8,
                            color="white" if abs(corr.values[i, j]) > 0.5 else "black")
            ax.set_title(f"Correlation Heatmap ({method.title()})", fontweight="bold")
            plt.tight_layout()
            st.pyplot(fig_m, use_container_width=True)
            plt.close(fig_m)

    if fig is not None:
        st.plotly_chart(fig, use_container_width=True, theme="streamlit")


# =========================================================
# PAGE D — EXPORT & REPORT
# =========================================================
elif page == "📤 Export & Report":
    st.markdown('<div class="app-title" style="font-size:1.6rem;">Export & Report</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-sub">Download your cleaned dataset and transformation log</div>', unsafe_allow_html=True)

    df = st.session_state["working_df"]
    orig = st.session_state["original_df"]
    if df is None:
        st.warning("Upload a dataset first."); st.stop()

    if orig is not None:
        sec("Before vs After Summary")
        impact_strip({
            "rb": len(orig), "ra": len(df),
            "cb": orig.shape[1], "ca": df.shape[1],
            "mb": int(orig.isna().sum().sum()), "ma": int(df.isna().sum().sum()),
        })
        rd = len(df) - len(orig)
        md = int(df.isna().sum().sum()) - int(orig.isna().sum().sum())
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Original Rows",    f"{len(orig):,}")
        c2.metric("Cleaned Rows",     f"{len(df):,}",  delta=f"{rd:+,}", delta_color="off")
        c3.metric("Original Missing", f"{int(orig.isna().sum().sum()):,}")
        c4.metric("Cleaned Missing",  f"{int(df.isna().sum().sum()):,}", delta=f"{md:+,}", delta_color="inverse")

    sec("Transformation Log")
    log = st.session_state["log"]
    if log:
        log_df = pd.DataFrame([{
            "Step":      i + 1,
            "Timestamp": e["timestamp"],
            "Action":    e["step"],
            "Params":    json.dumps(e.get("params", {})),
            "Columns":   ", ".join(e.get("affected_columns", [])) or "—",
        } for i, e in enumerate(log)])
        st.dataframe(log_df, use_container_width=True, height=300)
    else:
        card("No transformations logged yet — apply some cleaning steps first.", "info")

    sec("Download Files")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.download_button("⬇ Cleaned CSV", to_csv_bytes(df), "cleaned_dataset.csv", "text/csv", use_container_width=True)
    with c2:
        if EXCEL_OK:
            st.download_button("⬇ Cleaned Excel", to_excel_bytes(df), "cleaned_dataset.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.caption("Install `openpyxl` for Excel export.")
    with c3:
        recipe = {
            "generated_at": datetime.now().isoformat(),
            "source_file":  st.session_state.get("filename", "unknown"),
            "steps":        log,
        }
        st.download_button("⬇ JSON Recipe", json.dumps(recipe, indent=2).encode(),
                           "transformation_recipe.json", "application/json", use_container_width=True)
    with c4:
        if orig is not None:
            st.download_button("⬇ Original CSV", to_csv_bytes(orig), "original_dataset.csv", "text/csv", use_container_width=True)

    sec("Final Dataset Profile")
    st.dataframe(profile_df(df), use_container_width=True)

    with st.expander("View JSON Recipe (raw)"):
        st.code(json.dumps(recipe if log else {"steps": []}, indent=2), language="json")
