import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher
from openpyxl import Workbook

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BT Hours Tracker",
    page_icon="🏝️",
    layout="wide",
)

HOUR_GOAL = 1500

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background-color: #0f1117; color: #e8eaf0; }
h1, h2, h3 { font-family: 'DM Mono', monospace !important; letter-spacing: -0.5px; }

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 2.2rem 3rem 4rem 3rem; max-width: 1200px; }

.main-title {
    font-family: 'DM Mono', monospace;
    font-size: 2rem; color: #7ee8a2;
    border-bottom: 2px solid #7ee8a2;
    padding-bottom: 8px; margin-bottom: 4px;
}
.subtitle {
    color: #888; font-size: 0.9rem;
    margin-bottom: 32px;
    font-family: 'DM Mono', monospace;
}

.section-header {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem; color: #555;
    text-transform: uppercase; letter-spacing: 3px;
    margin: 28px 0 12px 0;
}

/* Upload zones */
.upload-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.75rem; color: #7ee8a2;
    text-transform: uppercase; letter-spacing: 2px;
    margin-bottom: 6px;
}
[data-testid="stFileUploader"] {
    background: #1a1d27;
    border: 1px dashed #2a2d3a;
    border-radius: 10px;
    padding: 0.4rem 0.8rem;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #7ee8a2; }
[data-testid="stFileUploader"] label {
    color: #aaa !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.8rem !important;
}

/* Stat cards */
.stat-card {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 10px;
    padding: 18px 20px;
}
.stat-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.62rem; color: #555;
    text-transform: uppercase; letter-spacing: 2px;
    margin-bottom: 6px;
}
.stat-value {
    font-family: 'DM Mono', monospace;
    font-size: 1.8rem; font-weight: 700;
    line-height: 1;
}
.stat-value.green  { color: #7ee8a2; }
.stat-value.blue   { color: #74b9ff; }
.stat-value.white  { color: #e8eaf0; }
.stat-value.amber  { color: #ffd93d; }

/* Metric cards */
.metric-card {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 10px;
    padding: 18px 20px;
    margin: 6px 0;
}
.metric-card.on-track { border-left: 4px solid #7ee8a2; }
.metric-card.warning  { border-left: 4px solid #ffd93d; }
.metric-card.behind   { border-left: 4px solid #74b9ff; }

/* Badges */
.badge {
    padding: 2px 10px; border-radius: 20px;
    font-size: 0.72rem; font-family: 'DM Mono', monospace;
    font-weight: 600; display: inline-block;
}
.badge-green  { background: #7ee8a222; color: #7ee8a2; border: 1px solid #7ee8a2; }
.badge-amber  { background: #ffd93d22; color: #ffd93d; border: 1px solid #ffd93d; }
.badge-blue   { background: #74b9ff22; color: #74b9ff; border: 1px solid #74b9ff; }
.badge-red    { background: #ff6b6b22; color: #ff6b6b; border: 1px solid #ff6b6b; }

/* Download button */
[data-testid="stDownloadButton"] button {
    background: #1a1d27 !important;
    color: #7ee8a2 !important;
    border: 1px solid #7ee8a2 !important;
    border-radius: 8px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.85rem !important;
    padding: 0.55rem 1.6rem !important;
    transition: all 0.2s !important;
    width: 100% !important;
}
[data-testid="stDownloadButton"] button:hover {
    background: #7ee8a222 !important;
    box-shadow: 0 0 12px #7ee8a244 !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid #2a2d3a !important;
    border-radius: 8px !important;
    overflow: hidden;
}

/* Expander */
[data-testid="stExpander"] {
    background: #1a1d27 !important;
    border: 1px solid #2a2d3a !important;
    border-radius: 8px !important;
}
.streamlit-expanderHeader {
    font-family: 'DM Mono', monospace !important;
    font-size: 0.82rem !important;
    color: #888 !important;
}

/* Alerts */
[data-testid="stAlert"] {
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
}

hr { border-color: #2a2d3a; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: #0d0f16 !important;
    border-right: 1px solid #1e2130 !important;
}
[data-testid="stSidebar"] * { color: #888 !important; }
[data-testid="stSidebar"] code { background: #1a1d27 !important; color: #7ee8a2 !important; }

/* Empty state box */
.empty-state {
    background: #1a1d27;
    border: 1px dashed #2a2d3a;
    border-radius: 10px;
    padding: 48px;
    text-align: center;
    margin-top: 24px;
}
.empty-state-title {
    font-family: 'DM Mono', monospace;
    font-size: 1rem; color: #444; margin-bottom: 8px;
}
.empty-state-sub {
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem; color: #2a2d3a;
}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">BT Hours Tracker</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="subtitle">AlohaABA Export  →  Direct Service BT Hours Report'
    f'&nbsp;&nbsp;·&nbsp;&nbsp;Goal: {HOUR_GOAL:,} hrs / BT</div>',
    unsafe_allow_html=True
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Expected Columns")
    st.markdown("**Billing Export** *(required)*")
    st.code("Service Name\nCompleted\nUnits\nStaff Name", language="text")
    st.markdown("**BT Contacts** *(optional)*")
    st.code("BT Name  (First Last)\nPhone\nEmail", language="text")
    st.markdown("---")
    st.markdown("**Logic**")
    st.markdown(
        f"- Filter: `Service Name = Direct Service BT`\n"
        f"- Filter: `Completed = Yes`\n"
        f"- Group by: `Staff Name`\n"
        f"- Hours = Units ÷ 4\n"
        f"- Goal: **{HOUR_GOAL:,} hrs** per BT\n"
        f"- 1 unit = 15 min · 4 units = 1 hr"
    )

# ── Helpers ───────────────────────────────────────────────────────────────────
def read_any(f):
    return pd.read_csv(f) if f.name.endswith(".csv") else pd.read_excel(f)

def normalize_cols(df):
    df.columns = df.columns.str.strip()
    return df

def norm_name(s):
    s = str(s).strip().lower().replace(",", " ")
    return " ".join(s.split())

def to_last_first(name):
    parts = str(name).strip().split()
    if len(parts) >= 2:
        return f"{parts[-1]}, {' '.join(parts[:-1])}"
    return str(name).strip()

# ── Processing ────────────────────────────────────────────────────────────────
def process_billing(df, bt_contacts_df=None):
    df = normalize_cols(df)
    col_map = {c.lower(): c for c in df.columns}

    service_col   = next((col_map[k] for k in col_map if "service name" in k), None)
    completed_col = next((col_map[k] for k in col_map if "completed" in k), None)
    units_col     = next((col_map[k] for k in col_map if "units" in k), None)
    bt_col        = next((col_map[k] for k in col_map if "staff name" in k or k in
                          ["bt", "behavior technician", "staff", "provider", "technician", "employee"]), None)

    missing = []
    if not service_col:   missing.append("Service Name")
    if not completed_col: missing.append("Completed")
    if not units_col:     missing.append("Units")
    if missing:
        st.error(f"Missing required columns: **{', '.join(missing)}**  \nDetected: {list(df.columns)}")
        return None

    df[service_col] = df[service_col].astype(str).str.strip()
    filtered = df[df[service_col].str.lower() == "direct service bt"].copy()
    if filtered.empty:
        st.warning("⚠️ No rows matched `Service Name = Direct Service BT`")
        return None

    filtered[completed_col] = filtered[completed_col].astype(str).str.strip()
    filtered = filtered[filtered[completed_col].str.lower() == "yes"].copy()
    if filtered.empty:
        st.warning("⚠️ No completed Direct Service BT rows found.")
        return None

    appt_status_col = next((c for c in filtered.columns if "appointment" in c.lower() and "status" in c.lower()), None)
    if appt_status_col:
        filtered = filtered[filtered[appt_status_col].astype(str).str.strip().str.lower() == "active"].copy()
        filtered[units_col] = pd.to_numeric(filtered[units_col], errors="coerce").fillna(0)

    if not bt_col:
        candidates = [col_map[k] for k in col_map if any(x in k for x in ["name", "client", "patient"])]
        bt_col = candidates[0] if candidates else None
    if not bt_col:
        st.error("❌ Could not find a Staff Name column.")
        return None

    filtered = filtered.rename(columns={bt_col: "Staff Name"})
    filtered["Phone"] = ""
    filtered["Email"] = ""

    match_log = []

    if bt_contacts_df is not None:
        bt_df = normalize_cols(bt_contacts_df)
        bt_required = {"BT Name", "Phone", "Email"}
        bt_missing = bt_required - set(bt_df.columns)
        if bt_missing:
            st.error(f"BT Contacts missing columns: {sorted(bt_missing)}")
        else:
            bt_df["BT_formatted"] = bt_df["BT Name"].apply(to_last_first)
            bt_df["bt_norm"]      = bt_df["BT_formatted"].apply(norm_name)
            staff_to_phone, staff_to_email = {}, {}

            for staff in filtered["Staff Name"].dropna().unique():
                staff_norm = norm_name(staff)
                best_score, best_row = 0.0, None
                for _, bt_row in bt_df.iterrows():
                    score = SequenceMatcher(None, staff_norm, bt_row["bt_norm"]).ratio()
                    if score > best_score:
                        best_score = score
                        best_row = bt_row

                if best_row is not None and best_score >= 0.8:
                    staff_to_phone[staff] = best_row["Phone"]
                    staff_to_email[staff] = best_row["Email"]
                    match_log.append({"Staff Name": staff, "Matched To": best_row["BT Name"], "Score": round(best_score, 2), "Status": "✅ Matched"})
                else:
                    match_log.append({"Staff Name": staff, "Matched To": best_row["BT Name"] if best_row is not None else "—", "Score": round(best_score, 2) if best_row is not None else 0, "Status": "⚠️ No match"})

            filtered["Phone"] = filtered["Staff Name"].map(staff_to_phone).fillna("")
            filtered["Email"] = filtered["Staff Name"].map(staff_to_email).fillna("")

    summary = filtered.groupby("Staff Name").agg(
        Total_Units=(units_col, "sum"),
        Phone=("Phone", "first"),
        Email=("Email", "first"),
    ).reset_index()
    summary.columns = ["BT Name", "Total Units", "Phone", "Email"]
    summary["Total Hours"]             = summary["Total Units"] / 4
    summary["Hours Remaining to 1500"] = HOUR_GOAL - summary["Total Hours"]
    summary = summary[["BT Name", "Phone", "Email", "Total Units", "Total Hours", "Hours Remaining to 1500"]]
    summary = summary[summary["Total Hours"] > 0]
    summary = summary.sort_values("BT Name").reset_index(drop=True)

    return summary, match_log

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(summary):
    has_contacts = summary["Phone"].astype(str).str.strip().ne("").any()

    wb = Workbook()
    ws = wb.active
    ws.title = "BT Hours Summary"

    if has_contacts:
        headers = ["BT Name", "Phone", "Email", "Total Units", "Total Hours", "Hours Remaining to 1500"]
    else:
        headers = ["BT Name", "Total Units", "Total Hours", "Hours Remaining to 1500"]

    ws.append(headers)

    for _, row in summary.iterrows():
        if has_contacts:
            ws.append([
                row["BT Name"], row["Phone"], row["Email"],
                row["Total Units"], row["Total Hours"], row["Hours Remaining to 1500"]
            ])
        else:
            ws.append([
                row["BT Name"], row["Total Units"],
                row["Total Hours"], row["Hours Remaining to 1500"]
            ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Upload section ────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">Upload Data</div>', unsafe_allow_html=True)
col_a, col_b = st.columns(2)

with col_a:
    st.markdown('<div class="upload-label">Billing Export</div>', unsafe_allow_html=True)
    billing_file = st.file_uploader(
        "", type=["xlsx", "xls", "csv"], key="billing",
        label_visibility="collapsed"
    )

with col_b:
    st.markdown('<div class="upload-label">BT Contacts &nbsp;<span style="color:#555;font-size:0.65rem">optional</span></div>', unsafe_allow_html=True)
    bt_contacts_file = st.file_uploader(
        "", type=["xlsx", "xls", "csv"], key="contacts",
        label_visibility="collapsed",
        help="Needs: BT Name (First Last), Phone, Email"
    )

# ── Main logic ────────────────────────────────────────────────────────────────
if billing_file:
    try:
        df             = read_any(billing_file)
        bt_contacts_df = read_any(bt_contacts_file) if bt_contacts_file else None

        result = process_billing(df, bt_contacts_df)

        if result is not None:
            summary, match_log = result

            # ── Contact match log ─────────────────────────────────────────────
            if bt_contacts_file and match_log:
                matched = sum(1 for m in match_log if "✅" in m["Status"])
                with st.expander(f"🔗 Contact matching — {matched}/{len(match_log)} matched"):
                    st.dataframe(pd.DataFrame(match_log), use_container_width=True, hide_index=True)

            # ── Download ──────────────────────────────────────────────────────
            st.markdown('<div class="section-header">Export</div>', unsafe_allow_html=True)
            dl_col, _ = st.columns([1, 2])
            with dl_col:
                excel_buf = build_excel(summary)
                st.download_button(
                    label="⬇  Download Excel Report",
                    data=excel_buf,
                    file_name="aloha_bt_hours_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)

elif not billing_file:
    st.markdown("""
    <div class="empty-state">
        <div class="empty-state-title">Upload your AlohaABA billing export above to get started</div>
        <div class="empty-state-sub">Accepts .csv · .xlsx · .xls</div>
    </div>
    """, unsafe_allow_html=True)