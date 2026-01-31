import re
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
from datetime import datetime

st.set_page_config(page_title="Student Admission & Exam Status Dashboard", layout="wide")

# -----------------------------
# Config
# -----------------------------
DEFAULT_FILE = "Student Profile.xlsx"
DEFAULT_SHEET = "Exam Status"

# Your business rule reference year (you earlier fixed to 2025)
REFERENCE_YEAR = 2025

# -----------------------------
# Helpers
# -----------------------------
def safe_num(x):
    try:
        if pd.isna(x):
            return 0.0
        return float(pd.to_numeric(x, errors="coerce") or 0)
    except:
        return 0.0

def parse_admission_year(batch):
    # Expected: "2020-JUL" or similar
    if pd.isna(batch):
        return np.nan
    m = re.match(r"^\s*(\d{4})", str(batch).strip())
    return int(m.group(1)) if m else np.nan

def course_duration(programme: str) -> int:
    # MBA / MCA => 2 year course, others default 3 year
    p = (programme or "").upper()
    if "MBA" in p or "MCA" in p:
        return 2
    return 3

def load_exam_status_df(uploaded_file):
    # Read pivot sheet where header row is at Excel row 4 (0-index header=3)
    df = pd.read_excel(uploaded_file, sheet_name=DEFAULT_SHEET, engine="openpyxl", header=3)

    # Standardize column names (strip)
    df.columns = [str(c).strip() for c in df.columns]

    # Drop fully empty rows
    df = df.dropna(how="all").copy()

    # Forward fill Programme, Batch (pivot style)
    if "Programme" in df.columns:
        df["Programme"] = df["Programme"].ffill()
    if "Batch" in df.columns:
        df["Batch"] = df["Batch"].ffill()

    # AdmissionYear
    df["AdmissionYear"] = df["Batch"].apply(parse_admission_year)

    # Clean agency
    if "Marketing Agency" not in df.columns:
        df["Marketing Agency"] = "UNKNOWN"
    df["Marketing Agency"] = df["Marketing Agency"].fillna("UNKNOWN").astype(str).str.strip()

    # Convert numeric cols
    for c in df.columns:
        if c not in ["Programme", "Batch", "Marketing Agency"]:
            df[c] = df[c].apply(safe_num)

    # Identify key columns
    col_cancel = "01. ADMISSION CANCELLED"
    col_passout = "02. PASSOUT"
    col_passout_conv = "02. PASSOUT & CONVOCATION"
    col_validity = "13. VALIDITY EXPIRED"
    col_total = "Grand Total"

    # Continuation / pending buckets (all except cancelled/passout/validity/total)
    continuation_cols = [c for c in df.columns if any([
        c.startswith("03."),
        c.startswith("04."),
        c.startswith("05."),
        c.startswith("06."),
        c.startswith("07."),
        c.startswith("08."),
        c.startswith("09."),
        c.startswith("10."),
        c.startswith("11."),
    ])]

    # Build metrics
    df["Cancelled"] = df[col_cancel] if col_cancel in df.columns else 0
    df["Passout_Total"] = (df[col_passout] if col_passout in df.columns else 0) + (df[col_passout_conv] if col_passout_conv in df.columns else 0)
    df["Validity_Expired"] = df[col_validity] if col_validity in df.columns else 0
    df["Continuation_Total"] = df[continuation_cols].sum(axis=1) if continuation_cols else 0
    df["GrandTotal"] = df[col_total] if col_total in df.columns else df[["Cancelled","Passout_Total","Validity_Expired","Continuation_Total"]].sum(axis=1)

    # Iteration rate = Validity Expired / Total
    df["Iteration_Rate_%"] = np.where(df["GrandTotal"] > 0, (df["Validity_Expired"] / df["GrandTotal"]) * 100, 0)

    # Eligibility rule (course + 2 years)
    df["CourseDuration"] = df["Programme"].astype(str).apply(course_duration)
    df["MaxAllowedYears"] = df["CourseDuration"] + 2
    df["YearsSinceAdmission"] = REFERENCE_YEAR - df["AdmissionYear"]
    df["Eligible_By_Rule"] = np.where(df["YearsSinceAdmission"] <= df["MaxAllowedYears"], 1, 0)

    # Placement eligible:
    # - For 3-year (UG): 5th sem onwards -> 10/11 + FINAL SEM + LAST SEM BACKLOG
    # - For 2-year (MBA/MCA): 3rd sem onwards -> 08/09 + FINAL SEM + LAST SEM BACKLOG
    def placement_cols_for_program(duration):
        cols = []
        if duration == 3:
            cols = [c for c in df.columns if c.startswith("10.") or c.startswith("11.") or c.startswith("06.") or c.startswith("07.")]
        else:
            cols = [c for c in df.columns if c.startswith("08.") or c.startswith("09.") or c.startswith("06.") or c.startswith("07.")]
        return cols

    placement_vals = []
    for i, row in df.iterrows():
        dur = int(row["CourseDuration"])
        cols = placement_cols_for_program(dur)
        placement_vals.append(float(row[cols].sum()) if cols else 0.0)

    df["Placement_Eligible"] = placement_vals
    df["Placement_Eligible_%"] = np.where(df["GrandTotal"] > 0, (df["Placement_Eligible"] / df["GrandTotal"]) * 100, 0)

    return df

def group_block(df, group_cols):
    g = df.groupby(group_cols, dropna=False).agg(
        Cancelled=("Cancelled","sum"),
        Passout_Total=("Passout_Total","sum"),
        Continuation_Total=("Continuation_Total","sum"),
        Validity_Expired=("Validity_Expired","sum"),
        Placement_Eligible=("Placement_Eligible","sum"),
        GrandTotal=("GrandTotal","sum"),
    ).reset_index()

    g["Cancelled_%"] = np.where(g["GrandTotal"] > 0, g["Cancelled"]/g["GrandTotal"]*100, 0)
    g["Passout_%"] = np.where(g["GrandTotal"] > 0, g["Passout_Total"]/g["GrandTotal"]*100, 0)
    g["Continuation_%"] = np.where(g["GrandTotal"] > 0, g["Continuation_Total"]/g["GrandTotal"]*100, 0)
    g["Validity_Expired_%"] = np.where(g["GrandTotal"] > 0, g["Validity_Expired"]/g["GrandTotal"]*100, 0)
    g["Placement_Eligible_%"] = np.where(g["GrandTotal"] > 0, g["Placement_Eligible"]/g["GrandTotal"]*100, 0)
    g["Iteration_Rate_%"] = g["Validity_Expired_%"]  # same definition

    return g

def bar_chart(df, x_col, y_col, title):
    c = (
        alt.Chart(df)
        .mark_bar()
        .encode(
            x=alt.X(f"{x_col}:N", sort="-y"),
            y=alt.Y(f"{y_col}:Q"),
            tooltip=[x_col, y_col]
        )
        .properties(height=320, title=title)
    )
    st.altair_chart(c, use_container_width=True)

# -----------------------------
# UI - Upload
# -----------------------------
st.title("📊 Student Admission & Exam Status Dashboard")

uploaded = st.sidebar.file_uploader("Upload file (.xlsx or .csv)", type=["xlsx","xls","csv"])

if not uploaded:
    st.info("Upload the Excel to start.")
    st.stop()

try:
    if str(uploaded.name).lower().endswith(".csv"):
        st.error("Please upload the Excel (.xlsx). Your current format is designed for the 'Exam Status' sheet.")
        st.stop()

    df = load_exam_status_df(uploaded)
except Exception as e:
    st.error(f"Failed to parse file: {e}")
    st.stop()

# -----------------------------
# Filters
# -----------------------------
st.sidebar.subheader("Filters")

programs = sorted(df["Programme"].dropna().unique().tolist())
years = sorted([int(x) for x in df["AdmissionYear"].dropna().unique().tolist()])
agencies = sorted(df["Marketing Agency"].dropna().unique().tolist())

sel_program = st.sidebar.multiselect("Program", programs, default=programs[:5] if len(programs) > 5 else programs)
sel_year = st.sidebar.multiselect("Admission Year", years, default=years)
sel_agency = st.sidebar.multiselect("Agency", agencies, default=agencies[:5] if len(agencies) > 5 else agencies)

f = df.copy()
if sel_program:
    f = f[f["Programme"].isin(sel_program)]
if sel_year:
    f = f[f["AdmissionYear"].isin(sel_year)]
if sel_agency:
    f = f[f["Marketing Agency"].isin(sel_agency)]

# -----------------------------
# Tabs
# -----------------------------
tab_summary, tab_program, tab_year, tab_agency, tab_year_agency, tab_place, tab_detail = st.tabs(
    ["Summary", "Program-wise", "Year-wise", "Agency-wise", "Year × Agency", "Placement", "Detailed"]
)

# -----------------------------
# Summary
# -----------------------------
with tab_summary:
    total_students = int(f["GrandTotal"].sum())
    passout = int(f["Passout_Total"].sum())
    continuation = int(f["Continuation_Total"].sum())
    cancelled = int(f["Cancelled"].sum())
    validity = int(f["Validity_Expired"].sum())
    iteration_rate = (validity / total_students * 100) if total_students else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Students", f"{total_students}")
    c2.metric("Passout (incl. Convocation)", f"{passout}")
    c3.metric("Continuation (All Buckets)", f"{continuation}")
    c4.metric("Admission Cancelled", f"{cancelled}")
    c5.metric("Validity Expired (Iteration)", f"{validity}")
    c6.metric("Iteration Rate (%)", f"{iteration_rate:.2f}%")

    st.divider()

    # Top agencies chart (by total)
    g_ag = group_block(f, ["Marketing Agency"]).sort_values("GrandTotal", ascending=False).head(15)
    st.subheader("Top Agencies by Total Students")
    bar_chart(g_ag, "Marketing Agency", "GrandTotal", "Top Agencies (Total Students)")

# -----------------------------
# Program-wise
# -----------------------------
with tab_program:
    g = group_block(f, ["Programme"]).sort_values("GrandTotal", ascending=False)
    st.subheader("Program-wise Summary")
    st.dataframe(g, use_container_width=True)

# -----------------------------
# Year-wise
# -----------------------------
with tab_year:
    g = group_block(f, ["AdmissionYear"]).sort_values("AdmissionYear")
    st.subheader("Year-wise Summary")
    st.dataframe(g, use_container_width=True)

# -----------------------------
# Agency-wise
# -----------------------------
with tab_agency:
    g = group_block(f, ["Marketing Agency"]).sort_values("GrandTotal", ascending=False)
    st.subheader("Agency-wise Summary")
    st.dataframe(g, use_container_width=True)

# -----------------------------
# Year × Agency (NEW)
# -----------------------------
with tab_year_agency:
    g = group_block(f, ["AdmissionYear", "Marketing Agency"]).sort_values(["AdmissionYear","GrandTotal"], ascending=[True,False])
    st.subheader("Year-wise + Agency-wise Summary (NEW)")
    st.dataframe(g, use_container_width=True)

    st.subheader("Top Agencies per Year (Total Students)")
    # show a chart for a selected year
    yr = st.selectbox("Select Year", sorted(g["AdmissionYear"].unique().tolist()))
    gy = g[g["AdmissionYear"] == yr].sort_values("GrandTotal", ascending=False).head(15)
    bar_chart(gy, "Marketing Agency", "GrandTotal", f"Top Agencies in {yr}")

# -----------------------------
# Placement
# -----------------------------
with tab_place:
    st.subheader("Placement Eligible (Rule-based by Semester Buckets)")
    g = group_block(f, ["Programme"]).sort_values("Placement_Eligible", ascending=False)
    st.dataframe(g[["Programme","Placement_Eligible","Placement_Eligible_%","GrandTotal"]], use_container_width=True)

# -----------------------------
# Detailed
# -----------------------------
with tab_detail:
    st.subheader("Detailed / Audit View (Row level)")
    show_cols = [
        "Programme","Batch","AdmissionYear","Marketing Agency",
        "GrandTotal","Cancelled","Passout_Total","Continuation_Total","Validity_Expired",
        "Iteration_Rate_%","CourseDuration","MaxAllowedYears","YearsSinceAdmission",
        "Eligible_By_Rule","Placement_Eligible","Placement_Eligible_%"
    ]
    st.dataframe(f[show_cols].fillna(""), use_container_width=True)
