# app.py
# Run: streamlit run app.py

import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Student Exam Status Dashboard", layout="wide")


# -----------------------------
# Helpers
# -----------------------------
def num(v) -> float:
    """Convert any cell to float number safely."""
    try:
        if pd.isna(v):
            return 0.0
        return float(pd.to_numeric(v, errors="coerce") or 0)
    except Exception:
        return 0.0


def load_and_parse_csv(file_path: str) -> pd.DataFrame:
    """
    Parses the report-style CSV:
    - First column contains Program name sometimes (e.g., BBA(13015)...)
    - Batch is in 'Unnamed: 1' like 2020-JUL
    - Cancelled in 'Unnamed: 2'
    - Total passout in 'Unnamed: 5'
    - Grand Total in 'Unnamed: 6'
    IMPORTANT: Program name + first batch are on SAME row in your file.
    """
    raw = pd.read_csv(file_path)

    col0 = raw.columns[0]

    COL_BATCH = "Unnamed: 1"
    COL_CANCELLED = "Unnamed: 2"
    COL_TOTAL_PASSOUT = "Unnamed: 5"
    COL_GRAND_TOTAL = "Unnamed: 6"

    # Batch format like 2020-JUL, 2021-JAN
    batch_pat = re.compile(r"^\d{4}-[A-Z]{3}$")

    current_program = None
    rows = []

    for _, r in raw.iterrows():
        first = "" if pd.isna(r.get(col0)) else str(r.get(col0)).strip()
        batch = "" if pd.isna(r.get(COL_BATCH)) else str(r.get(COL_BATCH)).strip()

        # Identify a program row (program name appears in first column)
        looks_like_program = (
            first
            and first.lower() not in {"programme", "no. of students"}
            and "(" in first
            and ")" in first
        )

        # If this row contains a program name, set it
        if looks_like_program:
            current_program = first

        # If this row has a valid batch and we already know program, record it
        if current_program and batch_pat.match(batch):
            cancelled = num(r.get(COL_CANCELLED))
            passout = num(r.get(COL_TOTAL_PASSOUT))
            grand_total = num(r.get(COL_GRAND_TOTAL))

            enrolled = max(grand_total - cancelled - passout, 0)
            year = int(batch.split("-")[0])

            rows.append(
                {
                    "Program": current_program,
                    "Batch": batch,
                    "Year": year,
                    "Cancelled": cancelled,
                    "Passout": passout,
                    "Enrolled": enrolled,
                    "GrandTotal": grand_total,
                }
            )

    tidy = pd.DataFrame(rows)
    return tidy


def add_percentage_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if df.empty:
        return df

    gt = df["GrandTotal"].replace(0, pd.NA)

    df["Cancelled_%"] = (df["Cancelled"] / gt) * 100
    df["Passout_%"] = (df["Passout"] / gt) * 100
    df["Enrolled_%"] = (df["Enrolled"] / gt) * 100

    df[["Cancelled_%", "Passout_%", "Enrolled_%"]] = (
        df[["Cancelled_%", "Passout_%", "Enrolled_%"]].fillna(0).round(2)
    )
    return df


def summarize(df: pd.DataFrame, group_cols):
    if df.empty:
        return df
    s = (
        df.groupby(group_cols)[["Cancelled", "Passout", "Enrolled", "GrandTotal"]]
        .sum()
        .reset_index()
    )
    return add_percentage_columns(s)


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("📊 Student Admission & Exam Status Dashboard")

FILE_PATH = "Student Profile(Exam Status).csv"

try:
    data = load_and_parse_csv(FILE_PATH)
except Exception as e:
    st.error(f"Error reading CSV: {e}")
    st.stop()

if data.empty:
    st.error("No data parsed from CSV. Please confirm the CSV format is same as shared.")
    st.stop()

data = add_percentage_columns(data)

# Sidebar filters
st.sidebar.header("Filters")

programs = sorted(data["Program"].unique().tolist())
years = sorted(data["Year"].unique().tolist())

sel_programs = st.sidebar.multiselect("Program", programs, default=programs)
sel_years = st.sidebar.multiselect("Year", years, default=years)

filtered = data[(data["Program"].isin(sel_programs)) & (data["Year"].isin(sel_years))].copy()

# -----------------------------
# Overall Summary
# -----------------------------
st.subheader("Overall Summary (Counts + %)")

tot = filtered[["Cancelled", "Passout", "Enrolled", "GrandTotal"]].sum()
gt = tot["GrandTotal"] if tot["GrandTotal"] else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Grand Total Students", int(tot["GrandTotal"]))
c2.metric("Cancelled Admission", int(tot["Cancelled"]), f"{(tot['Cancelled']/gt*100 if gt else 0):.2f}%")
c3.metric("Total Passout", int(tot["Passout"]), f"{(tot['Passout']/gt*100 if gt else 0):.2f}%")
c4.metric("Currently Enrolled", int(tot["Enrolled"]), f"{(tot['Enrolled']/gt*100 if gt else 0):.2f}%")

st.divider()

# -----------------------------
# Year-wise
# -----------------------------
st.subheader("📅 Year-wise Analysis")
year_summary = summarize(filtered, ["Year"]).sort_values("Year")
st.dataframe(year_summary, use_container_width=True)
st.bar_chart(year_summary.set_index("Year")[["Cancelled_%", "Passout_%", "Enrolled_%"]])

st.divider()

# -----------------------------
# Program-wise
# -----------------------------
st.subheader("🎓 Program-wise Analysis")
prog_summary = summarize(filtered, ["Program"]).sort_values("Program")
st.dataframe(prog_summary, use_container_width=True, height=420)
st.bar_chart(prog_summary.set_index("Program")[["Cancelled_%", "Passout_%", "Enrolled_%"]])

st.divider()

# -----------------------------
# Program × Year
# -----------------------------
st.subheader("📘 Program × Year Combined")
py_summary = summarize(filtered, ["Program", "Year"]).sort_values(["Program", "Year"])
st.dataframe(py_summary, use_container_width=True)

with st.expander("🔍 Batch-wise Raw Data"):
    st.dataframe(filtered, use_container_width=True)
