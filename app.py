# app.py (V3 - Agency-wise + Placement-wise)
# Run: streamlit run app.py

import re
import pandas as pd
import streamlit as st

CURRENT_YEAR = 2025

st.set_page_config(page_title="Student Admission & Exam Status Dashboard", layout="wide")


# -----------------------------
# Helpers
# -----------------------------
def to_num(x) -> float:
    try:
        if pd.isna(x):
            return 0.0
        return float(pd.to_numeric(x, errors="coerce") or 0.0)
    except Exception:
        return 0.0


def normalize_program_key(program: str) -> str:
    if not isinstance(program, str):
        return ""
    return program.split("(")[0].strip().upper()


def find_header_row(raw: pd.DataFrame) -> int:
    # In your report, "Batch" appears in column Unnamed: 1
    for i in range(len(raw)):
        v = raw.iloc[i].get("Unnamed: 1", "")
        if isinstance(v, str) and v.strip().lower() == "batch":
            return i
    return -1


def read_uploaded(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        # Default sheet for new file
        raw = pd.read_excel(uploaded_file, sheet_name="Exam Status")
        return raw
    elif name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    else:
        raise ValueError("Unsupported file type. Upload .csv or .xlsx")


def parse_report(uploaded_file) -> pd.DataFrame:
    raw = read_uploaded(uploaded_file)

    header_row = find_header_row(raw)
    if header_row == -1:
        raise ValueError("Could not find the header row that contains 'Batch'.")

    headers = raw.iloc[header_row].tolist()
    headers = [str(x).strip() if not pd.isna(x) else "" for x in headers]
    if headers[0] == "":
        headers[0] = "Programme"

    df = raw.iloc[header_row + 1 :].copy()
    df.columns = headers

    # Column names (as in your sheet)
    col_batch = "Batch"
    col_agency = "Marketing Agency"

    col_cancel = "01. ADMISSION CANCELLED"
    col_passout = "02. PASSOUT"
    col_passout_conv = "02. PASSOUT & CONVOCATION"
    col_pursuing = "03. PURSUING"

    col_cont1 = "04. Programme Continuation (1 Sem Pending)"
    col_cont2 = "05. Programme Continuation (2 Sem Pending)"
    col_final_pending = "06. FINAL SEM PENDING"
    col_last_backlog = "07. LAST SEM (BACKLOG)"
    col_cont3 = "08. Programme Continuation (3 Sem Pending)"
    col_cont4 = "09. Programme Continuation (4 Sem Pending)"
    col_cont5 = "10. Programme Continuation (5 Sem Pending)"
    col_cont6 = "11. Programme Continuation (6 Sem Pending)"

    col_validity = "13. VALIDITY EXPIRED"
    col_total = "Grand Total"

    needed = [
        col_batch, col_agency,
        col_cancel, col_passout, col_passout_conv, col_pursuing,
        col_cont1, col_cont2, col_final_pending, col_last_backlog,
        col_cont3, col_cont4, col_cont5, col_cont6,
        col_validity, col_total
    ]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected column(s): {missing}")

    # Batch cells are merged in Excel → many agency rows have blank Batch
    batch_pat = re.compile(r"^\d{4}-[A-Z]{3}$")

    current_program = None
    current_batch = None

    rows = []

    for _, r in df.iterrows():
        prog_cell = r.get(df.columns[0], "")
        prog_cell = "" if pd.isna(prog_cell) else str(prog_cell).strip()

        batch = r.get(col_batch, "")
        batch = "" if pd.isna(batch) else str(batch).strip()

        agency = r.get(col_agency, "")
        agency = "" if pd.isna(agency) else str(agency).strip()

        # Detect program header row
        if prog_cell and "(" in prog_cell and ")" in prog_cell:
            current_program = prog_cell
            current_batch = None  # reset when program changes

        # Detect batch (if present)
        if batch_pat.match(batch):
            current_batch = batch

        # We create a record ONLY if we have:
        # program + batch (possibly forward-filled) + agency
        if current_program and current_batch and agency:
            adm_year = int(current_batch.split("-")[0])

            cancelled = to_num(r.get(col_cancel))
            passout = to_num(r.get(col_passout))
            passout_conv = to_num(r.get(col_passout_conv))
            pursuing = to_num(r.get(col_pursuing))

            cont1 = to_num(r.get(col_cont1))
            cont2 = to_num(r.get(col_cont2))
            final_pending = to_num(r.get(col_final_pending))
            last_backlog = to_num(r.get(col_last_backlog))
            cont3 = to_num(r.get(col_cont3))
            cont4 = to_num(r.get(col_cont4))
            cont5 = to_num(r.get(col_cont5))
            cont6 = to_num(r.get(col_cont6))

            validity_expired = to_num(r.get(col_validity))
            grand_total = to_num(r.get(col_total))

            # Some agency rows may have blanks for totals; skip if grand_total is 0 and all buckets are 0
            if grand_total == 0 and (cancelled + passout + passout_conv + pursuing + cont1 + cont2 + final_pending +
                                    last_backlog + cont3 + cont4 + cont5 + cont6 + validity_expired) == 0:
                continue

            rows.append(
                {
                    "Program": current_program,
                    "ProgramKey": normalize_program_key(current_program),
                    "Agency": agency,
                    "Batch": current_batch,
                    "AdmissionYear": adm_year,
                    "Cancelled": cancelled,
                    "Passout": passout,
                    "Passout_Convocation": passout_conv,
                    "Pursuing": pursuing,
                    "Cont_1Sem": cont1,
                    "Cont_2Sem": cont2,
                    "FinalSem_Pending": final_pending,
                    "LastSem_Backlog": last_backlog,
                    "Cont_3Sem": cont3,
                    "Cont_4Sem": cont4,
                    "Cont_5Sem": cont5,
                    "Cont_6Sem": cont6,
                    "Validity_Expired": validity_expired,
                    "GrandTotal": grand_total,
                }
            )

    tidy = pd.DataFrame(rows)
    if tidy.empty:
        raise ValueError("Parsed 0 rows. Please verify that the file matches the Exam Status report structure.")

    # Derived totals
    tidy["Passout_Total"] = tidy["Passout"] + tidy["Passout_Convocation"]

    tidy["Continuation_Total"] = (
        tidy["Pursuing"] +
        tidy["Cont_1Sem"] + tidy["Cont_2Sem"] +
        tidy["FinalSem_Pending"] + tidy["LastSem_Backlog"] +
        tidy["Cont_3Sem"] + tidy["Cont_4Sem"] + tidy["Cont_5Sem"] + tidy["Cont_6Sem"]
    )

    # Iteration (VALIDITY EXPIRED)
    tidy["Iteration_Rate_%"] = tidy.apply(
        lambda r: (r["Validity_Expired"] / r["GrandTotal"] * 100) if r["GrandTotal"] else 0.0,
        axis=1,
    )

    # Placement eligibility (5th sem onwards = 2 sem pending or less incl backlog)
    tidy["Placement_Eligible"] = (
        tidy["Cont_2Sem"] +
        tidy["Cont_1Sem"] +
        tidy["FinalSem_Pending"] +
        tidy["LastSem_Backlog"]
    )
    tidy["Placement_Eligible_%"] = tidy.apply(
        lambda r: (r["Placement_Eligible"] / r["GrandTotal"] * 100) if r["GrandTotal"] else 0.0,
        axis=1,
    )

    # Data consistency
    tidy["Total_Check_Sum"] = tidy["Cancelled"] + tidy["Passout_Total"] + tidy["Continuation_Total"] + tidy["Validity_Expired"]
    tidy["Unaccounted"] = tidy["GrandTotal"] - tidy["Total_Check_Sum"]

    return tidy


def add_percent_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    gt = df["GrandTotal"].replace(0, pd.NA)

    df["Cancelled_%"] = (df["Cancelled"] / gt) * 100
    df["Passout_%"] = (df["Passout_Total"] / gt) * 100
    df["Continuation_%"] = (df["Continuation_Total"] / gt) * 100
    df["Validity_Expired_%"] = (df["Validity_Expired"] / gt) * 100

    df[["Cancelled_%", "Passout_%", "Continuation_%", "Validity_Expired_%"]] = (
        df[["Cancelled_%", "Passout_%", "Continuation_%", "Validity_Expired_%"]].fillna(0).round(2)
    )

    # Placement %
    df["Placement_Eligible_%"] = (df["Placement_Eligible"] / gt) * 100
    df["Placement_Eligible_%"] = df["Placement_Eligible_%"].fillna(0).round(2)

    return df


def summarize(df: pd.DataFrame, group_cols):
    s = (
        df.groupby(group_cols, dropna=False)[
            ["Cancelled", "Passout_Total", "Continuation_Total", "Validity_Expired",
             "Placement_Eligible", "GrandTotal"]
        ]
        .sum()
        .reset_index()
    )
    s = add_percent_columns(s)
    s["Iteration_Rate_%"] = s["Validity_Expired_%"]
    return s


# -----------------------------
# UI
# -----------------------------
st.title("📊 Student Admission & Exam Status Dashboard")

with st.sidebar:
    st.header("Upload Data")
    uploaded = st.file_uploader("Upload file (.xlsx or .csv)", type=["xlsx", "xls", "csv"])
    st.caption("Eligibility reference year is fixed at 2025")

if not uploaded:
    st.info("Upload the Excel/CSV to start.")
    st.stop()

try:
    data = parse_report(uploaded)
except Exception as e:
    st.error(f"Failed to parse file: {e}")
    st.stop()

# Filters
with st.sidebar:
    st.header("Filters")
    programs = sorted(data["Program"].unique().tolist())
    years = sorted(data["AdmissionYear"].unique().tolist())
    agencies = sorted(data["Agency"].unique().tolist())

    sel_programs = st.multiselect("Program", programs, default=programs)
    sel_years = st.multiselect("Admission Year", years, default=years)
    sel_agencies = st.multiselect("Agency", agencies, default=agencies)

fdf = data[
    (data["Program"].isin(sel_programs)) &
    (data["AdmissionYear"].isin(sel_years)) &
    (data["Agency"].isin(sel_agencies))
].copy()

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["Summary", "Program-wise", "Year-wise", "Agency-wise", "Placement", "Detailed"]
)

# -----------------------------
# Tab 1: Summary
# -----------------------------
with tab1:
    st.subheader("Summary")

    tot = fdf[["GrandTotal", "Cancelled", "Passout_Total", "Continuation_Total", "Validity_Expired", "Placement_Eligible"]].sum()
    gt = tot["GrandTotal"] if tot["GrandTotal"] else 0

    iteration_rate = (tot["Validity_Expired"] / gt * 100) if gt else 0
    placement_rate = (tot["Placement_Eligible"] / gt * 100) if gt else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Students", int(tot["GrandTotal"]))
    c2.metric("Passout (incl. Convocation)", int(tot["Passout_Total"]), f"{(tot['Passout_Total']/gt*100 if gt else 0):.2f}%")
    c3.metric("Continuation", int(tot["Continuation_Total"]), f"{(tot['Continuation_Total']/gt*100 if gt else 0):.2f}%")
    c4.metric("Admission Cancelled", int(tot["Cancelled"]), f"{(tot['Cancelled']/gt*100 if gt else 0):.2f}%")
    c5.metric("Validity Expired (Iteration)", int(tot["Validity_Expired"]), f"{iteration_rate:.2f}%")
    c6.metric("Placement Eligible", int(tot["Placement_Eligible"]), f"{placement_rate:.2f}%")

# -----------------------------
# Tab 2: Program-wise
# -----------------------------
with tab2:
    st.subheader("Program-wise")
    ps = summarize(fdf, ["Program"]).sort_values("Iteration_Rate_%", ascending=False)
    st.dataframe(ps, use_container_width=True, height=450)
    st.caption("Iteration Rate (%) by Program")
    st.bar_chart(ps.set_index("Program")[["Iteration_Rate_%"]])

# -----------------------------
# Tab 3: Year-wise
# -----------------------------
with tab3:
    st.subheader("Year-wise")
    ys = summarize(fdf, ["AdmissionYear"]).sort_values("AdmissionYear")
    st.dataframe(ys, use_container_width=True)
    st.caption("Placement Eligible (%) by Year")
    st.bar_chart(ys.set_index("AdmissionYear")[["Placement_Eligible_%"]])

# -----------------------------
# Tab 4: Agency-wise
# -----------------------------
with tab4:
    st.subheader("Agency-wise")
    ag = summarize(fdf, ["Agency"]).sort_values("GrandTotal", ascending=False)
    st.dataframe(ag, use_container_width=True, height=450)

    st.caption("Top Agencies by Total Students")
    st.bar_chart(ag.set_index("Agency")[["GrandTotal"]])

# -----------------------------
# Tab 5: Placement
# -----------------------------
with tab5:
    st.subheader("Placement Eligibility View")

    st.markdown(
        """
        **Placement Eligible Logic**  
        Placement Eligible = **2 Sem Pending + 1 Sem Pending + Final Sem Pending + Last Sem Backlog**  
        (i.e., students in **5th semester onwards** or equivalent)
        """
    )

    # Program + Year placement view
    py = summarize(fdf, ["Program", "AdmissionYear"]).sort_values(["Program", "AdmissionYear"])
    st.dataframe(py[["Program", "AdmissionYear", "Placement_Eligible", "Placement_Eligible_%", "GrandTotal"]], use_container_width=True)

    st.caption("Placement Eligible (%) by Program")
    pp = summarize(fdf, ["Program"]).sort_values("Placement_Eligible_%", ascending=False)
    st.bar_chart(pp.set_index("Program")[["Placement_Eligible_%"]])

# -----------------------------
# Tab 6: Detailed
# -----------------------------
with tab6:
    st.subheader("Detailed")
    out = fdf.copy()
    out["Iteration_Rate_%"] = out["Iteration_Rate_%"].round(2)
    out["Placement_Eligible_%"] = out["Placement_Eligible_%"].round(2)

    if (out["Unaccounted"].abs() > 0.0001).any():
        st.warning("Some rows have mismatch (GrandTotal ≠ sum of buckets). Please verify those batches/agencies.")

    st.dataframe(out.sort_values(["Program", "AdmissionYear", "Batch", "Agency"]), use_container_width=True, height=600)
