# app.py (V6 - Correct Vendor logic: Vendor = Marketing Agency from SOE Students, not sheet name)
# Run: streamlit run app.py

import re
import pandas as pd
import streamlit as st

CURRENT_YEAR = 2025
VALIDITY_EXPIRED_LABEL = "13. VALIDITY EXPIRED"

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


def normalize_vendor(v: str) -> str:
    """
    Normalize vendor/agency names so you get correct grouping:
    DL to OL => D2L
    JARO EDUCATION => JARO
    etc.
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "UNKNOWN"

    s = str(v).strip().upper()
    s = re.sub(r"\s+", " ", s)

    # common normalizations
    mapping = {
        "DL TO OL": "D2L",
        "D L TO O L": "D2L",
        "D2L": "D2L",
        "JARO EDUCATION": "JARO",
        "JARO EDU": "JARO",
        "BV MAIN": "BVMAIN",
        "B V MAIN": "BVMAIN",
        "COLLEGE DEKH0": "COLLEGEDEKHO",
        "COLLEGE DEKHO": "COLLEGEDEKHO",
        "TELEPHONY WHATSAPP": "TELEPHONY/WHATSAPP",
        "WHATSAPP": "WHATSAPP",
        "ORGANIC SEARCH": "ORGANIC SEARCH",
        "SOCIAL": "SOCIAL",
        "DIRECT": "DIRECT",
        "OFFLINE": "OFFLINE",
        "WALK IN": "WALK IN",
        "REFERRAL": "REFERRAL",
        "BVDU": "BVDU",
        "TCIL": "TCIL",
        "FACEBOOK": "FACEBOOK",
        "INSTAGRAM": "INSTAGRAM",
        "QR CODE": "QR CODE",
    }

    return mapping.get(s, s)


def find_header_row(raw: pd.DataFrame) -> int:
    for i in range(len(raw)):
        v = raw.iloc[i].get("Unnamed: 1", "")
        if isinstance(v, str) and v.strip().lower() == "batch":
            return i
    return -1


# -----------------------------
# Exam Status Parser (unchanged)
# -----------------------------
def parse_exam_status_report(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        raw = pd.read_excel(uploaded_file, sheet_name="Exam Status")
    elif name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        raise ValueError("Unsupported file type. Upload .csv or .xlsx")

    header_row = find_header_row(raw)
    if header_row == -1:
        raise ValueError("Could not find the header row that contains 'Batch'.")

    headers = raw.iloc[header_row].tolist()
    headers = [str(x).strip() if not pd.isna(x) else "" for x in headers]
    if headers[0] == "":
        headers[0] = "Programme"

    df = raw.iloc[header_row + 1 :].copy()
    df.columns = headers

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

        if prog_cell and "(" in prog_cell and ")" in prog_cell:
            current_program = prog_cell
            current_batch = None

        if batch_pat.match(batch):
            current_batch = batch

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

            if grand_total == 0 and (
                cancelled + passout + passout_conv + pursuing + cont1 + cont2 +
                final_pending + last_backlog + cont3 + cont4 + cont5 + cont6 + validity_expired
            ) == 0:
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

    tidy["Passout_Total"] = tidy["Passout"] + tidy["Passout_Convocation"]

    tidy["Continuation_Total"] = (
        tidy["Pursuing"] +
        tidy["Cont_1Sem"] + tidy["Cont_2Sem"] +
        tidy["FinalSem_Pending"] + tidy["LastSem_Backlog"] +
        tidy["Cont_3Sem"] + tidy["Cont_4Sem"] + tidy["Cont_5Sem"] + tidy["Cont_6Sem"]
    )

    tidy["Iteration_Rate_%"] = tidy.apply(
        lambda r: (r["Validity_Expired"] / r["GrandTotal"] * 100) if r["GrandTotal"] else 0.0,
        axis=1,
    )

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
# SOE Students Loader (UPDATED: bring vendor/agency column)
# -----------------------------
@st.cache_data(show_spinner=False)
def load_soe_students(student_profile_file) -> pd.DataFrame:
    df = pd.read_excel(student_profile_file, sheet_name="SOE Students")

    base_needed = ["Enrollment No", "Exam Status", "Programme", "Batch", "Student Name", "PRN"]
    missing = [c for c in base_needed if c not in df.columns]
    if missing:
        raise ValueError(f"'SOE Students' sheet missing columns: {missing}")

    # vendor column candidates (whatever exists in your file)
    vendor_candidates = [
        "Marketing Agency", "Agency", "Vendor", "Vendor Name", "Source", "Lead Source"
    ]
    vendor_col = None
    for c in vendor_candidates:
        if c in df.columns:
            vendor_col = c
            break

    df = df.copy()
    df["Enrollment No"] = df["Enrollment No"].astype(str).str.strip()
    df["Exam Status"] = df["Exam Status"].astype(str).str.strip()
    df["Programme"] = df["Programme"].astype(str).str.strip()
    df["Batch"] = df["Batch"].astype(str).str.strip()
    df["Student Name"] = df["Student Name"].astype(str).str.strip()
    df["PRN"] = df["PRN"].astype(str).str.strip()

    if vendor_col:
        df["VendorRaw"] = df[vendor_col].astype(str).str.strip()
        df["Vendor"] = df["VendorRaw"].apply(normalize_vendor)
    else:
        df["VendorRaw"] = None
        df["Vendor"] = "UNKNOWN"

    return df


# -----------------------------
# SOE Fees Write Off Parser (no vendor here; vendor comes from SOE Students link)
# -----------------------------
BVP_ENR_PATTERN = re.compile(r"\((BVP\d+)\)")

def _find_particulars_row(df: pd.DataFrame) -> int:
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.strip().str.lower()
        if (row == "particulars").any():
            return i
    return -1


def _parse_sheet_writeoff(sheet_df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    df = sheet_df.copy()

    pr = _find_particulars_row(df)
    if pr == -1:
        return pd.DataFrame()

    particulars_col_idx = None
    for j, val in enumerate(df.iloc[pr].tolist()):
        if isinstance(val, str) and val.strip().lower() == "particulars":
            particulars_col_idx = j
            break
    if particulars_col_idx is None:
        return pd.DataFrame()

    amt_cols = [particulars_col_idx + 1, particulars_col_idx + 2, particulars_col_idx + 3, particulars_col_idx + 4]
    max_col = df.shape[1] - 1
    if any(c > max_col for c in amt_cols):
        return pd.DataFrame()

    sno_col = 0
    records = []

    for i in range(pr + 1, len(df)):
        sno = df.iloc[i, sno_col]
        particulars = df.iloc[i, particulars_col_idx]

        if pd.isna(particulars):
            continue

        particulars_s = str(particulars).strip()
        if particulars_s.lower() == "grand total":
            break

        if pd.isna(sno) and not BVP_ENR_PATTERN.search(particulars_s):
            continue

        opening = df.iloc[i, amt_cols[0]]
        debit = df.iloc[i, amt_cols[1]]
        credit = df.iloc[i, amt_cols[2]]
        closing = df.iloc[i, amt_cols[3]]

        enr = None
        m = BVP_ENR_PATTERN.search(particulars_s)
        if m:
            enr = m.group(1)

        records.append(
            {
                "Sheet": sheet_name,
                "Enrollment No": enr,
                "Particulars": particulars_s,
                "OpeningBalance": to_num(opening),
                "Debit": to_num(debit),
                "Credit": to_num(credit),
                "ClosingBalance": to_num(closing),
            }
        )

    out = pd.DataFrame(records)
    out = out[~out["Enrollment No"].isna()].copy()
    return out


@st.cache_data(show_spinner=False)
def parse_writeoff_workbook(writeoff_file) -> pd.DataFrame:
    xl = pd.ExcelFile(writeoff_file)
    all_parts = []
    for sh in xl.sheet_names:
        sdf = pd.read_excel(writeoff_file, sheet_name=sh)
        t = _parse_sheet_writeoff(sdf, sh)
        if not t.empty:
            all_parts.append(t)
    if not all_parts:
        return pd.DataFrame()
    return pd.concat(all_parts, ignore_index=True)


# -----------------------------
# UI
# -----------------------------
st.title("📊 Student Admission & Exam Status Dashboard")

with st.sidebar:
    st.header("Upload Data")
    exam_file = st.file_uploader(
        "1) Upload Exam Status file (Student Profile.xlsx)",
        type=["xlsx", "xls", "csv"],
        key="exam_file"
    )
    writeoff_file = st.file_uploader(
        "2) Upload SOE Fees Write Off file (.xlsx)",
        type=["xlsx", "xls"],
        key="writeoff_file"
    )
    st.caption("Eligibility reference year is fixed at 2025")

if not exam_file:
    st.info("Upload the Exam Status file (Student Profile.xlsx) to start.")
    st.stop()

try:
    data = parse_exam_status_report(exam_file)
except Exception as e:
    st.error(f"Failed to parse Exam Status file: {e}")
    st.stop()

try:
    soe_students = load_soe_students(exam_file)
except Exception as e:
    st.warning(f"SOE Students linking is unavailable: {e}")
    soe_students = None

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

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    ["Summary", "Program-wise", "Year-wise", "Agency-wise", "Placement", "Detailed", "SOE Fees Write Off"]
)

# -----------------------------
# Tab 1
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
# Tab 2
# -----------------------------
with tab2:
    st.subheader("Program-wise")
    ps = summarize(fdf, ["Program"]).sort_values("Iteration_Rate_%", ascending=False)
    st.dataframe(ps, use_container_width=True, height=450)
    st.caption("Iteration Rate (%) by Program")
    st.bar_chart(ps.set_index("Program")[["Iteration_Rate_%"]])

# -----------------------------
# Tab 3
# -----------------------------
with tab3:
    st.subheader("Year-wise")
    ys = summarize(fdf, ["AdmissionYear"]).sort_values("AdmissionYear")
    st.dataframe(ys, use_container_width=True)
    st.caption("Placement Eligible (%) by Year")
    st.bar_chart(ys.set_index("AdmissionYear")[["Placement_Eligible_%"]])

# -----------------------------
# Tab 4
# -----------------------------
with tab4:
    st.subheader("Agency-wise")
    ag = summarize(fdf, ["Agency"]).sort_values("GrandTotal", ascending=False)
    st.dataframe(ag, use_container_width=True, height=450)
    st.caption("Top Agencies by Total Students")
    st.bar_chart(ag.set_index("Agency")[["GrandTotal"]])

# -----------------------------
# Tab 5
# -----------------------------
with tab5:
    st.subheader("Placement Eligibility View")
    st.markdown(
        """
        **Placement Eligible Logic**  
        Placement Eligible = **2 Sem Pending + 1 Sem Pending + Final Sem Pending + Last Sem Backlog**
        """
    )
    py = summarize(fdf, ["Program", "AdmissionYear"]).sort_values(["Program", "AdmissionYear"])
    st.dataframe(py[["Program", "AdmissionYear", "Placement_Eligible", "Placement_Eligible_%", "GrandTotal"]], use_container_width=True)

# -----------------------------
# Tab 6
# -----------------------------
with tab6:
    st.subheader("Detailed")
    out = fdf.copy()
    out["Iteration_Rate_%"] = out["Iteration_Rate_%"].round(2)
    out["Placement_Eligible_%"] = out["Placement_Eligible_%"].round(2)

    if (out["Unaccounted"].abs() > 0.0001).any():
        st.warning("Some rows have mismatch (GrandTotal ≠ sum of buckets). Please verify those batches/agencies.")

    st.dataframe(out.sort_values(["Program", "AdmissionYear", "Batch", "Agency"]), use_container_width=True, height=600)

# -----------------------------
# Tab 7 (UPDATED)
# -----------------------------
with tab7:
    st.subheader("SOE Fees Write Off (Linked to Validity Expired)")

    if not writeoff_file:
        st.info("Upload the SOE Fees Write Off (.xlsx) file to view this analysis.")
        st.stop()

    if soe_students is None:
        st.warning("Cannot link write-off to students because 'SOE Students' sheet is not available.")
        st.stop()

    try:
        wdf = parse_writeoff_workbook(writeoff_file)
    except Exception as e:
        st.error(f"Failed to parse SOE Fees Write Off file: {e}")
        st.stop()

    if wdf.empty:
        st.warning("No parsable rows found in the write-off workbook.")
        st.stop()

    # Validity Expired students
    vs = soe_students[soe_students["Exam Status"] == VALIDITY_EXPIRED_LABEL].copy()

    # Link (now includes Vendor)
    linked = wdf.merge(
        vs[["Enrollment No", "Student Name", "PRN", "Programme", "Batch", "Exam Status", "Vendor", "VendorRaw"]],
        on="Enrollment No",
        how="inner"
    )

    # show detected vendor list (this is what you wanted: identify & mention them)
    st.markdown("### Detected Vendors / Agencies (from SOE Students)")
    vend_list = (
        linked.groupby(["Vendor"], dropna=False)["Enrollment No"]
        .nunique()
        .reset_index(name="Unique Students")
        .sort_values("Unique Students", ascending=False)
    )
    st.dataframe(vend_list, use_container_width=True, height=280)

    # Metrics
    total_rows = len(wdf)
    total_students_linkable = wdf["Enrollment No"].nunique()
    total_closing = wdf["ClosingBalance"].sum()

    linked_students = linked["Enrollment No"].nunique()
    linked_closing = linked["ClosingBalance"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Write-off rows (linkable)", int(total_rows))
    c2.metric("Unique Enrollment Nos", int(total_students_linkable))
    c3.metric("Total Closing Balance", f"{total_closing:,.0f}")
    c4.metric("Validity Expired: Closing Balance", f"{linked_closing:,.0f}")

    st.caption("Vendor is taken from SOE Students (Marketing Agency/Agency/Vendor column).")

    # A) Program-wise / Batch-wise / Vendor-wise (FULL)
    st.markdown("### Validity Expired: Programme + Batch + Vendor (Program-wise + Batch-wise + Vendor-wise)")
    pbv = (
        linked.groupby(["Programme", "Batch", "Vendor"], dropna=False)[["ClosingBalance", "Debit", "Credit"]]
        .sum()
        .reset_index()
        .sort_values("ClosingBalance", ascending=False)
    )
    st.dataframe(pbv, use_container_width=True, height=480)

    # B) Vendor-wise only
    st.markdown("### Vendor-wise Closing Balance")
    vw = (
        linked.groupby(["Vendor"], dropna=False)[["ClosingBalance", "Debit", "Credit"]]
        .sum()
        .reset_index()
        .sort_values("ClosingBalance", ascending=False)
    )
    st.dataframe(vw, use_container_width=True, height=320)

    # C) Programme-wise only
    st.markdown("### Programme-wise Closing Balance")
    pw = (
        linked.groupby(["Programme"], dropna=False)[["ClosingBalance", "Debit", "Credit"]]
        .sum()
        .reset_index()
        .sort_values("ClosingBalance", ascending=False)
    )
    st.dataframe(pw, use_container_width=True, height=320)

    # Top students
    st.markdown("### Top Validity Expired Students by Closing Balance")
    topn = (
        linked.groupby(["Enrollment No", "Student Name", "Programme", "Batch", "Vendor"], dropna=False)[["ClosingBalance", "Debit", "Credit"]]
        .sum()
        .reset_index()
        .sort_values("ClosingBalance", ascending=False)
        .head(50)
    )
    st.dataframe(topn, use_container_width=True, height=520)

    with st.expander("Show linked raw rows"):
        st.dataframe(
            linked.sort_values(["Vendor", "Programme", "Batch", "ClosingBalance"], ascending=[True, True, True, False]),
            use_container_width=True,
            height=650
        )
