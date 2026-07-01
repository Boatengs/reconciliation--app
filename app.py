"""
app.py
Internal tool: upload billing file + rate table (+ optional rectification file),
get back a validated reconciliation workbook and a summary report.

Run locally:   streamlit run app.py
Deploy free:   push this folder to a GitHub repo, then deploy on
               https://share.streamlit.io (Streamlit Community Cloud).
"""

import io
from datetime import date

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Pt

from reconcile import load_billing, load_rates, load_rectification, run_reconciliation, build_summary

st.set_page_config(page_title="Cigna Premium Reconciliation", layout="wide")
st.title("Insurance Premium Reconciliation")
st.caption("Internal tool — upload billing + rate files, get a validated workbook and report back.")

# ---------------------------------------------------------------------------
# Sidebar inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("1. Files")
    billing_file = st.file_uploader("Billing file (.xlsx)", type=["xlsx"])
    rates_file = st.file_uploader("Rate table (.xlsx)", type=["xlsx"])
    rect_file = st.file_uploader("Rectification file (.xlsx) — optional", type=["xlsx"])

    st.header("2. Settings")
    rate_is_annual = st.radio(
        "The rate table figures are:",
        ["Annual premiums", "Monthly premiums"],
        help="If unsure, run once, check the warning banner that appears if the "
             "wrong setting produces a large uniform gap across most records.",
    ) == "Annual premiums"
    tolerance = st.slider("Tolerance % (flags CHECK when |gap| exceeds this)", 1, 50, 10)

    period_label = st.text_input("Period label (for the report title)", value="")

    run_btn = st.button("Run reconciliation", type="primary", use_container_width=True)

# ---------------------------------------------------------------------------
# Excel builder
# ---------------------------------------------------------------------------
def build_excel(validation_df, summary, warning, rect_df, tolerance, rate_is_annual):
    wb = Workbook()
    wb.remove(wb.active)

    green_font = Font(name="Calibri", sz=12, b=True, color="FF006100")
    green_fill = PatternFill("solid", fgColor="FFC6EFCE")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_top = Alignment(horizontal="center", vertical="top")
    data_font = Font(name="Calibri", sz=11)
    bold_font = Font(name="Calibri", sz=11, b=True)
    num_fmt = "#,##0.00"

    def style_header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = green_font
            cell.fill = green_fill
            cell.border = border
            cell.alignment = center_top
        ws.row_dimensions[1].height = 16

    # ---- Validation ----
    ws = wb.create_sheet("Validation")
    headers = ["Unique Identifier", "Category", "Age range", "Country", "N of days insured",
               "Code", "RatePremiumUSD", "ExpectedUSD", "Premium medical", "DifferenceUSD",
               "PercentDiff", "Status", "Reason"]
    ws.append(headers)
    style_header(ws, len(headers))
    for _, row in validation_df.iterrows():
        ws.append([
            row["UID"], row["Category"], row["AgeRange"], row["Country"], row["Days"],
            row["Code"], row["Rate"], round(row["ExpectedUSD"], 2), round(row["PremiumBilled"], 2),
            round(row["DifferenceUSD"], 2),
            (None if pd.isna(row["PercentDiff"]) else row["PercentDiff"]),
            row["Status"], row["Reason"],
        ])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).font = data_font
            if c in (7, 8, 9, 10, 11):
                ws.cell(row=r, column=c).number_format = num_fmt
    widths = [22, 12, 12, 14, 14, 14, 16, 14, 14, 14, 12, 10, 32]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"

    # ---- Summary ----
    ws2 = wb.create_sheet("Summary")
    ws2.append(["Total Rows", "OK", "CHECK", "No Rate", "Total Billed", "Total Expected",
                "Total Difference", "Avg % Diff (OK only)", "Tolerance %", "Rates are"])
    style_header(ws2, 10)
    ws2.append([
        summary["total_rows"], summary["ok"], summary["check"], summary["no_rate"],
        round(summary["total_billed"], 2), round(summary["total_expected"], 2),
        round(summary["total_diff"], 2),
        (None if summary["avg_ok_pct"] is None else round(summary["avg_ok_pct"], 2)),
        tolerance, ("Annual" if rate_is_annual else "Monthly"),
    ])
    for c in range(1, 11):
        ws2.cell(row=2, column=c).font = data_font
        if c in (5, 6, 7, 8):
            ws2.cell(row=2, column=c).number_format = num_fmt
    for i, w in enumerate([12, 8, 10, 10, 16, 16, 16, 18, 12, 12], start=1):
        ws2.column_dimensions[ws2.cell(row=1, column=i).column_letter].width = w
    if warning:
        ws2["A4"] = "Sanity check warning:"
        ws2["A4"].font = Font(name="Calibri", sz=11, b=True, color="FF9C0006")
        ws2["A5"] = warning
        ws2["A5"].font = Font(name="Calibri", sz=11, color="FF9C0006")
        ws2["A5"].alignment = Alignment(wrap_text=True, vertical="top")
        ws2.merge_cells("A5:J8")

    # ---- Exceptions ----
    ws3 = wb.create_sheet("Exceptions")
    ws3.append(headers)
    style_header(ws3, len(headers))
    exc = validation_df[validation_df["Status"] == "CHECK"]
    for _, row in exc.iterrows():
        ws3.append([
            row["UID"], row["Category"], row["AgeRange"], row["Country"], row["Days"],
            row["Code"], row["Rate"], round(row["ExpectedUSD"], 2), round(row["PremiumBilled"], 2),
            round(row["DifferenceUSD"], 2), row["PercentDiff"], row["Status"], row["Reason"],
        ])
    if len(exc) == 0:
        ws3["A2"] = "No exceptions found"
        ws3["A2"].font = data_font
    for r in range(2, ws3.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws3.cell(row=r, column=c).font = data_font
    for i, w in enumerate(widths, start=1):
        ws3.column_dimensions[ws3.cell(row=1, column=i).column_letter].width = w

    # ---- Rectification (if provided) ----
    if rect_df is not None:
        ws4 = wb.create_sheet("Rectification")
        ws4.append(["Unique Identifier", "Premium (rectified)", "Previous Due", "Difference"])
        style_header(ws4, 4)
        for _, row in rect_df.iterrows():
            ws4.append([row["UID"], round(row["PremiumRect"], 2), round(row["PreviousDue"], 2), round(row["Diff"], 2)])
        r = ws4.max_row + 1
        ws4.cell(row=r, column=1, value="TOTAL").font = bold_font
        for c, col in zip((2, 3, 4), ("B", "C", "D")):
            cell = ws4.cell(row=r, column=c, value=f"=SUM({col}2:{col}{r-1})")
            cell.font = bold_font
            cell.number_format = num_fmt
        for i, w in enumerate([22, 18, 16, 16], start=1):
            ws4.column_dimensions[ws4.cell(row=1, column=i).column_letter].width = w
        for rr in range(2, r):
            for c in range(2, 5):
                ws4.cell(row=rr, column=c).number_format = num_fmt
                ws4.cell(row=rr, column=c).font = data_font
            ws4.cell(row=rr, column=1).font = data_font

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Report builder
# ---------------------------------------------------------------------------
def build_report(summary, warning, rect_df, tolerance, rate_is_annual, period_label):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading("Premium Reconciliation Report", level=1)
    title.alignment = 1
    if period_label:
        sub = doc.add_paragraph(period_label)
        sub.alignment = 1

    doc.add_heading("Key figures", level=2)
    doc.add_paragraph(f"Billed total: {summary['total_billed']:,.2f} USD", style="List Bullet")
    doc.add_paragraph(f"Expected total: {summary['total_expected']:,.2f} USD", style="List Bullet")
    doc.add_paragraph(f"Difference: {summary['total_diff']:,.2f} USD", style="List Bullet")
    doc.add_paragraph(
        f"Records: {summary['total_rows']} — OK: {summary['ok']} — CHECK: {summary['check']} "
        f"— No Rate: {summary['no_rate']} — Tolerance: \u00b1{tolerance}%", style="List Bullet")
    doc.add_paragraph(f"Rate table treated as: {'Annual' if rate_is_annual else 'Monthly'} premiums", style="List Bullet")

    if warning:
        doc.add_heading("Warning", level=2)
        p = doc.add_paragraph(warning)
        p.runs[0].font.color.rgb = None  # keep default; Word will render plain

    if rect_df is not None:
        doc.add_heading("Rectification", level=2)
        total_rect = rect_df["PremiumRect"].sum()
        total_prev = rect_df["PreviousDue"].sum()
        total_diff = rect_df["Diff"].sum()
        doc.add_paragraph(f"Rectification premium total: {total_rect:,.2f} USD", style="List Bullet")
        doc.add_paragraph(f"Previous due total: {total_prev:,.2f} USD", style="List Bullet")
        doc.add_paragraph(f"Difference (new charges): {total_diff:,.2f} USD", style="List Bullet")
        new_charge_rows = rect_df[(rect_df["PreviousDue"] == 0) & (rect_df["PremiumRect"] != 0)]
        if len(new_charge_rows) > 0:
            doc.add_paragraph(
                f"{len(new_charge_rows)} record(s) show no previous due on file — worth a quick check "
                f"to confirm these are legitimately new charges.", style="List Bullet")

    doc.add_heading("Conclusion", level=2)
    if summary["check"] == 0 and summary["no_rate"] == 0:
        doc.add_paragraph(
            "All records reconcile within tolerance. Billed figures are accurate to the bill "
            "and consistent with expected premium.")
    else:
        doc.add_paragraph(
            f"{summary['check']} of {summary['total_rows']} records fall outside tolerance"
            + (f" and {summary['no_rate']} have no matching rate." if summary["no_rate"] else ".")
            + " Review the Exceptions sheet before approving payment on the affected records."
        )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------
if run_btn:
    if not billing_file or not rates_file:
        st.error("Please upload at least a billing file and a rate table.")
        st.stop()

    with st.spinner("Reading files..."):
        try:
            billing_df = load_billing(billing_file)
            rates_df = load_rates(rates_file)
            rect_df = load_rectification(rect_file) if rect_file else None
        except Exception as e:
            st.error(f"Could not read one of the files: {e}")
            st.stop()

    with st.spinner("Reconciling..."):
        validation_df, warning = run_reconciliation(billing_df, rates_df, tolerance, rate_is_annual)
        summary = build_summary(validation_df)

    st.success(f"Done — {summary['total_rows']} records reviewed.")

    if warning:
        st.warning(warning)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Billed total", f"${summary['total_billed']:,.2f}")
    c2.metric("Expected total", f"${summary['total_expected']:,.2f}")
    c3.metric("Difference", f"${summary['total_diff']:,.2f}")
    c4.metric("Flagged (CHECK)", f"{summary['check']} / {summary['total_rows']}")

    st.subheader("Validation detail")
    st.dataframe(validation_df, use_container_width=True, height=350)

    if rect_df is not None:
        st.subheader("Rectification detail")
        st.dataframe(rect_df, use_container_width=True, height=200)

    excel_buf = build_excel(validation_df, summary, warning, rect_df, tolerance, rate_is_annual)
    report_buf = build_report(summary, warning, rect_df, tolerance, rate_is_annual, period_label)

    dcol1, dcol2 = st.columns(2)
    dcol1.download_button(
        "Download reconciled workbook (.xlsx)", excel_buf,
        file_name=f"Reconciliation_{date.today().isoformat()}.xlsx",
        use_container_width=True,
    )
    dcol2.download_button(
        "Download report (.docx)", report_buf,
        file_name=f"Reconciliation_Report_{date.today().isoformat()}.docx",
        use_container_width=True,
    )
else:
    st.info("Upload your billing file and rate table on the left, then click **Run reconciliation**.")
