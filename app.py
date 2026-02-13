import datetime as dt
import streamlit as st

from engine import transform_ap_ledger
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="AP Ledger Transformer", layout="centered")
st.title("AP Ledger Transformer")
st.caption("Upload AP ledger → Run → Download aging report")

uploaded = st.file_uploader("Upload AP ledger Excel (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.stop()

excel_bytes = uploaded.getvalue()

# Let user choose the Top Sheet (because people rename tabs)
try:
    wb = load_workbook(BytesIO(excel_bytes), data_only=True)
    sheetnames = wb.sheetnames
except Exception:
    st.error("This file cannot be opened as an Excel workbook. Upload a valid .xlsx.")
    st.stop()

default_idx = sheetnames.index("Top Sheet") if "Top Sheet" in sheetnames else 0
top_sheet_name = st.selectbox("Top summary sheet tab", sheetnames, index=default_idx)

as_of = st.date_input("As-of date (aging cutoff)", value=dt.date.today())

run = st.button("Run aging analysis", type="primary")

if not run:
    st.stop()

with st.spinner("Processing..."):
    try:
        out_bytes = transform_ap_ledger(
            excel_bytes=excel_bytes,
            as_of=as_of,
            top_sheet_name=top_sheet_name,
        )
    except Exception as e:
        st.error("Processing failed.")
        st.code(str(e))
        st.stop()

st.success("Done.")
st.download_button(
    "Download aging report (.xlsx)",
    data=out_bytes,
    file_name=f"AP_Aging_Report_{as_of.isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
