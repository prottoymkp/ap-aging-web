import streamlit as st

st.set_page_config(page_title="AP Ledger Transformer", layout="centered")
st.title("AP Ledger Transformer")
st.caption("Phase 1: upload → download (no transformation yet).")

uploaded = st.file_uploader("Upload AP ledger Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    st.success("File received.")
    st.download_button(
        label="Download the same file (Phase 1 test)",
        data=uploaded.getvalue(),
        file_name=f"uploaded_{uploaded.name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload a file to enable download.")
