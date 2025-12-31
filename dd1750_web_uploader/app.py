import os, tempfile
import streamlit as st

from dd1750_core import (
    load_cfg, generate_dd1750_from_pdf, generate_dd1750_from_excel, merge_dd1750_pdfs
)

st.set_page_config(page_title="BOM ➜ DD1750 Generator", layout="wide")
st.title("BOM ➜ DD Form 1750 Generator")
st.caption("Upload a BOM (PDF/Excel) + a flat DD1750 template PDF and download a multi-page DD1750 + audit file.")

tab1, tab2 = st.tabs(["Generate DD1750", "Merge existing DD1750 PDFs"])

with tab1:
    st.subheader("Upload")
    col1, col2 = st.columns(2)
    with col1:
        bom_file = st.file_uploader("BOM (PDF or Excel)", type=["pdf","xlsx","xls"])
    with col2:
        template_file = st.file_uploader("DD1750 template (flat PDF, e.g., blank_flat.pdf)", type=["pdf"])

    st.subheader("Options")
    left, right = st.columns(2)

    with left:
        label = st.selectbox("Label under description", options=["NSN","SN"], index=0)
        page_start = st.number_input("Start parsing PDF at page (0-based)", min_value=0, value=0, step=1)
        force_ocr = st.checkbox("Force OCR (for scanned PDFs)", value=False)
        ocr_dpi = st.slider("OCR DPI", min_value=150, max_value=400, value=250, step=10)

    with right:
        st.markdown("**Excel mapping (only used if BOM is Excel):**")
        sheet = st.text_input("Sheet name (blank = active sheet)", value="")
        col_desc = st.text_input("Description header", value="Description")
        col_mat = st.text_input("Material/NSN header", value="Material")
        col_qty = st.text_input("Qty header", value="OH QTY")

    if st.button("Generate DD1750"):
        if not bom_file or not template_file:
            st.error("Upload BOTH the BOM and the template PDF.")
        else:
            with st.spinner("Generating..."):
                with tempfile.TemporaryDirectory() as tmp:
                    cfg = load_cfg(os.path.join(os.path.dirname(__file__), "config.yaml"))

                    bom_path = os.path.join(tmp, bom_file.name)
                    tpl_path = os.path.join(tmp, template_file.name)

                    with open(bom_path, "wb") as f:
                        f.write(bom_file.getbuffer())
                    with open(tpl_path, "wb") as f:
                        f.write(template_file.getbuffer())

                    out_pdf = os.path.join(tmp, "DD1750_OUTPUT.pdf")
                    out_audit = os.path.join(tmp, "DD1750_OUTPUT_AUDIT.csv")

                    try:
                        if bom_file.name.lower().endswith((".xlsx",".xls")):
                            items = generate_dd1750_from_excel(
                                bom_path, tpl_path, cfg, out_pdf, out_audit,
                                sheet=sheet or None, col_desc=col_desc, col_mat=col_mat, col_qty=col_qty,
                                label=label
                            )
                        else:
                            items = generate_dd1750_from_pdf(
                                bom_path, tpl_path, cfg, out_pdf, out_audit,
                                force_ocr=force_ocr, ocr_dpi=ocr_dpi, page_start=page_start,
                                label=label
                            )
                    except Exception as e:
                        st.exception(e)
                        st.stop()

                    st.success(f"Done. Line items generated: {len(items)}")

                    with open(out_pdf, "rb") as f:
                        st.download_button("Download DD1750 PDF", data=f, file_name=f"DD1750_{os.path.splitext(bom_file.name)[0]}.pdf")
                    with open(out_audit, "rb") as f:
                        st.download_button("Download AUDIT CSV", data=f, file_name=f"DD1750_{os.path.splitext(bom_file.name)[0]}_AUDIT.csv")

                    st.info("If AUDIT flags SUSPICIOUS_QTY, it’s almost always OCR noise. Re-run with higher DPI or verify those few lines quickly.")

with tab2:
    st.subheader("Merge DD1750 PDFs")
    merge_files = st.file_uploader("Upload DD1750 PDFs to merge", type=["pdf"], accept_multiple_files=True)
    keep_all_pages = st.checkbox("Keep all pages from each input (normally OFF)", value=False)

    if st.button("Merge"):
        if not merge_files:
            st.error("Upload at least one PDF.")
        else:
            with st.spinner("Merging..."):
                with tempfile.TemporaryDirectory() as tmp:
                    paths = []
                    for uf in sorted(merge_files, key=lambda x: x.name):
                        p = os.path.join(tmp, uf.name)
                        with open(p, "wb") as f:
                            f.write(uf.getbuffer())
                        paths.append(p)

                    out_pdf = os.path.join(tmp, "DD1750_MERGED.pdf")
                    try:
                        merge_dd1750_pdfs(paths, out_pdf, keep_all_pages=keep_all_pages)
                    except Exception as e:
                        st.exception(e)
                        st.stop()

                    with open(out_pdf, "rb") as f:
                        st.download_button("Download merged PDF", data=f, file_name="DD1750_MERGED.pdf")
                    st.success(f"Merged {len(paths)} file(s).")
