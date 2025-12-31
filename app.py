import os
import tempfile
import streamlit as st

from dd1750_core import (
    load_cfg,
    generate_dd1750_from_pdf,
    generate_dd1750_from_excel,
    merge_dd1750_pdfs,
)

st.set_page_config(page_title="BOM-to-1750", layout="wide")

st.title("BOM-to-1750")
st.caption("Upload a BOM (PDF or Excel) + a blank DD1750 template PDF. Generates a multi-page DD1750 + audit CSV.")

cfg = load_cfg("config.yaml")

tab1, tab2 = st.tabs(["Generate DD1750", "Merge existing DD1750 PDFs"])

with tab1:
    c1, c2 = st.columns(2)

    with c1:
        bom_file = st.file_uploader("Upload BOM (PDF or Excel)", type=["pdf", "xlsx"])
        template_file = st.file_uploader("Upload blank DD1750 template PDF", type=["pdf"])

        st.subheader("Options")
        label = st.selectbox("Label under description", ["NSN", "SN"], index=0)

        page_start = st.number_input("Start parsing PDF at page (0-based)", min_value=0, value=0, step=1)

        force_ocr = st.checkbox("Force OCR (for scanned PDFs)", value=False)
        ocr_dpi = st.slider("OCR DPI", min_value=150, max_value=400, value=250, step=10)

    with c2:
        st.subheader("Excel mapping (only used if BOM is Excel)")
        sheet_name = st.text_input("Sheet name (blank = active sheet)", value="")
        col_desc = st.text_input("Description header", value="Description")
        col_mat = st.text_input("Material/NSN header", value="Material")
        col_qty = st.text_input("Qty header", value="OH QTY")

        st.info("Tip: For B49-style TM 'Component Listing / Hand Receipt' PDFs, keep Force OCR OFF. "
                "OCR can misread gridlines and create insane quantities.")

    if st.button("Generate DD1750", type="primary", disabled=not (bom_file and template_file)):
        with tempfile.TemporaryDirectory() as td:
            bom_path = os.path.join(td, bom_file.name)
            tpl_path = os.path.join(td, template_file.name)
            with open(bom_path, "wb") as f:
                f.write(bom_file.read())
            with open(tpl_path, "wb") as f:
                f.write(template_file.read())

            out_pdf = os.path.join(td, "DD1750_OUTPUT.pdf")
            out_csv = os.path.join(td, "DD1750_AUDIT.csv")

            try:
                if bom_path.lower().endswith(".xlsx"):
                    items = generate_dd1750_from_excel(
                        excel_path=bom_path,
                        template_pdf=tpl_path,
                        cfg=cfg,
                        out_pdf=out_pdf,
                        out_audit_csv=out_csv,
                        sheet=sheet_name.strip() or None,
                        col_desc=col_desc,
                        col_mat=col_mat,
                        col_qty=col_qty,
                        label=label,
                    )
                else:
                    items = generate_dd1750_from_pdf(
                        bom_pdf=bom_path,
                        template_pdf=tpl_path,
                        cfg=cfg,
                        out_pdf=out_pdf,
                        out_audit_csv=out_csv,
                        force_ocr=force_ocr,
                        ocr_dpi=int(ocr_dpi),
                        page_start=int(page_start),
                        label=label,
                    )
            except Exception as e:
                st.error(f"Generation failed: {e}")
                st.exception(e)
            else:
                st.success(f"Generated {len(items)} line items.")

                with open(out_pdf, "rb") as f:
                    st.download_button("Download DD1750 PDF", f, file_name="DD1750_OUTPUT.pdf", mime="application/pdf")

                with open(out_csv, "rb") as f:
                    st.download_button("Download AUDIT CSV", f, file_name="DD1750_AUDIT.csv", mime="text/csv")

with tab2:
    st.subheader("Merge existing DD1750 PDFs into one file")
    pdfs = st.file_uploader("Upload DD1750 PDFs", type=["pdf"], accept_multiple_files=True)
    keep_all = st.checkbox("Keep all pages from each PDF (OFF keeps only page 1 from each file)", value=False)

    if st.button("Merge PDFs", disabled=not pdfs):
        with tempfile.TemporaryDirectory() as td:
            paths = []
            for fobj in pdfs:
                p = os.path.join(td, fobj.name)
                with open(p, "wb") as f:
                    f.write(fobj.read())
                paths.append(p)
            out_pdf = os.path.join(td, "DD1750_MERGED.pdf")
            try:
                merge_dd1750_pdfs(paths, out_pdf, keep_all_pages=keep_all)
            except Exception as e:
                st.error(f"Merge failed: {e}")
                st.exception(e)
            else:
                with open(out_pdf, "rb") as f:
                    st.download_button("Download merged PDF", f, file_name="DD1750_MERGED.pdf", mime="application/pdf")
