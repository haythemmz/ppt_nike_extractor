import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import extract_products_from_pptx


st.set_page_config(
    page_title="Nike PPT Extractor",
    layout="centered",
)

st.title("Nike PowerPoint Extractor")
st.write("Upload a Nike PowerPoint file and generate the extracted Excel output.")

uploaded_file = st.file_uploader(
    "Upload PowerPoint file",
    type=["pptx"],
)

if uploaded_file is not None:
    st.success(f"Selected file: {uploaded_file.name}")

    if st.button("Extract to Excel"):
        with st.spinner("Processing PowerPoint..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                pptx_path = tmpdir / uploaded_file.name
                output_path = tmpdir / uploaded_file.name.replace(".pptx", "_extracted.xlsx")

                with open(pptx_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                products, style_colors, raw_text = extract_products_from_pptx(pptx_path)

                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    products.to_excel(writer, sheet_name="Products", index=False)
                    style_colors.to_excel(writer, sheet_name="Style_Colors", index=False)
                    raw_text.to_excel(writer, sheet_name="Raw_Slide_Text", index=False)

                st.success("Extraction completed.")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="Download Excel file",
                        data=f,
                        file_name=output_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                st.write(f"Products: {len(products)}")
                st.write(f"Style-color rows: {len(style_colors)}")