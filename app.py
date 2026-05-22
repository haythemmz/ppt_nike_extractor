import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import extract_products_from_pptx


st.set_page_config(
    page_title="Nike PPT Extractor",
    layout="wide",
)

st.title("Nike PowerPoint Extractor")

st.write(
    "Upload a Nike PowerPoint file, preview the extracted Excel sheets, then download the Excel output."
)

uploaded_file = st.file_uploader(
    "Upload PowerPoint file",
    type=["pptx"],
)

if uploaded_file is not None:
    st.success(f"Selected file: {uploaded_file.name}")

    if st.button("Extract and preview Excel"):
        with st.spinner("Processing PowerPoint... Please wait."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                pptx_path = tmpdir / uploaded_file.name

                with open(pptx_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                products, style_colors, raw_text = extract_products_from_pptx(pptx_path)

                excel_buffer = io.BytesIO()

                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    products.to_excel(writer, sheet_name="Products", index=False)
                    style_colors.to_excel(writer, sheet_name="Style_Colors", index=False)
                    raw_text.to_excel(writer, sheet_name="Raw_Slide_Text", index=False)

                excel_buffer.seek(0)

                st.session_state["products"] = products
                st.session_state["style_colors"] = style_colors
                st.session_state["raw_text"] = raw_text
                st.session_state["excel_bytes"] = excel_buffer.getvalue()
                st.session_state["output_file_name"] = (
                    Path(uploaded_file.name).stem + "_extracted.xlsx"
                )

                st.success("Extraction completed successfully.")

if "products" in st.session_state:
    products = st.session_state["products"]
    style_colors = st.session_state["style_colors"]
    raw_text = st.session_state["raw_text"]

    st.subheader("Extraction summary")

    col1, col2, col3 = st.columns(3)

    col1.metric("Products", len(products))
    col2.metric("Style-color rows", len(style_colors))
    col3.metric("Raw slide rows", len(raw_text))

    st.subheader("Excel preview")

    tab1, tab2, tab3 = st.tabs(
        [
            "Products",
            "Style Colors",
            "Raw Slide Text",
        ]
    )

    with tab1:
        st.dataframe(
            products,
            use_container_width=True,
            hide_index=True,
        )

    with tab2:
        st.dataframe(
            style_colors,
            use_container_width=True,
            hide_index=True,
        )

    with tab3:
        st.dataframe(
            raw_text,
            use_container_width=True,
            hide_index=True,
        )

    st.download_button(
        label="Download Excel file",
        data=st.session_state["excel_bytes"],
        file_name=st.session_state["output_file_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )