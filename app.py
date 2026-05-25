import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import extract_products_from_pptx


def find_column(columns: list[str], candidates: list[str]) -> str | None:
    normalized = {str(column).strip().lower(): column for column in columns}
    for candidate in candidates:
        match = normalized.get(candidate.lower())
        if match is not None:
            return match
    return None


def normalize_style_color(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def build_nike_lookup(nike_data: pd.DataFrame) -> tuple[pd.DataFrame | None, str | None]:
    style_col = find_column(nike_data.columns.tolist(), ["Style-Color", "Style Color", "style_color_key", "style_color"])
    wholesale_col = find_column(nike_data.columns.tolist(), ["Wholesale", "Wholesale Price", "wholesale_price"])
    retail_col = find_column(nike_data.columns.tolist(), ["Retail", "Retail Price", "retail_price"])

    missing = [
        label
        for label, column in (
            ("Style-Color", style_col),
            ("Wholesale", wholesale_col),
            ("Retail", retail_col),
        )
        if column is None
    ]
    if missing:
        return None, "Missing required Nike CSV columns: " + ", ".join(missing)

    nike = nike_data[[style_col, wholesale_col, retail_col]].copy()
    nike.columns = ["style_color_key", "nike_wholesale_price", "nike_retail_price"]
    nike["style_color_key"] = nike["style_color_key"].map(normalize_style_color)
    nike["nike_wholesale_price"] = pd.to_numeric(nike["nike_wholesale_price"], errors="coerce")
    nike["nike_retail_price"] = pd.to_numeric(nike["nike_retail_price"], errors="coerce")
    nike = nike[nike["style_color_key"] != ""]

    nike_lookup = (
        nike.groupby("style_color_key", as_index=False)
        .agg(
            nike_wholesale_price=("nike_wholesale_price", "first"),
            nike_retail_price=("nike_retail_price", "first"),
            nike_row_count=("style_color_key", "size"),
            nike_wholesale_price_options=("nike_wholesale_price", lambda x: ",".join(map(str, sorted(x.dropna().unique())))),
            nike_retail_price_options=("nike_retail_price", lambda x: ",".join(map(str, sorted(x.dropna().unique())))),
        )
    )
    return nike_lookup, None


def prices_match(left, right) -> bool:
    if pd.isna(left) or pd.isna(right):
        return False
    return round(float(left), 2) == round(float(right), 2)


def compare_ppt_to_nike(style_colors: pd.DataFrame, nike_data: pd.DataFrame) -> tuple[pd.DataFrame | None, str | None]:
    nike_lookup, error = build_nike_lookup(nike_data)
    if error:
        return None, error

    ppt = style_colors.copy()
    ppt["style_color_key"] = ppt["style_color_key"].map(normalize_style_color)
    ppt["wholesale_price"] = pd.to_numeric(ppt["wholesale_price"], errors="coerce")
    ppt["retail_price"] = pd.to_numeric(ppt["retail_price"], errors="coerce")

    comparison = ppt.merge(nike_lookup, on="style_color_key", how="left")
    comparison["exists_in_nike"] = comparison["nike_row_count"].notna()
    comparison["wholesale_price_match"] = comparison.apply(
        lambda row: prices_match(row["wholesale_price"], row["nike_wholesale_price"]),
        axis=1,
    )
    comparison["retail_price_match"] = comparison.apply(
        lambda row: prices_match(row["retail_price"], row["nike_retail_price"]),
        axis=1,
    )
    comparison["all_checks_pass"] = (
        comparison["exists_in_nike"]
        & comparison["wholesale_price_match"]
        & comparison["retail_price_match"]
    )

    preferred_columns = [
        "slide_number",
        "category",
        "product_name",
        "style_code",
        "color_code",
        "style_color_key",
        "exists_in_nike",
        "wholesale_price",
        "nike_wholesale_price",
        "wholesale_price_match",
        "retail_price",
        "nike_retail_price",
        "retail_price_match",
        "all_checks_pass",
        "nike_row_count",
        "nike_wholesale_price_options",
        "nike_retail_price_options",
        "status",
    ]
    existing_columns = [column for column in preferred_columns if column in comparison.columns]
    return comparison[existing_columns], None


def filter_dataframe(df: pd.DataFrame, key_prefix: str) -> pd.DataFrame:
    if df.empty:
        return df

    filtered = df.copy()

    filter_columns = [
        ("slide_number", "Slide"),
        ("category", "Category"),
        ("product_name", "Product name"),
        ("style_code", "Style code"),
        ("style_codes", "Style code"),
        ("style_color_key", "Style-color code"),
    ]

    available_filters = [
        (column, label)
        for column, label in filter_columns
        if column in filtered.columns
    ]

    if available_filters:
        columns = st.columns(len(available_filters))
    else:
        columns = []

    for filter_column, (column, label) in zip(columns, available_filters):
        values = sorted(filtered[column].dropna().astype(str).unique().tolist())
        selected_value = filter_column.selectbox(
            label,
            options=["All"] + values,
            key=f"{key_prefix}_{column}_filter",
        )
        if selected_value != "All":
            filtered = filtered[filtered[column].astype(str) == selected_value]

    st.caption(f"Showing {len(filtered):,} of {len(df):,} rows")
    return filtered


st.set_page_config(
    page_title="Nike PPT Extractor",
    layout="wide",
)

st.title("Nike PowerPoint Extractor")

st.write(
    "Upload a Nike PowerPoint file with products or accessories (bags/backpacks), preview the extracted Excel sheets, then download the Excel output. "
    "Supports both regular apparel products and Nike accessories with special bag style codes (e.g., N.100.3478.091)."
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
        products_view = filter_dataframe(products, "products")
        st.dataframe(
            products_view,
            use_container_width=True,
            hide_index=True,
        )

    with tab2:
        style_colors_view = filter_dataframe(style_colors, "style_colors")
        st.dataframe(
            style_colors_view,
            use_container_width=True,
            hide_index=True,
        )

    with tab3:
        raw_text_view = filter_dataframe(raw_text, "raw_text")
        st.dataframe(
            raw_text_view,
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("Compare with Nike CSV")

    nike_csv = st.file_uploader(
        "Upload Nike data CSV",
        type=["csv"],
        key="nike_csv_upload",
    )

    if nike_csv is not None:
        try:
            nike_data = pd.read_csv(nike_csv)
            comparison, comparison_error = compare_ppt_to_nike(style_colors, nike_data)
        except Exception as exc:
            comparison = None
            comparison_error = f"Could not read or compare the Nike CSV: {exc}"

        if comparison_error:
            st.error(comparison_error)
        elif comparison is not None:
            missing_count = int((~comparison["exists_in_nike"]).sum())
            wholesale_mismatch_count = int(
                (comparison["exists_in_nike"] & ~comparison["wholesale_price_match"]).sum()
            )
            retail_mismatch_count = int(
                (comparison["exists_in_nike"] & ~comparison["retail_price_match"]).sum()
            )
            pass_count = int(comparison["all_checks_pass"].sum())

            comp_col1, comp_col2, comp_col3, comp_col4 = st.columns(4)
            comp_col1.metric("Matched rows", pass_count)
            comp_col2.metric("Missing in Nike", missing_count)
            comp_col3.metric("Wholesale mismatches", wholesale_mismatch_count)
            comp_col4.metric("Retail mismatches", retail_mismatch_count)

            show_issues_only = st.checkbox(
                "Show only missing or price mismatches",
                value=True,
            )

            comparison_view = comparison
            if show_issues_only:
                comparison_view = comparison[~comparison["all_checks_pass"]]

            comparison_view = filter_dataframe(comparison_view, "comparison")

            st.dataframe(
                comparison_view,
                use_container_width=True,
                hide_index=True,
            )

            comparison_csv = comparison.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download comparison CSV",
                data=comparison_csv,
                file_name="ppt_vs_nike_comparison.csv",
                mime="text/csv",
            )

    st.download_button(
        label="Download Excel file",
        data=st.session_state["excel_bytes"],
        file_name=st.session_state["output_file_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
