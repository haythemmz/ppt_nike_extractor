#!/usr/bin/env python
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from extractor import extract_products_from_pptx


def read_expected(path: Path) -> pd.DataFrame:
    expected = pd.read_csv(path)
    if "style_color_key" not in expected.columns:
        raise ValueError("Expected CSV must contain a style_color_key column.")
    expected = expected.copy()
    expected["style_color_key"] = expected["style_color_key"].astype(str).str.strip().str.upper()
    if "slide_number" in expected.columns:
        expected["slide_number"] = pd.to_numeric(expected["slide_number"], errors="coerce").astype("Int64")
    return expected


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare expected style-color keys with PPT extractor output.")
    parser.add_argument("pptx", type=Path, help="PowerPoint file to extract.")
    parser.add_argument("expected_csv", type=Path, help="CSV with at least style_color_key, optionally slide_number.")
    parser.add_argument("--out", type=Path, default=Path("style_color_validation.xlsx"), help="Output report workbook.")
    args = parser.parse_args()

    _, extracted, _, qa_issues = extract_products_from_pptx(args.pptx)
    expected = read_expected(args.expected_csv)

    actual = extracted.copy()
    if "style_color_key" not in actual.columns:
        actual["style_color_key"] = pd.Series(dtype="str")
    actual["style_color_key"] = actual["style_color_key"].astype(str).str.strip().str.upper()
    if "slide_number" in expected.columns and expected["slide_number"].notna().any():
        merge_keys = ["slide_number", "style_color_key"]
        if "slide_number" not in actual.columns:
            actual["slide_number"] = pd.Series(dtype="Int64")
        actual["slide_number"] = pd.to_numeric(actual["slide_number"], errors="coerce").astype("Int64")
    else:
        merge_keys = ["style_color_key"]

    comparison = expected.merge(
        actual,
        on=merge_keys,
        how="outer",
        suffixes=("_expected", "_actual"),
        indicator=True,
    )

    missing = comparison[comparison["_merge"] == "left_only"].copy()
    extra = comparison[comparison["_merge"] == "right_only"].copy()
    matched = comparison[comparison["_merge"] == "both"].copy()

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        matched.to_excel(writer, sheet_name="Matched", index=False)
        missing.to_excel(writer, sheet_name="Missing", index=False)
        extra.to_excel(writer, sheet_name="Extra", index=False)
        qa_issues.to_excel(writer, sheet_name="QA_Issues", index=False)

    print(f"Matched: {len(matched)}")
    print(f"Missing: {len(missing)}")
    print(f"Extra: {len(extra)}")
    print(f"QA warnings: {len(qa_issues)}")
    print(f"Report written to: {args.out}")


if __name__ == "__main__":
    main()
