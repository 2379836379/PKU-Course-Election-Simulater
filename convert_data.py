#!/usr/bin/env python3
from __future__ import annotations

"""Excel to Parquet converter for course data."""

import argparse
from pathlib import Path

import pandas as pd


TABLE_TYPE_COL = "\u8868\u683c\u7c7b\u578b"
UNDERGRAD_TABLE_TYPE = "\u672c\u79d1\u751f\u8bfe\u8868"
PRIMARY_DEDUP_KEYS = ["\u8bfe\u7a0b\u53f7", "\u73ed\u53f7"]
SECONDARY_DEDUP_KEYS = [
    "\u8bfe\u7a0b\u540d",
    "\u6388\u8bfe\u6559\u5e08",
    "\u9662\u7cfb",
    "\u53c2\u8003\u5b66\u5206",
    "\u4e0a\u8bfe\u65f6\u95f4",
]
AUDIENCE_COL = "\u4fee\u8bfb\u5bf9\u8c61"
REQUIRED_COLS = set([TABLE_TYPE_COL] + PRIMARY_DEDUP_KEYS + SECONDARY_DEDUP_KEYS + [AUDIENCE_COL])


def merge_course_group(group: pd.DataFrame) -> pd.Series:
    row = group.iloc[0].copy()
    audiences = []
    for value in group[AUDIENCE_COL].astype(str):
        value = value.strip()
        if not value or value.lower() == "nan":
            continue
        audiences.append(value)
    row[AUDIENCE_COL] = "\uff0c".join(dict.fromkeys(audiences))
    return row


def convert_excel_to_parquet(excel_file: Path, parquet_file: Path) -> bool:
    print(f"Reading {excel_file}...")

    if not excel_file.exists():
        print(f"Error: file not found: {excel_file}")
        print("Hint: by default the script looks for test.xlsx in the current working directory.")
        return False

    try:
        df = pd.read_excel(excel_file)
        print(f"Loaded {len(df)} source rows")

        missing_cols = REQUIRED_COLS - {str(c) for c in df.columns}
        if missing_cols:
            print(f"Error: Excel is missing required columns: {', '.join(sorted(missing_cols))}")
            print(f"Current columns: {list(df.columns)}")
            return False

        filtered_df = df[df[TABLE_TYPE_COL].astype(str).str.strip() == UNDERGRAD_TABLE_TYPE].copy()
        print(f"Kept {len(filtered_df)} rows where {TABLE_TYPE_COL} == {UNDERGRAD_TABLE_TYPE}")

        print("Processing rows...")
        first_pass = (
            filtered_df.groupby(PRIMARY_DEDUP_KEYS, sort=False, dropna=False, group_keys=False)
            .apply(merge_course_group, include_groups=False)
            .reset_index()
        )
        result_df = (
            first_pass.groupby(SECONDARY_DEDUP_KEYS, sort=False, dropna=False, group_keys=False)
            .apply(merge_course_group, include_groups=False)
            .reset_index()
        )
        print(f"Dedup complete: {len(result_df)} rows")

        parquet_file.parent.mkdir(parents=True, exist_ok=True)
        print(f"Writing {parquet_file}...")
        result_df.to_parquet(parquet_file, compression="snappy", index=False)

        excel_size = excel_file.stat().st_size / (1024 * 1024)
        parquet_size = parquet_file.stat().st_size / (1024 * 1024)

        print("\nConversion succeeded!")
        print(f"  Excel size:   {excel_size:.2f} MB")
        print(f"  Parquet size: {parquet_size:.2f} MB")
        if excel_size > 0:
            saved = (excel_size - parquet_size) / excel_size * 100
            print(f"  Space saved:  {saved:.1f}%")
        return True

    except Exception as exc:
        print(f"Error: conversion failed - {exc}")
        return False


def parse_args() -> argparse.Namespace:
    base_dir = Path.cwd()

    parser = argparse.ArgumentParser(
        description="CourseElection data converter (Excel -> Parquet)"
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=base_dir / "test.xlsx",
        help="Input Excel path (default: ./test.xlsx)",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=base_dir / "courses.parquet",
        help="Output Parquet path (default: ./courses.parquet)",
    )
    return parser.parse_args()


if __name__ == "__main__":
    print("=" * 60)
    print("CourseElection data converter")
    print("=" * 60)
    print()

    args = parse_args()
    success = convert_excel_to_parquet(args.excel, args.out)

    if success:
        print("\nNext app launch will use the regenerated Parquet file.")
    else:
        print("\nConversion failed.")
