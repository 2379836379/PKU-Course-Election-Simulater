#!/usr/bin/env python3
"""数据格式转换工具

将 Excel 课表数据转换为 Parquet（默认 snappy 压缩），以提升读取速度。

用法示例（Windows / PowerShell）：
  python ./PKU-Course-Election/convert_data.py
  py -3 ./PKU-Course-Election/convert_data.py
  python ./PKU-Course-Election/convert_data.py --excel ./test.xlsx --out ./courses.parquet

说明：脚本默认读取“当前工作目录”下的 test.xlsx，并在当前工作目录输出 courses.parquet；也可以用参数指定路径。
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


def convert_excel_to_parquet(excel_file: Path, parquet_file: Path) -> bool:
    print(f"正在读取 {excel_file}...")

    if not excel_file.exists():
        print(f"错误: 找不到文件 {excel_file}")
        print("提示: 默认按当前工作目录查找 test.xlsx；也可以用 --excel 指定路径。")
        return False

    try:
        df = pd.read_excel(excel_file)
        print(f"成功读取 {len(df)} 条原始记录")

        required_cols = {"课程号", "班号", "修读对象"}
        missing_cols = required_cols - {str(c) for c in df.columns}
        if missing_cols:
            print(f"错误: Excel 缺少必要列: {', '.join(sorted(missing_cols))}")
            print(f"当前表头: {list(df.columns)}")
            return False

        print("正在处理数据...")
        grouped = df.groupby(["课程号", "班号"], as_index=False)

        processed_rows = []
        for _, group in grouped:
            row = group.iloc[0].copy()
            if len(group) > 1:
                row["修读对象"] = "，".join(group["修读对象"].astype(str).unique())
            processed_rows.append(row)

        result_df = pd.DataFrame(processed_rows)
        print(f"处理完成，共 {len(result_df)} 条课程记录")

        parquet_file.parent.mkdir(parents=True, exist_ok=True)
        print(f"正在保存为 {parquet_file}...")
        result_df.to_parquet(parquet_file, compression="snappy", index=False)

        excel_size = excel_file.stat().st_size / (1024 * 1024)
        parquet_size = parquet_file.stat().st_size / (1024 * 1024)

        print("\n转换成功!")
        print(f"  Excel 文件大小:   {excel_size:.2f} MB")
        print(f"  Parquet 文件大小: {parquet_size:.2f} MB")
        if excel_size > 0:
            print(f"  空间节省:        {((excel_size - parquet_size) / excel_size * 100):.1f}%")
        print("\n预计读取速度提升: 5-10 倍")

        return True

    except Exception as e:
        print(f"错误: 转换失败 - {e}")
        return False


def parse_args() -> argparse.Namespace:
    base_dir = Path.cwd()

    parser = argparse.ArgumentParser(
        description="CourseElection 数据格式转换工具（Excel -> Parquet）"
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=base_dir / "test.xlsx",
        help="输入 Excel 路径（默认：当前工作目录下的 test.xlsx）",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=base_dir / "courses.parquet",
        help="输出 Parquet 路径（默认：当前工作目录下的 courses.parquet）",
    )
    return parser.parse_args()


if __name__ == "__main__":
    print("=" * 60)
    print("CourseElection 数据格式转换工具")
    print("=" * 60)
    print()

    args = parse_args()
    success = convert_excel_to_parquet(args.excel, args.out)

    if success:
        print("\n下次启动应用时将自动使用 Parquet 格式，加载速度更快！")
    else:
        print("\n转换失败，请检查错误信息")

