from pathlib import Path
import argparse
import pandas as pd

from bundle_logic import process_workbook


def main():
    parser = argparse.ArgumentParser(description="Generate bundle plan from PUF panel workbook")
    parser.add_argument("input_file", help="Path to input Excel file")
    parser.add_argument("-o", "--output-dir", default="output", help="Output folder")
    args = parser.parse_args()

    input_file = Path(args.input_file)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    orders_df, layer_df, bundles_df, summary_df = process_workbook(input_file)

    output_file = output_dir / f"{input_file.stem}_bundle_output.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        orders_df.to_excel(writer, index=False, sheet_name="ORDER_DATA_CLEAN")
        layer_df.to_excel(writer, index=False, sheet_name="LAYER_DATA_CLEAN")
        bundles_df.to_excel(writer, index=False, sheet_name="BUNDLE_PLAN")
        summary_df.to_excel(writer, index=False, sheet_name="SUMMARY")

    print(f"Created: {output_file}")
    print(f"Bundle rows: {len(bundles_df)}")
    print(f"Orders summarized: {len(summary_df)}")


if __name__ == "__main__":
    main()