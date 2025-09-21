import os
import pandas as pd
import argparse
import logging
from tqdm import tqdm

# Setup logging
logging.basicConfig(
    filename="merge_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def merge_excels(folder_path, output_file, all_sheets=False, drop_duplicates=False):
    # List only Excel files
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("⚠️ No Excel files found to merge.")
        logging.warning("No Excel files found in the folder.")
        return

    all_data = []

    for file in tqdm(excel_files, desc="Merging files"):
        file_path = os.path.join(folder_path, file)
        try:
            if all_sheets:
                # Read all sheets
                xls = pd.ExcelFile(file_path)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    df["__source_file"] = file
                    df["__sheet_name"] = sheet
                    all_data.append(df)
            else:
                # Only first sheet
                df = pd.read_excel(file_path)
                df["__source_file"] = file
                all_data.append(df)

            logging.info(f"Loaded {file} successfully")

        except Exception as e:
            print(f"❌ Skipped {file}: {e}")
            logging.error(f"Failed to read {file}: {e}")

    if not all_data:
        print("⚠️ No valid Excel data to merge.")
        logging.warning("No valid data found to merge.")
        return

    merged = pd.concat(all_data, ignore_index=True)

    if drop_duplicates:
        before = len(merged)
        merged.drop_duplicates(inplace=True)
        after = len(merged)
        print(f"ℹ️ Removed {before - after} duplicate rows")
        logging.info(f"Removed {before - after} duplicates")

    # Save in different formats
    output_path = os.path.join(folder_path, output_file)
    try:
        if output_file.endswith(".csv"):
            merged.to_csv(output_path, index=False)
        elif output_file.endswith(".json"):
            merged.to_json(output_path, orient="records")
        else:  # Default Excel
            merged.to_excel(output_path, index=False)

        print(f"✅ Merge completed. Saved as {output_path}")
        logging.info(f"Merge completed. Saved as {output_path}")

    except Exception as e:
        print(f"❌ Failed to save file: {e}")
        logging.error(f"Failed to save output: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Merge multiple Excel files into one")

    parser.add_argument("--folder", required=True, help="Folder path containing Excel files")
    parser.add_argument("--output", default="merged_data.xlsx",
                        help="Output file name (supports .xlsx, .csv, .json)")
    parser.add_argument("--all-sheets", action="store_true",
                        help="Merge all sheets from each Excel file")
    parser.add_argument("--drop-duplicates", action="store_true",
                        help="Remove duplicate rows in merged file")

    args = parser.parse_args()

    if os.path.exists(args.folder):
        merge_excels(args.folder, args.output, args.all_sheets, args.drop_duplicates)
    else:
        print("❌ Invalid folder path. Please check and try again.")
        logging.error(f"Invalid folder path: {args.folder}")


