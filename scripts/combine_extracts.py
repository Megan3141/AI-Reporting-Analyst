from pathlib import Path

import pandas as pd


FOLDER = Path("path/to/extract_outputs")

SOURCE_FILES = [
    "season_extract.xlsx",
    "recovery_extract.xlsx",
]

OUTPUT_FILE = FOLDER / "master_extract.xlsx"

CATEGORY_SHEETS = [
    "Category_A",
    "Category_B",
    "Category_C",
    "Category_D",
]

combined_categories = {}
combined_logs = []


for category in CATEGORY_SHEETS:
    frames = []

    for source_file in SOURCE_FILES:
        file_path = FOLDER / source_file

        if not file_path.exists():
            print("Missing file:", source_file)
            continue

        try:
            df = pd.read_excel(file_path, sheet_name=category)

            df = df[df["Metric 1"] != 0]
            df = df[df["Metric 1"].notna()]

            df["Source Extract File"] = source_file
            frames.append(df)

            print(f"Loaded {category} from {source_file}: {len(df)} rows")

        except ValueError:
            print(f"No {category} sheet in {source_file}")

    if frames:
        combined_categories[category] = pd.concat(frames, ignore_index=True)


for source_file in SOURCE_FILES:
    file_path = FOLDER / source_file

    if not file_path.exists():
        continue

    try:
        log_df = pd.read_excel(file_path, sheet_name="Extraction_Log")
        log_df["Source Extract File"] = source_file
        combined_logs.append(log_df)

    except ValueError:
        print(f"No Extraction_Log in {source_file}")


with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for category, df in combined_categories.items():
        df.to_excel(
            writer,
            sheet_name=category[:31],
            index=False,
        )

    if combined_logs:
        final_log = pd.concat(combined_logs, ignore_index=True)

        final_log.to_excel(
            writer,
            sheet_name="Extraction_Log",
            index=False,
        )


print()
print("Master file created:")
print(OUTPUT_FILE)