
import os
import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

# Replace these with your own local input/output paths when running.
INPUT_FOLDER = "path/to/input/folder"
OUTPUT_FILE = "path/to/output/merged_stock_data.xlsx"


# Required fields used to identify the correct header row.
REQUIRED_HEADERS = [
    "S/No",
    "Name of Facility",
    "Facility ID",
    "LGA",
    "Ward",
    "Product Code",
    "Expiry Date",
    "Beginning Balance",
    "Quantity Received",
    "Quantity Dispensed",
    "Losses",
    "Adjustment",
    "Physical Stock on Hand",
    "Selling Price",
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_excel_engine(file_path):
    """Select the appropriate Excel engine based on file extension."""
    extension = os.path.splitext(file_path)[1].lower()

    if extension in [".xlsx", ".xlsm"]:
        return "openpyxl"

    if extension == ".xls":
        return "xlrd"

    return None


def find_header_row(file_path, required_headers):
    """
    Search all rows in an Excel file and identify the row
    containing the required column headings.
    """

    engine = get_excel_engine(file_path)

    if engine is None:
        return None

    try:
        preview = pd.read_excel(
            file_path,
            header=None,
            engine=engine
        )

        for row_number in range(len(preview)):
            row_values = (
                preview.iloc[row_number]
                .astype(str)
                .str.strip()
                .str.lower()
                .tolist()
            )

            matches = 0

            for header in required_headers:
                header_lower = header.lower()

                if any(header_lower in value for value in row_values):
                    matches += 1

            # Require most of the expected headers to be present.
            if matches >= max(5, len(required_headers) * 0.5):
                return row_number

    except Exception as error:
        print(f"Could not inspect {file_path}: {error}")

    return None


# ============================================================
# 1. INVENTORY EXCEL FILES
# ============================================================

excel_files = []

for filename in os.listdir(INPUT_FOLDER):

    if filename.lower().endswith((".xlsx", ".xls", ".xlsm")):
        excel_files.append(
            os.path.join(INPUT_FOLDER, filename)
        )


print(f"Excel files found: {len(excel_files)}")


# ============================================================
# 2. IDENTIFY HEADER ROWS
# ============================================================

files_with_headers = []
files_without_headers = []

for file_path in excel_files:

    header_row = find_header_row(
        file_path,
        REQUIRED_HEADERS
    )

    if header_row is not None:

        files_with_headers.append(
            (file_path, header_row)
        )

    else:

        files_without_headers.append(
            file_path
        )


print(f"Files with valid headers: {len(files_with_headers)}")
print(f"Files without valid headers: {len(files_without_headers)}")


# ============================================================
# 3. READ AND MERGE VALID FILES
# ============================================================

dataframes = []

for file_path, header_row in files_with_headers:

    engine = get_excel_engine(file_path)

    try:

        dataframe = pd.read_excel(
            file_path,
            header=header_row,
            engine=engine
        )

        # Remove completely empty rows.
        dataframe = dataframe.dropna(
            how="all"
        )

        # Keep source information for traceability.
        dataframe["Source File"] = os.path.basename(
            file_path
        )

        dataframes.append(dataframe)

        print(
            f"Loaded: {os.path.basename(file_path)} "
            f"({len(dataframe)} rows)"
        )

    except Exception as error:

        print(
            f"Error reading "
            f"{os.path.basename(file_path)}: {error}"
        )


# ============================================================
# 4. COMBINE DATA
# ============================================================

if not dataframes:

    raise ValueError(
        "No valid Excel files were found."
    )


merged_data = pd.concat(
    dataframes,
    ignore_index=True
)


print(
    f"Total rows after merging: "
    f"{len(merged_data)}"
)


# ============================================================
# 5. CLEAN STOCK RECORDS
# ============================================================

# Remove records without a Product Code.
if "Product Code" in merged_data.columns:

    merged_data = merged_data[
        merged_data["Product Code"].notna()
    ]


# Remove records without Quantity Dispensed.
if "Quantity Dispensed" in merged_data.columns:

    merged_data = merged_data[
        merged_data["Quantity Dispensed"].notna()
    ]


# ============================================================
# 6. REMOVE ZERO-STOCK RECORDS
# ============================================================

stock_columns = [
    "Beginning Balance",
    "Quantity Received",
    "Quantity Dispensed",
    "Losses",
    "Adjustment",
    "Physical Stock on Hand",
]


available_stock_columns = [
    column
    for column in stock_columns
    if column in merged_data.columns
]


if available_stock_columns:

    numeric_stock = merged_data[
        available_stock_columns
    ].apply(
        pd.to_numeric,
        errors="coerce"
    ).fillna(0)

    # Remove rows where all key stock fields are zero.
    merged_data = merged_data[
        numeric_stock.sum(axis=1) != 0
    ]


# ============================================================
# 7. REMOVE RECORDS WITH NO DISPENSING OR STOCK
# ============================================================

if (
    "Quantity Dispensed" in merged_data.columns
    and
    "Physical Stock on Hand" in merged_data.columns
):

    quantity_dispensed = pd.to_numeric(
        merged_data["Quantity Dispensed"],
        errors="coerce"
    ).fillna(0)

    physical_stock = pd.to_numeric(
        merged_data["Physical Stock on Hand"],
        errors="coerce"
    ).fillna(0)

    merged_data = merged_data[
        (quantity_dispensed != 0)
        |
        (physical_stock != 0)
    ]


# ============================================================
# 8. SAVE CLEANED DATA
# ============================================================

os.makedirs(
    os.path.dirname(OUTPUT_FILE),
    exist_ok=True
)

merged_data.to_excel(
    OUTPUT_FILE,
    index=False
)


# ============================================================
# 9. SUMMARY
# ============================================================

print("\nProcessing completed.")
print(
    f"Files processed: "
    f"{len(files_with_headers)}"
)
print(
    f"Files skipped: "
    f"{len(files_without_headers)}"
)
print(
    f"Final records: "
    f"{len(merged_data)}"
)
print(
    f"Output saved to: "
    f"{OUTPUT_FILE}"
)
