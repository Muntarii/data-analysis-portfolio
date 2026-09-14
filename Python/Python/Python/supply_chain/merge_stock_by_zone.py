import os
import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

# Replace with the folder containing the zone subfolders.
PARENT_FOLDER = "path/to/parent/folder"

# Name of the zone folders to process.
ZONES = [
    "zone1",
    "zone2",
    "zone3",
]

# Output file.
OUTPUT_FILE = "path/to/output/merged_data.xlsx"


# ============================================================
# MERGE FILES FROM EACH ZONE
# ============================================================

all_data = []

for zone in ZONES:

    zone_folder = os.path.join(
        PARENT_FOLDER,
        zone
    )

    if not os.path.exists(zone_folder):

        print(
            f"Folder not found: {zone}"
        )

        continue

    for filename in os.listdir(zone_folder):

        if not filename.lower().endswith(".xlsx"):
            continue

        file_path = os.path.join(
            zone_folder,
            filename
        )

        try:

            dataframe = pd.read_excel(
                file_path
            )

            # Preserve the origin of each record.
            dataframe["Source Zone"] = zone
            dataframe["Source File"] = filename

            all_data.append(
                dataframe
            )

            print(
                f"Loaded: {zone}/{filename}"
            )

        except Exception as error:

            print(
                f"Error reading {filename}: "
                f"{error}"
            )


# ============================================================
# COMBINE ALL DATA
# ============================================================

if not all_data:

    raise ValueError(
        "No Excel files were successfully loaded."
    )


merged_data = pd.concat(
    all_data,
    ignore_index=True
)


# ============================================================
# SAVE OUTPUT
# ============================================================

output_folder = os.path.dirname(
    OUTPUT_FILE
)

if output_folder:

    os.makedirs(
        output_folder,
        exist_ok=True
    )


merged_data.to_excel(
    OUTPUT_FILE,
    index=False
)


# ============================================================
# SUMMARY
# ============================================================

print("\nProcessing completed.")
print(
    f"Total records: {len(merged_data)}"
)
print(
    f"Output: {OUTPUT_FILE}"
)
