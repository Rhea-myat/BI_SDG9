import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
import re

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(__file__).resolve().parent / "Datasets_v2"
INPUT_FILE = BASE_DIR / "ICT Access and Usage by Individuals.xlsx"

OUTPUT_DIR = Path(__file__).resolve().parent / "output_ict_usage"

print("BASE_DIR:", BASE_DIR)
print("INPUT_FILE:", INPUT_FILE)
print("Exists:", INPUT_FILE.exists())

KEEP_AGGREGATES = False
# False = keep only country-level rows
# can change to True, if want to keep rows like OECD, EU, Non-OECD economies etc, but right now we want to focus on country-level data only

# =========================================================
# HELPERS
# =========================================================
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def col_letters_to_index(col_letters: str) -> int:
    result = 0
    for ch in col_letters:
        result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
    return result

def parse_cell_ref(cell_ref: str):
    match = re.match(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        return None, None
    return match.group(1), int(match.group(2))

def load_shared_strings(zip_file):
    shared_strings = []
    if "xl/sharedStrings.xml" not in zip_file.namelist():  #check if sharedStrings.xml exists, if not return empty list
        return shared_strings

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))  # read and parse sharedStrings.xml to get the list of shared strings used in the workbook
    # iterates each shared string items
    for si in root.findall("a:si", NS):
        parts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(t.text or "") # store reconstructed string value for this shared string item, handling both simple and rich text cases
        shared_strings.append("".join(parts))
    return shared_strings

def load_sheet_rows(zip_file, sheet_path="xl/worksheets/sheet1.xml"):
    shared_strings = load_shared_strings(zip_file)
    root = ET.fromstring(zip_file.read(sheet_path))

    rows = []
    for row in root.findall(".//a:sheetData/a:row", NS):
        row_data = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r")
            cell_type = cell.attrib.get("t")
            col_letters, _ = parse_cell_ref(ref)

            if not col_letters:
                continue

            col_index = col_letters_to_index(col_letters)

            value_elem = cell.find("a:v", NS)
            value = value_elem.text if value_elem is not None else None

            if cell_type == "s" and value is not None:
                value = shared_strings[int(value)]
            elif cell_type == "inlineStr":
                is_elem = cell.find("a:is", NS)
                if is_elem is not None:
                    texts = [t.text or "" for t in is_elem.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
                    value = "".join(texts)

            row_data[col_index] = value
        rows.append(row_data)

    return rows

def clean_country_name(name: str) -> str:
    if pd.isna(name):
        return name

    name = str(name)
    name = name.replace("·", "").replace("\u2007", " ").strip()
    name = re.sub(r"\s+", " ", name).strip()
    return name

def is_aggregate_area(name: str) -> bool:
    if pd.isna(name):
        return True

    name_lower = str(name).lower()
    aggregate_keywords = [
        "oecd",
        "non-oecd",
        "european union",
        "countries",
        "world",
        "total",
        "average"
    ]
    return any(k in name_lower for k in aggregate_keywords)

# =========================================================
# EXTRACT
# =========================================================
with zipfile.ZipFile(INPUT_FILE, "r") as z:
    rows = load_sheet_rows(z, "xl/worksheets/sheet1.xml")

measure_text = rows[1].get(2) if len(rows) > 1 else None
unit_text = rows[2].get(2) if len(rows) > 2 else None
breakdown_text = rows[3].get(2) if len(rows) > 3 else None

print("Measure:", measure_text)
print("Unit:", unit_text)
print("Breakdown:", breakdown_text)

# Excel row 5 = years, Excel row 6 = area label, Excel row 7+ = data
header_row = rows[4]
area_header_row = rows[5]

year_map = {}
for col_idx, value in header_row.items():
    if col_idx >= 4:
        if value is not None and str(value).strip().isdigit():
            year_map[col_idx] = int(str(value).strip())

print("Year map:", year_map)

# =========================================================
# TRANSFORM: WIDE -> LONG
# =========================================================
records = []

for row in rows[6:]:
    area = row.get(2)
    if area is None:
        continue

    area = clean_country_name(area)

    if area.startswith("©") or area == "ICT Access and Usage by Individuals":
        continue

    if (not KEEP_AGGREGATES) and is_aggregate_area(area):
        continue

    for col_idx, year in year_map.items():
        raw_value = row.get(col_idx)

        if raw_value in (None, ""):
            continue

        try:
            value = float(raw_value)
        except ValueError:
            continue

        records.append({
            "Country_Name": area,
            "Year": year,
            "Quarter": None,                 # yearly data
            "YearQuarter": str(year),        # consistent with shared Dim_Time
            "ICT_Type": "Individuals using the Internet - last 3 months",
            "ICTUsagePercentage": value,
            "UnitOfMeasure": "Percentage of population",
            "Breakdown": breakdown_text
        })

df_long = pd.DataFrame(records)

print("\nLong-format preview:")
print(df_long.head())
print("\nShape:", df_long.shape)

if df_long.empty:
    raise ValueError("df_long is empty. Check header rows and year_map.")

# =========================================================
# DIMENSIONS
# =========================================================

# ---- Dim_Country ----
dim_country = (
    df_long[["Country_Name"]]
    .drop_duplicates()
    .sort_values("Country_Name")
    .reset_index(drop=True)
)
dim_country["Country_Key"] = range(1, len(dim_country) + 1)
dim_country = dim_country[["Country_Key", "Country_Name"]]

# ---- Dim_Time ----
dim_time = (
    df_long[["Year", "Quarter", "YearQuarter"]]
    .drop_duplicates()
    .sort_values(["Year", "YearQuarter"])
    .reset_index(drop=True)
)
dim_time["Time_Key"] = range(1, len(dim_time) + 1)
dim_time = dim_time[["Time_Key", "Year", "Quarter", "YearQuarter"]]

# ---- Dim_ICT_Usage ----
dim_ict_usage = (
    df_long[["ICT_Type"]]
    .drop_duplicates()
    .reset_index(drop=True)
)
dim_ict_usage["ICT_Key"] = range(1, len(dim_ict_usage) + 1)
dim_ict_usage = dim_ict_usage[["ICT_Key", "ICT_Type"]]

# =========================================================
# FACT TABLE
# =========================================================
fact_ictusage = (
    df_long
    .merge(dim_country, on="Country_Name", how="left")
    .merge(dim_time, on=["Year", "Quarter", "YearQuarter"], how="left")
    .merge(dim_ict_usage, on="ICT_Type", how="left")
    .copy()
)

fact_ictusage["ICT_ID"] = range(1, len(fact_ictusage) + 1)

fact_ictusage = fact_ictusage[
    ["ICT_ID", "Country_Key", "Time_Key", "ICT_Key", "ICTUsagePercentage"]
].sort_values(["Country_Key", "Time_Key"]).reset_index(drop=True)

# =========================================================
# DATA QUALITY CHECKS
# =========================================================
print("\n--- Data Quality Checks ---")
print("Null Country_Key:", fact_ictusage["Country_Key"].isna().sum())
print("Null Time_Key:", fact_ictusage["Time_Key"].isna().sum())
print("Null ICT_Key:", fact_ictusage["ICT_Key"].isna().sum())
print(
    "Duplicate business rows:",
    fact_ictusage.duplicated(subset=["Country_Key", "Time_Key", "ICT_Key"]).sum()
)

# =========================================================
# LOAD TO CSV
# =========================================================
dim_country.to_csv(OUTPUT_DIR / "dim_country.csv", index=False)
dim_time.to_csv(OUTPUT_DIR / "dim_time.csv", index=False)
dim_ict_usage.to_csv(OUTPUT_DIR / "dim_ict_usage.csv", index=False)
fact_ictusage.to_csv(OUTPUT_DIR / "fact_ictusage.csv", index=False)
df_long.to_csv(OUTPUT_DIR / "stg_ict_usage_long.csv", index=False)

print("\nFiles saved to:", OUTPUT_DIR)
for file in OUTPUT_DIR.iterdir():
    print("-", file.name)
