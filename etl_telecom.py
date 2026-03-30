import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
import re

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(__file__).resolve().parent / "Datasets_v2"
INPUT_FILE = BASE_DIR / "Broadband and telecom databases_v2.xlsx"
OUTPUT_DIR = Path(__file__).resolve().parent / "output_telecom"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print("INPUT_FILE:", INPUT_FILE)
print("Exists:", INPUT_FILE.exists())

# =========================================================
# XML HELPERS
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
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return shared_strings

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    for si in root.findall("a:si", NS):
        parts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(t.text or "")
        shared_strings.append("".join(parts))
    return shared_strings

def get_sheet_path(zip_file, target_sheet_name="Table"):
    workbook_root = ET.fromstring(zip_file.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))

    rel_map = {}
    for rel in rels_root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            rel_map[rid] = target

    for sheet in workbook_root.findall(".//a:sheets/a:sheet", NS):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if name == target_sheet_name and rid in rel_map:
            target = rel_map[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target
            return target

    raise ValueError(f"Sheet '{target_sheet_name}' not found.")

def load_sheet_rows(zip_file, sheet_path):
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

def clean_text(x):
    if x is None:
        return None
    x = str(x).strip()
    x = re.sub(r"\s+", " ", x)
    return x

# =========================================================
# EXTRACT
# =========================================================
with zipfile.ZipFile(INPUT_FILE, "r") as z:
    sheet_path = get_sheet_path(z, "Table")
    print("Using sheet path:", sheet_path)
    rows = load_sheet_rows(z, sheet_path)

print("Total rows loaded:", len(rows))

print("\nPreview first rows:")
for i, r in enumerate(rows[:12], start=1):
    print(f"Excel row {i}: {r}")

# =========================================================
# GET TIME COLUMNS FROM EXCEL ROW 3
# =========================================================
# Python index 2 = Excel row 3
time_header_row = rows[2]

time_map = {}
for col_idx, value in time_header_row.items():
    val = clean_text(value)
    if val and re.fullmatch(r"\d{4}-Q[1-4]", val):
        time_map[col_idx] = val

print("\nDetected time_map:")
print(time_map)

if not time_map:
    raise ValueError("No quarter columns found in row 3.")

# =========================================================
# PARSE BLOCKS
# =========================================================
records = []

current_measure = None
current_network_access = None
current_network_type = None

for i, row in enumerate(rows[4:], start=5):  # start at Excel row 5
    col2 = clean_text(row.get(2))  # column B
    col3 = clean_text(row.get(3))  # column C

    if not col2:
        continue

    # Detect metadata rows
    if col2.startswith("Measure:"):
        current_measure = col2.replace("Measure:", "", 1).strip()
        continue

    if col2.startswith("Network access mode:"):
        current_network_access = col2.replace("Network access mode:", "", 1).strip()
        continue

    if col2.startswith("Network type:"):
        current_network_type = col2.replace("Network type:", "", 1).strip()
        continue

    # Skip title/footer rows
    if col2 in ["Broadband and telecom databases", "Reference area", "Time period"]:
        continue

    # Country data row
    country = col2
    unit = col3  # "Subscriptions", "Data usage"

    # If measure metadata is missing, fall back to column C
    measure_type = current_measure if current_measure else unit

    for col_idx, yq in time_map.items():
        raw_value = row.get(col_idx)
        if raw_value in (None, "", " "):
            continue

        try:
            value = float(raw_value)
        except ValueError:
            continue

        records.append({
            "Country_Name": country,
            "MeasureType": measure_type,
            "NetworkAccessType": current_network_access,
            "NetworkType": current_network_type,
            "YearQuarter": yq,
            "MeasureValue": value,
            "Unit": unit
        })

df_long = pd.DataFrame(records)

print("\nLong-format preview:")
print(df_long.head())
print("\nLong-format shape:", df_long.shape)

if df_long.empty:
    raise ValueError("No fact rows parsed. Check metadata block structure.")

# =========================================================
# CLEAN TEXT FIELDS
# =========================================================
for col in ["Country_Name", "MeasureType", "NetworkAccessType", "NetworkType", "Unit"]:
    df_long[col] = df_long[col].apply(clean_text)

df_long["Year"] = df_long["YearQuarter"].str.split("-").str[0].astype(int)
df_long["Quarter"] = df_long["YearQuarter"].str.split("-").str[1]

print("\nUnique MeasureType values:")
print(sorted(df_long["MeasureType"].dropna().unique()))

print("\nUnique Unit values:")
print(sorted(df_long["Unit"].dropna().unique()))

# =========================================================
# DIMENSIONS
# =========================================================
dim_country = (
    df_long[["Country_Name"]]
    .drop_duplicates()
    .sort_values("Country_Name")
    .reset_index(drop=True)
)
dim_country["Country_Key"] = range(1, len(dim_country) + 1)
dim_country = dim_country[["Country_Key", "Country_Name"]]

quarter_order = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
dim_time = (
    df_long[["Year", "Quarter", "YearQuarter"]]
    .drop_duplicates()
    .copy()
)
dim_time["QuarterOrder"] = dim_time["Quarter"].map(quarter_order)
dim_time = dim_time.sort_values(["Year", "QuarterOrder"]).reset_index(drop=True)
dim_time["Time_Key"] = range(1, len(dim_time) + 1)
dim_time = dim_time[["Time_Key", "Year", "Quarter", "YearQuarter"]]

dim_networktype = (
    df_long[["NetworkType"]]
    .drop_duplicates()
    .sort_values("NetworkType")
    .reset_index(drop=True)
)
dim_networktype["NetworkTypeKey"] = range(1, len(dim_networktype) + 1)
dim_networktype = dim_networktype[["NetworkTypeKey", "NetworkType"]]

dim_networkaccess = (
    df_long[["NetworkAccessType"]]
    .drop_duplicates()
    .sort_values("NetworkAccessType")
    .reset_index(drop=True)
)
dim_networkaccess["NetworkAccessKey"] = range(1, len(dim_networkaccess) + 1)
dim_networkaccess = dim_networkaccess[["NetworkAccessKey", "NetworkAccessType"]]

dim_measuretype = (
    df_long[["MeasureType"]]
    .drop_duplicates()
    .sort_values("MeasureType")
    .reset_index(drop=True)
)
dim_measuretype["Measure_Key"] = range(1, len(dim_measuretype) + 1)
dim_measuretype = dim_measuretype[["Measure_Key", "MeasureType"]]

dim_unit = (
    df_long[["Unit"]]
    .drop_duplicates()
    .sort_values("Unit")
    .reset_index(drop=True)
)
dim_unit["Unit_Key"] = range(1, len(dim_unit) + 1)
dim_unit = dim_unit[["Unit_Key", "Unit"]]

# =========================================================
# FACT
# =========================================================
fact_telecom = (
    df_long
    .merge(dim_country, on="Country_Name", how="left")
    .merge(dim_time, on=["Year", "Quarter", "YearQuarter"], how="left")
    .merge(dim_networktype, on="NetworkType", how="left")
    .merge(dim_networkaccess, on="NetworkAccessType", how="left")
    .merge(dim_measuretype, on="MeasureType", how="left")
    .merge(dim_unit, on="Unit", how="left")
    .copy()
)

fact_telecom["TelecomFact_ID"] = range(1, len(fact_telecom) + 1)

fact_telecom = fact_telecom[[
    "TelecomFact_ID",
    "Country_Key",
    "Time_Key",
    "NetworkTypeKey",
    "NetworkAccessKey",
    "Measure_Key",
    "Unit_Key",
    "MeasureValue"
]].sort_values(
    ["Country_Key", "Time_Key", "NetworkTypeKey", "NetworkAccessKey", "Measure_Key"]
).reset_index(drop=True)

# =========================================================
# DATA QUALITY
# =========================================================
print("\n--- Data Quality Checks ---")
print("Null Country_Key:", fact_telecom["Country_Key"].isna().sum())
print("Null Time_Key:", fact_telecom["Time_Key"].isna().sum())
print("Null NetworkTypeKey:", fact_telecom["NetworkTypeKey"].isna().sum())
print("Null NetworkAccessKey:", fact_telecom["NetworkAccessKey"].isna().sum())
print("Null Measure_Key:", fact_telecom["Measure_Key"].isna().sum())

dup_count = fact_telecom.duplicated(
    subset=["Country_Key", "Time_Key", "NetworkTypeKey", "NetworkAccessKey", "Measure_Key"]
).sum()
print("Duplicate business rows:", dup_count)

print("\nMeasure types loaded:")
print(dim_measuretype)

# =========================================================
# LOAD
# =========================================================
df_long.to_csv(OUTPUT_DIR / "stg_telecom_long.csv", index=False)
dim_country.to_csv(OUTPUT_DIR / "dim_country.csv", index=False)
dim_time.to_csv(OUTPUT_DIR / "dim_time.csv", index=False)
dim_networktype.to_csv(OUTPUT_DIR / "dim_networktype.csv", index=False)
dim_networkaccess.to_csv(OUTPUT_DIR / "dim_networkaccesstype.csv", index=False)
dim_measuretype.to_csv(OUTPUT_DIR / "dim_measuretype.csv", index=False)
dim_unit.to_csv(OUTPUT_DIR / "dim_unit.csv", index=False)
fact_telecom.to_csv(OUTPUT_DIR / "fact_telecom.csv", index=False)

print("\nFiles saved to:", OUTPUT_DIR)
for f in OUTPUT_DIR.iterdir():
    print("-", f.name)
