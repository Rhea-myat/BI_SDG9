import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
import re
# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(__file__).resolve().parent / "Datasets_v2"
INPUT_FILE = BASE_DIR / "Patents by Technology.xlsx"
OUTPUT_DIR = Path(__file__).resolve().parent / "output_patent"
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
    x = str(x).replace("\xa0", " ")
    x = re.sub(r"\s+", " ", x).strip()
    return x

# =========================================================
# EXTRACT
# =========================================================
with zipfile.ZipFile(INPUT_FILE, "r") as z:
    sheet_path = get_sheet_path(z, "Table")
    print("Using sheet path:", sheet_path)
    rows = load_sheet_rows(z, sheet_path)

print("Total rows loaded:", len(rows))

print("\nPreview first 12 rows:")
for i, r in enumerate(rows[:12], start=1):
    print(f"Excel row {i}: {r}")

# =========================================================
# HEADER ROWS
# Excel row 3 = Year
# Excel row 4 = Patent authority
# Excel row 5 = Measure
# Excel row 6 = Agent role
# Excel row 8 = Reference area
# Excel row 9+ = data
# =========================================================
year_row = rows[2]
authority_row = rows[3]
measure_row = rows[4]
agent_row = rows[5]

# Build metadata map for each data column
column_meta = {}

for col_idx, year_val in year_row.items():
    year_clean = clean_text(year_val)

    if col_idx < 4:
        continue
    if not year_clean or not str(year_clean).isdigit():
        continue

    column_meta[col_idx] = {
        "Year": int(year_clean),
        "PatentAuthority": clean_text(authority_row.get(col_idx)),
        "PatentMeasure": clean_text(measure_row.get(col_idx)),
        "AgentRole": clean_text(agent_row.get(col_idx))
    }

print("\nDetected data columns:", len(column_meta))
print("Sample metadata:")
for k in list(column_meta.keys())[:5]:
    print(k, column_meta[k])

# =========================================================
# TRANSFORM: WIDE -> LONG
# =========================================================
records = []

for row in rows[8:]:  # Excel row 9 onward
    country = clean_text(row.get(2))

    if not country:
        continue
    if country.startswith("©"):
        continue
    if country in ["Patents by technology", "Reference area"]:
        continue

    for col_idx, meta in column_meta.items():
        raw_value = row.get(col_idx)

        if raw_value in (None, "", " "):
            continue

        try:
            value = float(raw_value)
        except ValueError:
            continue

        records.append({
            "Country_Name": country,
            "Year": meta["Year"],
            "Quarter": None,
            "YearQuarter": str(meta["Year"]),
            "PatentAuthority": meta["PatentAuthority"],
            "PatentMeasure": meta["PatentMeasure"],
            "AgentRole": meta["AgentRole"],
            "PatentCount": value
        })

df_long = pd.DataFrame(records)

print("\nLong-format preview:")
print(df_long.head())
print("\nShape:", df_long.shape)

if df_long.empty:
    raise ValueError("No rows were parsed into df_long.")

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

dim_time = (
    df_long[["Year", "Quarter", "YearQuarter"]]
    .drop_duplicates()
    .sort_values(["Year", "YearQuarter"])
    .reset_index(drop=True)
)
dim_time["Time_Key"] = range(1, len(dim_time) + 1)
dim_time = dim_time[["Time_Key", "Year", "Quarter", "YearQuarter"]]

dim_patentauthority = (
    df_long[["PatentAuthority"]]
    .drop_duplicates()
    .sort_values("PatentAuthority")
    .reset_index(drop=True)
)
dim_patentauthority["PatentAuthorityKey"] = range(1, len(dim_patentauthority) + 1)
dim_patentauthority = dim_patentauthority[["PatentAuthorityKey", "PatentAuthority"]]

dim_patentmeasure = (
    df_long[["PatentMeasure"]]
    .drop_duplicates()
    .sort_values("PatentMeasure")
    .reset_index(drop=True)
)
dim_patentmeasure["PatentMeasureKey"] = range(1, len(dim_patentmeasure) + 1)
dim_patentmeasure = dim_patentmeasure[["PatentMeasureKey", "PatentMeasure"]]

dim_agentrole = (
    df_long[["AgentRole"]]
    .drop_duplicates()
    .sort_values("AgentRole")
    .reset_index(drop=True)
)
dim_agentrole["AgentRoleKey"] = range(1, len(dim_agentrole) + 1)
dim_agentrole = dim_agentrole[["AgentRoleKey", "AgentRole"]]

# =========================================================
# FACT TABLE
# =========================================================
fact_patent = (
    df_long
    .merge(dim_country, on="Country_Name", how="left")
    .merge(dim_time, on=["Year", "Quarter", "YearQuarter"], how="left")
    .merge(dim_patentauthority, on="PatentAuthority", how="left")
    .merge(dim_patentmeasure, on="PatentMeasure", how="left")
    .merge(dim_agentrole, on="AgentRole", how="left")
    .copy()
)

fact_patent["Patent_ID"] = range(1, len(fact_patent) + 1)

fact_patent = fact_patent[[
    "Patent_ID",
    "Country_Key",
    "Time_Key",
    "PatentAuthorityKey",
    "PatentMeasureKey",
    "AgentRoleKey",
    "PatentCount"
]].sort_values(
    ["Country_Key", "Time_Key", "PatentAuthorityKey", "PatentMeasureKey", "AgentRoleKey"]
).reset_index(drop=True)

# =========================================================
# DATA QUALITY CHECKS
# =========================================================
print("\n--- Data Quality Checks ---")
print("Null Country_Key:", fact_patent["Country_Key"].isna().sum())
print("Null Time_Key:", fact_patent["Time_Key"].isna().sum())
print("Null PatentAuthorityKey:", fact_patent["PatentAuthorityKey"].isna().sum())
print("Null PatentMeasureKey:", fact_patent["PatentMeasureKey"].isna().sum())
print("Null AgentRoleKey:", fact_patent["AgentRoleKey"].isna().sum())

dup_count = fact_patent.duplicated(
    subset=[
        "Country_Key",
        "Time_Key",
        "PatentAuthorityKey",
        "PatentMeasureKey",
        "AgentRoleKey"
    ]
).sum()
print("Duplicate business rows:", dup_count)

# =========================================================
# LOAD
# =========================================================
df_long.to_csv(OUTPUT_DIR / "stg_patent_long.csv", index=False)
dim_country.to_csv(OUTPUT_DIR / "dim_country.csv", index=False)
dim_time.to_csv(OUTPUT_DIR / "dim_time.csv", index=False)
dim_patentauthority.to_csv(OUTPUT_DIR / "dim_patentauthority.csv", index=False)
dim_patentmeasure.to_csv(OUTPUT_DIR / "dim_patentmeasure.csv", index=False)
dim_agentrole.to_csv(OUTPUT_DIR / "dim_agentrole.csv", index=False)
fact_patent.to_csv(OUTPUT_DIR / "fact_patent.csv", index=False)

print("\nFiles saved to:", OUTPUT_DIR)
for f in OUTPUT_DIR.iterdir():
    print("-", f.name)