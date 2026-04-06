import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

# ============================================
# CONFIG
# ============================================
BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "R&D RE.xlsx"
OUTPUT_DIR = BASE_DIR / "output_rnd"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

FACT_SHEET_NAME = "Fact_R&D"
RD_DIMENSION_SHEET_NAME = "Dim_R&D"
RD_FK_COLUMN = "R&D_Key"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

COUNTRY_NAME_MAP = {
    "AR": "Argentina",
    "AT": "Austria",
    "AU": "Australia",
    "BE": "Belgium",
    "BG": "Bulgaria",
    "CA": "Canada",
    "CH": "Switzerland",
    "CL": "Chile",
    "CO": "Colombia",
    "CR": "Costa Rica",
    "CZ": "Czechia",
    "DE": "Germany",
    "DK": "Denmark",
    "EE": "Estonia",
    "ES": "Spain",
    "EU27_2020": "European Union (27 countries, from 2020)",
    "FI": "Finland",
    "FR": "France",
    "GB": "United Kingdom",
    "GR": "Greece",
    "HR": "Croatia",
    "HU": "Hungary",
    "IE": "Ireland",
    "IL": "Israel",
    "IS": "Iceland",
    "IT": "Italy",
    "JP": "Japan",
    "KR": "Korea",
    "LT": "Lithuania",
    "LU": "Luxembourg",
    "LV": "Latvia",
    "MX": "Mexico",
    "NL": "Netherlands",
    "NO": "Norway",
    "NZ": "New Zealand",
    "OECD": "OECD",
    "PL": "Poland",
    "PT": "Portugal",
    "RO": "Romania",
    "RU": "Russia",
    "SE": "Sweden",
    "SG": "Singapore",
    "SI": "Slovenia",
    "SK": "Slovakia",
    "TR": "Turkey",
    "TW": "Chinese Taipei",
    "US": "United States",
    "ZA": "South Africa",
}


def col_letters_to_index(col_letters: str) -> int:
    result = 0
    for ch in col_letters:
        result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
    return result


def parse_cell_ref(cell_ref: str):
    match = re.match(r"([A-Z]+)(\d+)", cell_ref or "")
    if not match:
        return None, None
    return match.group(1), int(match.group(2))


def load_shared_strings(zip_file):
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    shared_strings = []
    for si in root.findall("a:si", NS):
        parts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(t.text or "")
        shared_strings.append("".join(parts))
    return shared_strings


def get_sheet_paths(zip_file):
    workbook_root = ET.fromstring(zip_file.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))

    rel_map = {}
    for rel in rels_root.findall("r:Relationship", REL_NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            rel_map[rid] = target if target.startswith("xl/") else f"xl/{target}"

    sheet_paths = {}
    for sheet in workbook_root.findall(".//a:sheets/a:sheet", NS):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        if name and rid in rel_map:
            sheet_paths[name] = rel_map[rid]
    return sheet_paths


def get_cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    value_elem = cell.find("a:v", NS)
    value = value_elem.text if value_elem is not None else None

    if cell_type == "s" and value is not None:
        return shared_strings[int(value)]

    if cell_type == "inlineStr":
        is_elem = cell.find("a:is", NS)
        if is_elem is not None:
            texts = [
                t.text or ""
                for t in is_elem.iter(
                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
                )
            ]
            return "".join(texts)

    return value


def load_sheet_dataframe(zip_file, sheet_path, shared_strings):
    root = ET.fromstring(zip_file.read(sheet_path))

    sheet_rows = []
    max_col_index = 0

    for row in root.findall(".//a:sheetData/a:row", NS):
        row_data = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r")
            col_letters, _ = parse_cell_ref(ref)
            if not col_letters:
                continue

            col_index = col_letters_to_index(col_letters)
            row_data[col_index] = get_cell_value(cell, shared_strings)
            max_col_index = max(max_col_index, col_index)

        if row_data:
            sheet_rows.append(row_data)

    if not sheet_rows:
        return pd.DataFrame()

    normalized_rows = []
    for row_data in sheet_rows:
        normalized_rows.append([row_data.get(col) for col in range(1, max_col_index + 1)])

    headers = normalized_rows[0]
    data_rows = normalized_rows[1:]

    columns = []
    for index, header in enumerate(headers, start=1):
        header_text = str(header).strip() if header is not None else ""
        columns.append(header_text if header_text else f"Column_{index}")

    return pd.DataFrame(data_rows, columns=columns)


def load_workbook_sheets(input_file: Path):
    with zipfile.ZipFile(input_file, "r") as zip_file:
        shared_strings = load_shared_strings(zip_file)
        sheet_paths = get_sheet_paths(zip_file)

        sheets = {}
        for sheet_name, sheet_path in sheet_paths.items():
            sheets[sheet_name] = load_sheet_dataframe(zip_file, sheet_path, shared_strings)
        return sheets


def map_country_names(dim_country_df: pd.DataFrame) -> pd.DataFrame:
    dim_country_df = dim_country_df.copy()
    if "Country_Name" not in dim_country_df.columns:
        return dim_country_df

    dim_country_df["Country_Name"] = (
        dim_country_df["Country_Name"]
        .astype(str)
        .str.strip()
        .map(lambda code: COUNTRY_NAME_MAP.get(code, code))
    )
    return dim_country_df


print("INPUT_FILE:", INPUT_FILE)
print("Exists:", INPUT_FILE.exists())

if not INPUT_FILE.exists():
    raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

# ============================================
# LOAD WORKBOOK
# ============================================
sheets = load_workbook_sheets(INPUT_FILE)
print("Sheets found:", list(sheets.keys()))

for sheet_name, df in sheets.items():
    print(f"\n--- {sheet_name} ---")
    print(df.head())
    print("Shape:", df.shape)
    print("Columns:", df.columns.tolist())

# ============================================
# GET FACT AND R&D DIMENSION
# ============================================
fact_df = sheets[FACT_SHEET_NAME].copy()
rd_dim_df = sheets[RD_DIMENSION_SHEET_NAME].copy()
dim_time_df = sheets["Dim_Year"].copy()
dim_country_df = map_country_names(sheets["Dim_Country"])

dim_time_df["Quarter"] = None
dim_time_df["YearQuarter"] = dim_time_df["Year"].astype(str)
dim_time_df = dim_time_df[["Time_Key", "Year", "Quarter", "YearQuarter"]]

# ============================================
# CHECK WHETHER R&D DIMENSION HAS ONLY ONE KEY
# ============================================
if RD_FK_COLUMN in rd_dim_df.columns:
    unique_keys = rd_dim_df[RD_FK_COLUMN].dropna().unique()
    print(f"\nUnique R&D Keys in dimension: {unique_keys}")

    if len(unique_keys) == 1:
        print("Only one R&D Key found. Dropping R&D dimension and FK from fact table.")

        if RD_FK_COLUMN in fact_df.columns:
            fact_df = fact_df.drop(columns=[RD_FK_COLUMN])

        export_rd_dimension = False
    else:
        print("More than one R&D Key found. Keeping R&D dimension and FK.")
        export_rd_dimension = True
else:
    print(f"Column '{RD_FK_COLUMN}' not found in R&D dimension.")
    export_rd_dimension = True

# ============================================
# EXPORT FACT TABLE
# ============================================
fact_output = OUTPUT_DIR / "fact_rnd.csv"
fact_df.to_csv(fact_output, index=False)
print(f"Exported fact table: {fact_output}")

# ============================================
# EXPORT OTHER DIMENSIONS
# ============================================
for sheet_name, df in sheets.items():
    if sheet_name == FACT_SHEET_NAME:
        continue

    if sheet_name == RD_DIMENSION_SHEET_NAME and not export_rd_dimension:
        print(f"Skipped exporting singleton dimension: {sheet_name}")
        continue

    if sheet_name == "Table1":
        print(f"Skipped exporting helper sheet: {sheet_name}")
        continue

    if sheet_name == "Dim_Year":
        output_file = OUTPUT_DIR / "dim_time.csv"
        dim_time_df.to_csv(output_file, index=False)
        print(f"Exported: {output_file}")
        continue

    if sheet_name == "Dim_Country":
        output_file = OUTPUT_DIR / "dim_country.csv"
        dim_country_df.to_csv(output_file, index=False)
        print(f"Exported: {output_file}")
        continue

    if sheet_name == "R&D":
        output_file = OUTPUT_DIR / "stg_rnd.csv"
        df.to_csv(output_file, index=False)
        print(f"Exported: {output_file}")
        continue

    safe_name = sheet_name.strip().replace(" ", "_").replace("&", "and")
    output_file = OUTPUT_DIR / f"{safe_name.lower()}.csv"
    df.to_csv(output_file, index=False)
    print(f"Exported: {output_file}")

print("\nDone.")
