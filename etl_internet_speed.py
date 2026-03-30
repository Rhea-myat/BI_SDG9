import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

# ============================================
# CONFIG
# ============================================
BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "INTERNETSpeedFINAL.xlsx"
OUTPUT_DIR = BASE_DIR / "output_internet"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

COUNTRY_NAME_MAP = {
    "AT": "Austria",
    "AU": "Australia",
    "BE": "Belgium",
    "BG": "Bulgaria",
    "CA": "Canada",
    "CH": "Switzerland",
    "CL": "Chile",
    "CO": "Colombia",
    "CY": "Cyprus",
    "CZ": "Czechia",
    "DE": "Germany",
    "DK": "Denmark",
    "EE": "Estonia",
    "EL": "Greece",
    "ES": "Spain",
    "FI": "Finland",
    "FR": "France",
    "HR": "Croatia",
    "HU": "Hungary",
    "IE": "Ireland",
    "IS": "Iceland",
    "IT": "Italy",
    "JP": "Japan",
    "KO": "Korea",
    "LT": "Lithuania",
    "LU": "Luxembourg",
    "LV": "Latvia",
    "ME": "Montenegro",
    "MT": "Malta",
    "NL": "Netherlands",
    "NO": "Norway",
    "NZ": "New Zealand",
    "PL": "Poland",
    "PT": "Portugal",
    "RO": "Romania",
    "SE": "Sweden",
    "SI": "Slovenia",
    "SK": "Slovakia",
    "TR": "Turkey",
    "UK": "United Kingdom",
    "US": "United States",
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


def load_selected_sheets(input_file: Path, sheet_names):
    with zipfile.ZipFile(input_file, "r") as zip_file:
        shared_strings = load_shared_strings(zip_file)
        sheet_paths = get_sheet_paths(zip_file)

        sheets = {}
        for sheet_name in sheet_names:
            sheet_path = sheet_paths[sheet_name]
            sheets[sheet_name] = load_sheet_dataframe(zip_file, sheet_path, shared_strings)
        return sheets


def map_country_names(dim_country_df: pd.DataFrame) -> pd.DataFrame:
    dim_country_df = dim_country_df.copy()
    if "Country" not in dim_country_df.columns:
        return dim_country_df

    dim_country_df["Country_Name"] = (
        dim_country_df["Country"]
        .astype(str)
        .str.strip()
        .map(lambda code: COUNTRY_NAME_MAP.get(code, code))
    )
    return dim_country_df[["Country_Key", "Country_Name"]]


print("INPUT_FILE:", INPUT_FILE)
print("Exists:", INPUT_FILE.exists())

if not INPUT_FILE.exists():
    raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

required_sheets = [
    "Fact_InternetSpeedFinal",
    "Dim_City",
    "Dim_Country",
    "Dim_Time",
    "Dim_InternetSpeed",
]
sheets = load_selected_sheets(INPUT_FILE, required_sheets)

fact_df = sheets["Fact_InternetSpeedFinal"].copy()
dim_city = sheets["Dim_City"].copy()
dim_country = map_country_names(sheets["Dim_Country"])
dim_time = sheets["Dim_Time"].copy()
dim_internetspeed = sheets["Dim_InternetSpeed"].copy()

print("Loaded sheets:", list(sheets.keys()))
print("Fact columns:", fact_df.columns.tolist())

if "REF_AREA" in fact_df.columns:
    fact_df = fact_df.drop(columns=["REF_AREA"])

fact_df.to_csv(OUTPUT_DIR / "fact_internetspeed.csv", index=False)
dim_city.to_csv(OUTPUT_DIR / "dim_city.csv", index=False)
dim_country.to_csv(OUTPUT_DIR / "dim_country.csv", index=False)
dim_time.to_csv(OUTPUT_DIR / "dim_time.csv", index=False)
dim_internetspeed.to_csv(OUTPUT_DIR / "dim_internetspeed.csv", index=False)

print("Exported fact_internetspeed.csv")
print("Exported dim_city.csv")
print("Exported dim_country.csv")
print("Exported dim_time.csv")
print("Exported dim_internetspeed.csv")
print("\nDone.")
