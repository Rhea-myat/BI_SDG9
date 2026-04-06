from pathlib import Path
import shutil

import pandas as pd

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(__file__).resolve().parent

TELECOM_DIR = BASE_DIR / "output_telecom"
PATENT_DIR = BASE_DIR / "output_patent"
ICT_DIR = BASE_DIR / "output_ict_usage"
RND_DIR = BASE_DIR / "output_rnd"
INTERNET_DIR = BASE_DIR / "output_internet"
FINAL_DIR = BASE_DIR / "output_final"

FINAL_DIR.mkdir(parents=True, exist_ok=True)


def load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing required file: {path}")
    return pd.read_csv(path)


def normalize_country_name(name):
    if pd.isna(name):
        return name
    return str(name).replace("·", "").strip()


def build_shared_dim_country(country_frames: list[pd.DataFrame]) -> pd.DataFrame:
    all_countries = pd.concat(
        [frame[["Country_Name"]] for frame in country_frames],
        ignore_index=True,
    ).dropna()
    all_countries["Country_Name"] = all_countries["Country_Name"].map(normalize_country_name)

    dim_country = (
        all_countries.drop_duplicates()
        .sort_values("Country_Name")
        .reset_index(drop=True)
    )
    dim_country["Country_Key"] = range(1, len(dim_country) + 1)
    return dim_country[["Country_Key", "Country_Name"]]


def build_shared_dim_time(time_frames: list[pd.DataFrame]) -> pd.DataFrame:
    normalized_frames = []
    for frame in time_frames:
        time_frame = frame[["Year", "Quarter", "YearQuarter"]].copy()
        time_frame["Quarter"] = time_frame["Quarter"].where(time_frame["Quarter"].notna(), None)
        time_frame["Quarter"] = time_frame["Quarter"].astype(object)
        time_frame["YearQuarter"] = time_frame["YearQuarter"].astype(str)
        normalized_frames.append(time_frame)

    all_time = pd.concat(normalized_frames, ignore_index=True).drop_duplicates()

    quarter_order = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
    all_time["QuarterOrder"] = all_time["Quarter"].map(quarter_order).fillna(0)
    dim_time = all_time.sort_values(["Year", "QuarterOrder", "YearQuarter"]).reset_index(drop=True)
    dim_time["Time_Key"] = range(1, len(dim_time) + 1)
    return dim_time[["Time_Key", "Year", "Quarter", "YearQuarter"]]


def remap_fact_keys(
    fact_df: pd.DataFrame,
    old_dim_country: pd.DataFrame,
    old_dim_time: pd.DataFrame,
    shared_dim_country: pd.DataFrame,
    shared_dim_time: pd.DataFrame,
) -> pd.DataFrame:
    old_dim_country = old_dim_country.copy()
    old_dim_time = old_dim_time.copy()
    shared_dim_time = shared_dim_time.copy()

    old_dim_country["Country_Name"] = old_dim_country["Country_Name"].map(normalize_country_name)

    old_dim_time["Quarter"] = old_dim_time["Quarter"].where(old_dim_time["Quarter"].notna(), None)
    old_dim_time["Quarter"] = old_dim_time["Quarter"].astype(object)
    old_dim_time["YearQuarter"] = old_dim_time["YearQuarter"].astype(str)

    shared_dim_time["Quarter"] = shared_dim_time["Quarter"].where(
        shared_dim_time["Quarter"].notna(), None
    )
    shared_dim_time["Quarter"] = shared_dim_time["Quarter"].astype(object)
    shared_dim_time["YearQuarter"] = shared_dim_time["YearQuarter"].astype(str)

    fact_df = (
        fact_df
        .merge(old_dim_country, on="Country_Key", how="left")
        .merge(old_dim_time, on="Time_Key", how="left")
        .drop(columns=["Country_Key", "Time_Key"])
        .merge(shared_dim_country, on="Country_Name", how="left")
        .merge(shared_dim_time, on=["Year", "Quarter", "YearQuarter"], how="left")
    )
    return fact_df


def copy_dimension(src_dir: Path, src_name: str, dest_name: str | None = None):
    src_path = src_dir / src_name
    if not src_path.exists():
        raise FileNotFoundError(f"Missing dimension file: {src_path}")
    shutil.copy2(src_path, FINAL_DIR / (dest_name or src_name))


# =========================================================
# LOAD DATASETS
# =========================================================
telecom_dim_country = load_csv(TELECOM_DIR / "dim_country.csv")
telecom_dim_time = load_csv(TELECOM_DIR / "dim_time.csv")
telecom_fact = load_csv(TELECOM_DIR / "fact_telecom.csv")

patent_dim_country = load_csv(PATENT_DIR / "dim_country.csv")
patent_dim_time = load_csv(PATENT_DIR / "dim_time.csv")
patent_fact = load_csv(PATENT_DIR / "fact_patent.csv")

ict_dim_country = load_csv(ICT_DIR / "dim_country.csv")
ict_dim_time = load_csv(ICT_DIR / "dim_time.csv")
ict_fact = load_csv(ICT_DIR / "fact_ictusage.csv")

rnd_dim_country = load_csv(RND_DIR / "dim_country.csv")
rnd_dim_time = load_csv(RND_DIR / "dim_time.csv")
rnd_fact = load_csv(RND_DIR / "fact_rnd.csv")

internet_dim_country = load_csv(INTERNET_DIR / "dim_country.csv")
internet_dim_time = load_csv(INTERNET_DIR / "dim_time.csv")
internet_fact = load_csv(INTERNET_DIR / "fact_internetspeed.csv")

# =========================================================
# BUILD SHARED DIMENSIONS
# =========================================================
shared_dim_country = build_shared_dim_country([
    telecom_dim_country,
    patent_dim_country,
    ict_dim_country,
    rnd_dim_country,
    internet_dim_country,
])
shared_dim_country.to_csv(FINAL_DIR / "dim_country.csv", index=False)
print("Created shared dim_country.csv")

shared_dim_time = build_shared_dim_time([
    telecom_dim_time,
    patent_dim_time,
    ict_dim_time,
    rnd_dim_time,
    internet_dim_time,
])
shared_dim_time.to_csv(FINAL_DIR / "dim_time.csv", index=False)
print("Created shared dim_time.csv")

# =========================================================
# REMAP FACT TABLES
# =========================================================
fact_telecom = remap_fact_keys(
    telecom_fact,
    telecom_dim_country,
    telecom_dim_time,
    shared_dim_country,
    shared_dim_time,
)
fact_telecom = fact_telecom[[
    "TelecomFact_ID",
    "Country_Key",
    "Time_Key",
    "NetworkTypeKey",
    "NetworkAccessKey",
    "Telecom_Key",
    "Unit_Key",
    "MeasureValue",
]]
fact_telecom.to_csv(FINAL_DIR / "fact_telecom.csv", index=False)
print("Created fact_telecom.csv")

fact_patent = remap_fact_keys(
    patent_fact,
    patent_dim_country,
    patent_dim_time,
    shared_dim_country,
    shared_dim_time,
)
fact_patent = fact_patent[[
    "Patent_ID",
    "Country_Key",
    "Time_Key",
    "PatentAuthorityKey",
    "Patent_Key",
    "AgentRoleKey",
    "PatentCount",
]]
fact_patent.to_csv(FINAL_DIR / "fact_patent.csv", index=False)
print("Created fact_patent.csv")

fact_ictusage = remap_fact_keys(
    ict_fact,
    ict_dim_country,
    ict_dim_time,
    shared_dim_country,
    shared_dim_time,
)
fact_ictusage = fact_ictusage[[
    "ICT_ID",
    "Country_Key",
    "Time_Key",
    "ICT_Key",
    "ICTUsagePercentage",
]]
fact_ictusage.to_csv(FINAL_DIR / "fact_ictusage.csv", index=False)
print("Created fact_ictusage.csv")

fact_rnd = remap_fact_keys(
    rnd_fact,
    rnd_dim_country,
    rnd_dim_time,
    shared_dim_country,
    shared_dim_time,
)
fact_rnd = fact_rnd[[
    "R&D_ID",
    "Country_Key",
    "Time_Key",
    "ExpenditureAmount",
    "ResearcherNumber",
]]
fact_rnd.to_csv(FINAL_DIR / "fact_rnd.csv", index=False)
print("Created fact_rnd.csv")

fact_internetspeed = remap_fact_keys(
    internet_fact,
    internet_dim_country,
    internet_dim_time,
    shared_dim_country,
    shared_dim_time,
)
fact_internetspeed = fact_internetspeed[[
    "InternetSpeed_ID",
    "City_Key",
    "Country_Key",
    "Time_Key",
    "Internet_Speed_Type_Key",
    "InternetSpeed",
]]
fact_internetspeed.to_csv(FINAL_DIR / "fact_internetspeed.csv", index=False)
print("Created fact_internetspeed.csv")

# =========================================================
# COPY SUBJECT-SPECIFIC DIMENSIONS
# =========================================================
copy_dimension(INTERNET_DIR, "dim_city.csv")
copy_dimension(TELECOM_DIR, "dim_unit.csv")
copy_dimension(TELECOM_DIR, "dim_telecom.csv")
copy_dimension(TELECOM_DIR, "dim_networktype.csv")
copy_dimension(TELECOM_DIR, "dim_networkaccesstype.csv")
copy_dimension(PATENT_DIR, "dim_patent.csv")
copy_dimension(PATENT_DIR, "dim_agentrole.csv")
copy_dimension(PATENT_DIR, "dim_patentauthority.csv", "dim_patentauthorizationtype.csv")
copy_dimension(INTERNET_DIR, "dim_internetspeed.csv")
copy_dimension(ICT_DIR, "dim_ict_usage.csv")
print("Copied subject-specific dimensions")

print("\nIntegration complete!")
