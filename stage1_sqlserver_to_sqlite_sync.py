import os
import re
import sqlite3
import warnings
import logging
from typing import Dict, List
import pyodbc
import pandas as pd
from openpyxl import load_workbook

# Senior Fix: Suppress non-critical Pandas SQLAlchemy warnings for clean CLI telemetry
warnings.filterwarnings("ignore", category=UserWarning)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("Stage1_MasterSync")

# 1. ENVIRONMENT & SECURITY CONFIGURATION
# Credentials strictly loaded via environment variables (Zero hardcoded secrets)
DB_SERVER = os.getenv("ENTERPRISE_DB_SERVER", "127.0.0.1")
DB_USER = os.getenv("ENTERPRISE_DB_USER", "etl_service_account")
DB_PASS = os.getenv("ENTERPRISE_DB_PASS", "")
DB_NAME = os.getenv("ENTERPRISE_DB_NAME", "Core_Operations_DW")

CURRENT_CUTOFF = os.getenv("REPORTING_CUTOFF", "15092026_vig2026")
MASTER_EXCEL_PATH = os.getenv("MASTER_TEMPLATE_PATH", "./templates/Master_Executive_Tracking.xlsx")


# 2. METADATA-DRIVEN SERVICE REGISTRY
# Maps business programs, dynamic survey rules, and exact Excel cell coordinates
SERVICE_SCHEMA_REGISTRY: List[Dict] = [
    {
        "service_name": "PROGRAM_ALPHA - CENTERS",
        "sheet": "Alpha_Centers",
        "start_row": 3,
        "end_row": 38,
        "col_contract_xl": "E",
        "hr_col_out_active": "H",
        "hr_col_out_inactive": "I",
        "hr_col_state_sql": "StaffStatus",
        "hr_col_id_sql": "StaffId",
        "base_col_out_active": "J",
        "base_col_out_inactive": "R",
        "survey_diag": "BASELINE_COMPLIANCE_SURVEY_ALPHA",
        "diag_col_out_active": "L",
        "diag_col_out_inactive": "S",
        "survey_followup": "FOLLOWUP_COMPLIANCE_SURVEY_ALPHA",
        "followup_col_out_active": "N",
        "followup_col_out_inactive": "T",
        "runs_sessions": "YES",
        "sessions_col_out_active": "P",
        "sessions_col_out_inactive": "U",
    },
    {
        "service_name": "PROGRAM_BETA - COMMUNITY_SPACES",
        "sheet": "Community_Spaces",
        "start_row": 3,
        "end_row": 15,
        "col_contract_xl": "E",
        "hr_col_out_active": "H",
        "hr_col_out_inactive": "I",
        "hr_col_state_sql": "StaffStatus",
        "hr_col_id_sql": "StaffId",
        "base_col_out_active": "J",
        "base_col_out_inactive": "R",
        "survey_diag": "BASELINE_COMPLIANCE_SURVEY_BETA",
        "diag_col_out_active": "L",
        "diag_col_out_inactive": "S",
        "survey_followup": "FOLLOWUP_COMPLIANCE_SURVEY_BETA",
        "followup_col_out_active": "N",
        "followup_col_out_inactive": "T",
        "runs_sessions": "YES",
        "sessions_col_out_active": "P",
        "sessions_col_out_inactive": "U",
    },
    {
        "service_name": "INCLUSION_SUPPORT - FIXED_CENTERS",
        "sheet": "Fixed_Centers",
        "start_row": 3,
        "end_row": 50,
        "col_contract_xl": "D",
        "hr_col_out_active": "H",
        "hr_col_out_inactive": "I",
        "hr_col_state_sql": "StaffStatus",
        "hr_col_id_sql": "StaffId",
        "base_col_out_active": "J",
        "base_col_out_inactive": "V",
        "survey_diag": "BASELINE_INCLUSION_FIXED",
        "diag_col_out_active": "L",
        "diag_col_out_inactive": "W",
        "survey_followup": "FOLLOWUP_INCLUSION_FIXED",
        "followup_col_out_active": "N",
        "followup_col_out_inactive": "X",
        "survey_diag2": "BARRIER_DIAGNOSTIC_FIXED",
        "diag_col_out_active2": "P",
        "diag_col_out_inactive2": "Y",
        "survey_followup2": "FOLLOWUP_BARRIER_DIAGNOSTIC_FIXED",
        "followup_col_out_active2": "R",
        "followup_col_out_inactive2": "Z",
        "metrics_col_out_active": "T",
        "metrics_col_out_inactive": "AA",
        "runs_sessions": "NO",
        "sessions_col_out_active": "",
        "sessions_col_out_inactive": "",
    }
]


# 3. STRING SANITIZATION & SQL NORMALIZATION
def clean_text(text) -> str:
    """Strips invisible whitespace, Unicode dashes, and floating-point artifacts (.0)."""
    if text is None:
        return ""
    if isinstance(text, (int, float)):
        text = str(int(text))
    text = str(text).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"[\s\t\n\r\xa0\-\–\—]+", "", text)


def clean_contract_id(text) -> str:
    """Restores leading zeros stripped by Excel on 10-digit numeric contract IDs."""
    cleaned = clean_text(text)
    if not cleaned:
        return ""
    if cleaned.isdigit() and len(cleaned) == 10:
        return "0" + cleaned
    return cleaned


def normalize_sql_column(column: str) -> str:
    """Generates T-SQL expression to strip spaces and multi-format hyphens on the fly."""
    return f"REPLACE(REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM({column}))), ' ', ''), '-', ''), N'–', ''), N'—', '')"

# 4. CORE DATABASE ENGINE
class DatabaseHandler:
    def __init__(self, server: str, database: str, username: str, password: str):
        drivers = [d for d in pyodbc.drivers() if "SQL Server" in d or "SQL Native Client" in d]
        if not drivers:
            raise EnvironmentError("No compatible SQL Server ODBC driver detected on host.")

        driver_to_use = sorted(drivers, key=lambda x: "ODBC Driver" in x, reverse=True)[0]
        logger.info(f"Initializing ODBC connection using driver: {driver_to_use}")

        self.conn_str = (
            f"DRIVER={{{driver_to_use}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password};"
            f"TrustServerCertificate=yes;"
        )
        self.conn = pyodbc.connect(self.conn_str)

    def get_standard_counts(
        self,
        base_table: str,
        service: str,
        contracts: List[str],
        deduplicate: bool = False,
        state_column: str = "EnrollmentStatus",
        id_column: str = "BeneficiaryId",
        survey_filter: str = None,
        trailing_underscore: bool = False
    ) -> pd.DataFrame:
        if not contracts:
            return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

        suffix = "_" if trailing_underscore else ""
        full_table = f"dbo.Ops_{base_table}_{CURRENT_CUTOFF}{suffix}"

        sql_count = f"COUNT(DISTINCT {id_column})" if deduplicate else f"COUNT({id_column})"
        placeholders = ",".join(["?"] * len(contracts))

        col_contract_sql = normalize_sql_column("ContractNumber")
        col_service_sql = normalize_sql_column("Service")
        col_state_sql = f"UPPER(LTRIM(RTRIM({state_column})))"

        survey_clause = ""
        if survey_filter:
            col_survey_sql = normalize_sql_column("SurveyName")
            survey_clause = f" AND {col_survey_sql} = ?"

        query = f"""
            SELECT
                {col_contract_sql} AS ContractNumber,
                {col_state_sql} AS Status,
                {sql_count} AS Total
            FROM {full_table}
            WHERE {col_service_sql} = ?
              AND {col_state_sql} IN ('ACTIVE', 'INACTIVE')
              AND {col_contract_sql} IN ({placeholders})
              {survey_clause}
            GROUP BY
                {col_contract_sql},
                {col_state_sql}
        """

        params = [clean_text(service)] + contracts
        if survey_filter:
            params.append(clean_text(survey_filter))

        df = pd.read_sql(query, self.conn, params=params)

        if not df.empty:
            df["ContractNumber"] = df["ContractNumber"].astype(str).apply(clean_contract_id)
            pivot_df = df.pivot_table(
                index="ContractNumber",
                columns="Status",
                values="Total",
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            for col in ["ACTIVE", "INACTIVE"]:
                if col not in pivot_df.columns:
                    pivot_df[col] = 0
            return pivot_df
        return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

    def get_survey_reconciliation_counts(self, service: str, contracts: List[str], survey_name: str) -> pd.DataFrame:
        """Cross-references primary beneficiary roster against diagnostic survey completions in RAM."""
        if not contracts:
            return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

        table_ben = f"dbo.Ops_Beneficiaries_{CURRENT_CUTOFF}"
        table_diag = f"dbo.Ops_Diagnostics_{CURRENT_CUTOFF}"
        col_service_sql = normalize_sql_column("Service")

        query_ben = f"""
            SELECT
                ContractNumber,
                UPPER(LTRIM(RTRIM(EnrollmentStatus))) AS Status,
                BeneficiaryId
            FROM {table_ben}
            WHERE {col_service_sql} = ?
        """
        df_ben = pd.read_sql(query_ben, self.conn, params=[clean_text(service)])
        if df_ben.empty:
            return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

        df_ben["ContractNumber"] = df_ben["ContractNumber"].apply(clean_contract_id)
        df_ben["BeneficiaryId"] = df_ben["BeneficiaryId"].apply(clean_text)
        df_ben["Status"] = df_ben["Status"].str.strip()

        df_ben = df_ben[df_ben["ContractNumber"].isin(contracts)]
        df_ben = df_ben[df_ben["Status"].isin(["ACTIVE", "INACTIVE"])]
        if df_ben.empty:
            return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

        survey_pattern = "%".join(survey_name.split())
        query_diag = f"""
            SELECT BeneficiaryId
            FROM {table_diag}
            WHERE {col_service_sql} = ?
              AND LTRIM(RTRIM(SurveyName)) LIKE ?
              AND (SequenceOrder = 1 OR SequenceOrder = '1')
        """
        df_diag = pd.read_sql(query_diag, self.conn, params=[clean_text(service), survey_pattern])
        if df_diag.empty:
            return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

        df_diag["BeneficiaryId"] = df_diag["BeneficiaryId"].apply(clean_text)
        df_diag = df_diag.drop_duplicates(subset=["BeneficiaryId"])

        df_merged = pd.merge(df_ben, df_diag, on="BeneficiaryId", how="inner")
        if df_merged.empty:
            return pd.DataFrame(columns=["ContractNumber", "ACTIVE", "INACTIVE"])

        pivot_df = df_merged.pivot_table(
            index="ContractNumber",
            columns="Status",
            values="BeneficiaryId",
            aggfunc=pd.Series.nunique,
            fill_value=0
        ).reset_index()

        for col in ["ACTIVE", "INACTIVE"]:
            if col not in pivot_df.columns:
                pivot_df[col] = 0
        return pivot_df

    def export_local_sqlite_staging(self, output_dir: str) -> str:
        """Snapshots remote SQL Server tables into a local SQLite database for high-speed offline processing."""
        os.makedirs(output_dir, exist_ok=True)
        sqlite_path = os.path.join(output_dir, f"Staging_Snapshot_{CURRENT_CUTOFF}.db")
        logger.info(f"Initiating SQL Server -> SQLite local snapshot: {sqlite_path}")

        tables_to_mirror = {
            f"Ops_Beneficiaries_{CURRENT_CUTOFF}": f"dbo.Ops_Beneficiaries_{CURRENT_CUTOFF}",
            f"Ops_Diagnostics_{CURRENT_CUTOFF}": f"dbo.Ops_Diagnostics_{CURRENT_CUTOFF}",
            f"Ops_HumanResources_{CURRENT_CUTOFF}": f"dbo.Ops_HumanResources_{CURRENT_CUTOFF}",
            f"Ops_SessionAttendance_{CURRENT_CUTOFF}_": f"dbo.Ops_SessionAttendance_{CURRENT_CUTOFF}_1",
            f"Ops_NutritionalMetrics_{CURRENT_CUTOFF}_": f"dbo.Ops_NutritionalMetrics_{CURRENT_CUTOFF}"
        }

        with sqlite3.connect(sqlite_path) as conn_sqlite:
            for local_name, remote_table in tables_to_mirror.items():
                logger.info(f"Mirroring remote table {remote_table} -> {local_name}...")
                df_chunk = pd.read_sql(f"SELECT * FROM {remote_table}", self.conn)
                df_chunk.to_sql(local_name, conn_sqlite, if_exists="replace", index=False)
                logger.info(f"Committed {len(df_chunk)} rows into local staging.")

        return sqlite_path

    def close(self):
        self.conn.close()

# 5. EXCEL COORDINATE INJECTION ENGINE
def is_header_cell(text) -> bool:
    text_lower = str(text).lower()
    return any(k in text_lower for k in ("contract", "contrato", "number", "id"))


def inject_metrics_into_workbook(excel_path: str, db: DatabaseHandler, config: Dict) -> None: #imp
    logger.info(f"Processing workbook injection for service: {config['service_name']}")
    wb = load_workbook(excel_path)

    if config["sheet"] not in wb.sheetnames:
        logger.warning(f"Target worksheet '{config['sheet']}' not found. Skipping.")
        return

    ws = wb[config["sheet"]]
    col_contract = config["col_contract_xl"]
    valid_rows = []
    extracted_contracts = []

    for row in range(config["start_row"], config["end_row"] + 1):
        cell_val = ws[f"{col_contract}{row}"].value
        if not cell_val or is_header_cell(cell_val):
            continue
        cleaned_id = clean_contract_id(cell_val)
        if cleaned_id:
            extracted_contracts.append(cleaned_id)
            valid_rows.append((row, cleaned_id))

    unique_contracts = list(set(extracted_contracts))
    if not unique_contracts:
        return

    svc = config["service_name"]

    df_hr = db.get_standard_counts(
        "HumanResources", svc, unique_contracts,
        deduplicate=True, state_column=config["hr_col_state_sql"], id_column=config["hr_col_id_sql"]
    )
    dict_hr = df_hr.set_index("ContractNumber").to_dict(orient="index") if not df_hr.empty else {}

    df_base = db.get_standard_counts("Beneficiaries", svc, unique_contracts, deduplicate=True)
    dict_base = df_base.set_index("ContractNumber").to_dict(orient="index") if not df_base.empty else {}

    dict_diag = {}
    if config.get("survey_diag"):
        df_d = db.get_survey_reconciliation_counts(svc, unique_contracts, config["survey_diag"])
        dict_diag = df_d.set_index("ContractNumber").to_dict(orient="index") if not df_d.empty else {}

    for row, contract in valid_rows:
        if config.get("hr_col_out_active") and contract in dict_hr:
            ws[f"{config['hr_col_out_active']}{row}"] = dict_hr[contract].get("ACTIVE", 0)
            ws[f"{config['hr_col_out_inactive']}{row}"] = dict_hr[contract].get("INACTIVE", 0)

        if config.get("base_col_out_active") and contract in dict_base:
            ws[f"{config['base_col_out_active']}{row}"] = dict_base[contract].get("ACTIVE", 0)
            ws[f"{config['base_col_out_inactive']}{row}"] = dict_base[contract].get("INACTIVE", 0)

        if config.get("diag_col_out_active") and contract in dict_diag:
            ws[f"{config['diag_col_out_active']}{row}"] = dict_diag[contract].get("ACTIVE", 0)
            ws[f"{config['diag_col_out_inactive']}{row}"] = dict_diag[contract].get("INACTIVE", 0)

    wb.save(excel_path)
    logger.info(f"Worksheet '{config['sheet']}' successfully updated.")


if __name__ == "__main__":
    db_engine = DatabaseHandler(DB_SERVER, DB_NAME, DB_USER, DB_PASS)
    try:
        for service_cfg in SERVICE_SCHEMA_REGISTRY:
            inject_metrics_into_workbook(MASTER_EXCEL_PATH, db_engine, service_cfg)
        db_engine.export_local_sqlite_staging("./staging_db")
    finally:
        db_engine.close()
