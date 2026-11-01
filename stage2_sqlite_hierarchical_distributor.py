import os
import re
import sqlite3
import warnings
import logging
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("Stage2_LocalDistributor")
# 1. LOCAL STAGING CONFIG
CURRENT_CUTOFF = os.getenv("REPORTING_CUTOFF", "15092026_vig2026")
LOCAL_SQLITE_PATH = os.getenv("LOCAL_SQLITE_PATH", f"./staging_db/Staging_Snapshot_{CURRENT_CUTOFF}.db")
REFERENCE_DB_PATH = os.getenv("REFERENCE_DB_PATH", "./staging_db/Static_Reference_Catalog.db")
OUTPUT_ROOT_DIR = os.getenv("OUTPUT_ROOT_DIR", "./output_deliverables")

REQUIRED_BENEFICIARY_COLS = [
    "RegionCode", "RegionName", "ContractYear", "ContractNumber",
    "ContractorEntity", "ServiceDepartment", "ServiceMunicipality",
    "Service", "BeneficiaryType", "BeneficiaryId", "DocumentType",
    "EnrollmentStatus"
]

SERVICE_DISTRIBUTION_RULES = [
    {
        "service_name": "PROGRAM_ALPHA - CENTERS",
        "surveys": [
            "BASELINE_COMPLIANCE_SURVEY_ALPHA",
            "FOLLOWUP_COMPLIANCE_SURVEY_ALPHA"
        ],
        "include_sessions": True,
        "include_metrics": False
    },
    {
        "service_name": "INCLUSION_SUPPORT - FIXED_CENTERS",
        "surveys": [
            "BASELINE_INCLUSION_FIXED",
            "FOLLOWUP_INCLUSION_FIXED",
            "BARRIER_DIAGNOSTIC_FIXED",
            "FOLLOWUP_BARRIER_DIAGNOSTIC_FIXED"
        ],
        "include_sessions": False,
        "include_metrics": True
    }
]



# 2. FILESYSTEM & REGEX UTILITIES
def sanitize_path_segment(name: str) -> str:
    """Strips OS-reserved characters to guarantee safe directory and file creation."""
    if pd.isna(name):
        return "UNASSIGNED_REGION"
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()


def build_resilient_like_pattern(service_name: str) -> str:
    """Replaces typographic dashes and surrounding spaces with SQL wildcards (%)."""
    return re.sub(r"\s*[-–—]\s*", "%", service_name)


#  PIVOT & MULTI-SHEET GENERATion
def generate_hierarchical_deliverables() -> None:
    if not os.path.exists(LOCAL_SQLITE_PATH):
        raise FileNotFoundError(f"Local SQLite staging database not found: {LOCAL_SQLITE_PATH}")

    conn = sqlite3.connect(LOCAL_SQLITE_PATH)

    # Pre-load static reference catalog once into RAM to eliminate redundant disk reads
    df_ref_catalog = pd.DataFrame()
    if os.path.exists(REFERENCE_DB_PATH):
        with sqlite3.connect(REFERENCE_DB_PATH) as conn_ref:
            cursor = conn_ref.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            if tables:
                df_ref_catalog = pd.read_sql(f"SELECT * FROM [{tables[0][0]}]", conn_ref)
                logger.info(f"Pre-loaded static reference table ({len(df_ref_catalog)} records) into RAM.")

    t_ben = f"Ops_Beneficiaries_{CURRENT_CUTOFF}"
    t_diag = f"Ops_Diagnostics_{CURRENT_CUTOFF}"
    t_hr = f"Ops_HumanResources_{CURRENT_CUTOFF}"
    t_enc = f"Ops_SessionAttendance_{CURRENT_CUTOFF}_"
    t_metrics = f"Ops_NutritionalMetrics_{CURRENT_CUTOFF}_"

    for srv_cfg in SERVICE_DISTRIBUTION_RULES:
        exact_service = srv_cfg["service_name"]
        like_pattern = build_resilient_like_pattern(exact_service)
        logger.info(f"Processing in-memory consolidation for: {exact_service}")

        cols_sql = ", ".join(REQUIRED_BENEFICIARY_COLS)
        df_ben = pd.read_sql(f"SELECT {cols_sql} FROM {t_ben} WHERE Service LIKE ?", conn, params=[like_pattern])
        if df_ben.empty:
            continue

        surveys = srv_cfg["surveys"]
        placeholders = ",".join(["?"] * len(surveys))
        q_diag = f"""
            SELECT BeneficiaryId, SurveyName
            FROM {t_diag}
            WHERE Service LIKE ? AND (SequenceOrder = 1 OR SequenceOrder = '1')
              AND SurveyName IN ({placeholders})
        """
        df_diag = pd.read_sql(q_diag, conn, params=[like_pattern] + surveys)

        # Vectorized in-memory pivot to flag survey completion per beneficiary
        if not df_diag.empty:
            df_diag["Completed"] = 1
            df_diag_pivot = df_diag.pivot_table(
                index="BeneficiaryId",
                columns="SurveyName",
                values="Completed",
                aggfunc="max",
                fill_value=0
            ).reset_index()
        else:
            df_diag_pivot = pd.DataFrame(columns=["BeneficiaryId"] + surveys)

        df_ben_final = pd.merge(df_ben, df_diag_pivot, on="BeneficiaryId", how="left")
        for col in surveys:
            if col not in df_ben_final.columns:
                df_ben_final[col] = 0
            df_ben_final[col] = df_ben_final[col].fillna(0).astype(int)

        df_hr = pd.read_sql(f"SELECT * FROM {t_hr} WHERE Service LIKE ?", conn, params=[like_pattern])
        df_enc = pd.read_sql(f"SELECT * FROM {t_enc} WHERE Service LIKE ?", conn, params=[like_pattern]) if srv_cfg["include_sessions"] else pd.DataFrame()
        df_met = pd.read_sql(f"SELECT * FROM {t_metrics} WHERE Service LIKE ?", conn, params=[like_pattern]) if srv_cfg["include_metrics"] else pd.DataFrame()

        # Partition and export multi-sheet Excel workbooks organized by Service -> Region -> Contract
        contracts = df_ben_final["ContractNumber"].dropna().unique()
        logger.info(f"Generating {len(contracts)} partitioned contract workbooks...")

        for contract in contracts:
            contract_ben = df_ben_final[df_ben_final["ContractNumber"] == contract]
            region_folder = sanitize_path_segment(contract_ben["RegionName"].iloc[0])
            clean_contract_file = sanitize_path_segment(contract)

            target_dir = os.path.join(OUTPUT_ROOT_DIR, sanitize_path_segment(exact_service), region_folder)
            os.makedirs(target_dir, exist_ok=True)
            file_path = os.path.join(target_dir, f"{clean_contract_file}.xlsx")

            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                contract_ben.to_excel(writer, sheet_name="Beneficiaries", index=False)
                df_hr[df_hr["ContractNumber"] == contract].to_excel(writer, sheet_name="Human_Resources", index=False)

                if srv_cfg["include_sessions"]:
                    df_enc[df_enc["ContractNumber"] == contract].to_excel(writer, sheet_name="Session_Attendance", index=False)
                    if not df_ref_catalog.empty:
                        df_ref_catalog.to_excel(writer, sheet_name="Reference_Catalog", index=False)

                if srv_cfg["include_metrics"]:
                    df_met[df_met["ContractNumber"] == contract].to_excel(writer, sheet_name="Operational_Metrics", index=False)

    conn.close()
    logger.info("Stage 2 hierarchical distribution completed successfully.")


if __name__ == "__main__":
    generate_hierarchical_deliverables()
