# Hybrid SQL Server-to-SQLite ETL & Metadata-Driven Reporting Engine

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![SQL Server](https://img.shields.io/badge/SQL_Server-T--SQL_%7C_pyodbc-CC2927?style=for-the-badge&logo=microsoftsqlserver&logoColor=white)
![SQLite](https://img.shields.io/badge/Local_Staging-SQLite3-003B57?style=for-the-badge&logo=sqlite&logoColor=white)
![Pandas](https://img.shields.io/badge/Vectorized_ETL-Pandas_%7C_OpenPyXL-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Security](https://img.shields.io/badge/Infosec-Zero_Hardcoded_Secrets-2EA44F?style=for-the-badge&logo=shield&logoColor=white)

> **Summary:** A production-grade, two-stage ETL and automated reporting framework engineered to eliminate database I/O bottlenecks during massive operational closures. By decoupling extraction from heavy analytical transformations via a local SQLite staging layer and a metadata-driven coordinate engine, this architecture reduced a **3-day manual consolidation cycle down to a 10-minute automated execution**.

---

## 1. System Architecture (Two-Stage Decoupled Pipeline)

Querying a remote production **SQL Server** instance hundreds of times inside a loop to generate contract-level multi-sheet Excel deliverables causes severe network saturation, lock contention, and connection timeouts. 

This engine solves the problem by splitting the workload into two specialized stages:

```mermaid
flowchart TB
    subgraph PROD ["1. Production Layer (Remote Network)"]
        SQL[("SQL Server DW\n(Raw Operational Tables)")]
    end

    subgraph STAGE1 ["2. Stage 1: Master Sync & KPI Injection (stage1_sqlserver_to_sqlite_sync.py)"]
        ODBC["pyodbc Connection Pool\n+ Dynamic Driver Resolution"]
        TSQL["On-the-Fly T-SQL Normalization\n& Regex Contract Rescue"]
        META["SERVICE_SCHEMA_REGISTRY\n(Metadata Coordinate Engine)"]
        MASTER_XL["Master Executive Workbook\n(Exact Cell Coordinate Injection)"]
    end

    subgraph STAGING ["3. Local Staging Layer (Disk & RAM)"]
        SQLITE[("Local SQLite Snapshot\n(Staging_Snapshot_Cutoff.db)")]
        REF_DB[("Static Reference DB\n(Pre-loaded once in RAM)")]
    end

    subgraph STAGE2 ["4. Stage 2: In-Memory Pivot & Distribution (stage2_sqlite_hierarchical_distributor.py)"]
        PANDAS["Vectorized Pandas Engine\n(Inner/Left Merges & Binary Pivots)"]
        TREE["Automated Directory Partitioner\n(/Service/Region/Contract_ID.xlsx)"]
    end

    SQL -->|"Parameterized Batch Queries"| ODBC
    ODBC --> TSQL
    TSQL <--> META
    TSQL -->|"OpenPyXL Coordinate Mapping"| MASTER_XL
    ODBC -->|"Full Table Mirroring (I/O Offload)"| SQLITE
    SQLITE -->|"Zero-Latency Local Reads"| PANDAS
    REF_DB -->|"Single RAM Cache Read"| PANDAS
    PANDAS -->|"Multi-Sheet Workbook Generation"| TREE
