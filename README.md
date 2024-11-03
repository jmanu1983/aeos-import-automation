# AEOS Import Automation

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?logo=powershell&logoColor=white)
![SQL Server](https://img.shields.io/badge/SQL%20Server-Staging%20%2B%20SP-CC2927?logo=microsoftsqlserver&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

A production-grade **PowerShell automation script** for importing accredited contractor personnel into the **Nedap AEOS** access control system. Handles CSV parsing, SQL Server staging, SOAP vendor management, and full audit logging.

## Features

- **Robust CSV Parsing** — Handles multiple encodings (UTF-8, Windows-1252, UTF-16), BOM detection, and mojibake repair
- **Smart Header Mapping** — Auto-detects column positions regardless of header naming variations
- **SQL Server Integration** — Schema introspection, BulkCopy staging, stored procedure execution
- **SOAP Vendor Management** — Automatic vendor creation/update via AEOS web services with retry logic
- **Vendor Lifecycle** — Unblocks vendors in DB, polls until vendors are synced
- **Full Audit Trail** — Detailed logging with UPSERT/BLOCK tracking per person
- **Secure Credentials** — Uses PowerShell CliXml encrypted credentials (Windows DPAPI)
- **Configurable** — All settings externalized to a JSON config file

## Architecture

```
CSV File (prestataires.csv)
    │
    ▼
┌──────────────────────┐
│  Parse & Validate    │  ← Encoding detection, header mapping
│  (PowerShell)        │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  SQL Server Staging  │  ← BulkCopy into PJ_PRESTATAIRES_IMPORT_STAGE
│  (SqlBulkCopy)       │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Vendor SOAP Sync    │  ← Add/Change vendors via AEOS SOAP API
│  (Optional)          │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Stored Procedure    │  ← dbo.sp_LoadAccreditesToImport
│  (SQL Server)        │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Archive & Audit     │  ← Move CSV, log UPSERT/BLOCK details
└──────────────────────┘
```

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Language | PowerShell 5.1+ |
| Database | SQL Server (SqlClient) |
| SOAP | Native HTTP (Invoke-WebRequest) |
| Credentials | PowerShell CliXml (DPAPI encryption) |
| Scheduling | Windows Task Scheduler |

## Installation

```bash
git clone https://github.com/jmanu1983/aeos-import-automation.git
cd aeos-import-automation
```

## Configuration

1. Copy the example configuration:
   ```powershell
   Copy-Item config\accredites.config.json.example config\accredites.config.json
   ```

2. Edit the config file with your environment settings (SQL Server, paths, SOAP URL).

3. Create encrypted credential files:
   ```powershell
   # SQL Server credentials
   Get-Credential | Export-Clixml -Path config\sql_cred.xml

   # AEOS SOAP credentials
   Get-Credential | Export-Clixml -Path secrets\aeos-soap.cred.clixml
   ```

## Usage

```powershell
# Run with default config path
.\bin\01-Import-Accredites.ps1

# Run with a custom config
.\bin\01-Import-Accredites.ps1 -ConfigPath "D:\custom\config.json"
```

### Scheduled Execution

Set up a Windows Task Scheduler job to run the script at regular intervals:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "D:\path\to\bin\01-Import-Accredites.ps1"
```

## Project Structure

```
aeos-import-automation/
├── bin/
│   └── 01-Import-Accredites.ps1    # Main import script
├── config/
│   ├── accredites.config.json.example  # Configuration template
│   └── PJ_PRESTATAIRES_IMPORT_LOAD.sql # Stored procedure
├── secrets/                         # Encrypted credentials (not in VCS)
├── logs/                            # Runtime logs (not in VCS)
└── README.md
```

## License

This project is licensed under the MIT License.
