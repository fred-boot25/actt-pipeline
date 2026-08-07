.\Backfill-ZipIntakeHashes.ps1 -ZipRoot "Q:\ftomaseck\ACTT_FILES\SOX_FY25_20250327"
.\Backfill-ZipIntakeHashes.ps1 -ZipRoot "Q:\ftomaseck\ACTT_FILES\SOX_FY26_20260324"
.\Load-ZipIntakeToSql.ps1 -ManifestPath "Q:\ftomaseck\ACTT_FILES\SOX_FY26_20260324\_Extracted\_ACTT_ZipIntakeManifest.csv"
.\Load-ACTTCsvToSql.ps1 -CsvRoot "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY26_20260324"


Get-Content -TotalCount 1 "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY26_20260324\BBCIDB_ACTT_MSSQL\ACTT_Platform_MSSQL_All.csv"
Get-Content -TotalCount 1 "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY26_20260324\BBGPSQL_ACTT_MSSQL\ACTT_Platform_MSSQL_All.csv"

Get-Item "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY26_20260324\BBGPSQL_ACTT_MSSQL\ACTT_Platform_MSSQL_All.csv" | Select-Object LastWriteTime