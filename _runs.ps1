cd C:\ACTT_COMMAND
.\ACTT_CONVERT_RUNNER.ps1 -SnapshotName "SOX_FY25_20250327"
.\ACTT_CONVERT_RUNNER.ps1 -SnapshotName "SOX_FY26_20260324"



<#
Unblock-File -Path "C:\ACTT_COMMAND\Load-ACTTCsvToSql.ps1"
Select-String -Path "C:\ACTT_COMMAND\Load-ACTTCsvToSql.ps1" -Pattern "OnlyBaseNames"

#>

.\Load-ACTTCsvToSql.ps1 -CsvRoot "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY25_20250327" -OnlyBaseNames "ACTT_Platform_MSSQL_All" -AllowNewColumns
.\Load-ACTTCsvToSql.ps1 -CsvRoot "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY26_20260324" -OnlyBaseNames "ACTT_Platform_MSSQL_All" -AllowNewColumns