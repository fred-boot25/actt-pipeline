<#
.SYNOPSIS
    Loads the CSV output of Convert-ACTTFoldersToCsv_v2 / Invoke-ACTTZipToPbiCsv into SQL Server,
    with automatic schema-drift detection if Deloitte changes the extraction script's columns.

.DESCRIPTION
    - Recursively finds every .csv file under -CsvRoot (your PBI_READ snapshot output).
    - Groups files by their base file name (e.g. "quickfixes.csv", "MSDB_SYSJOBSTEPS.csv") -
      each distinct base name becomes one target table under the "actt" schema.
    - Every column loads as NVARCHAR(MAX), matching how the source ACTT script declares columns.
    - Compares each CSV's header row to the existing table's columns before loading:
        - New column in the CSV     -> logged to actt.SchemaDrift.
                                        Blocks the load for that file UNLESS -AllowNewColumns is passed,
                                        in which case the column is added via ALTER TABLE and logged as such.
        - Column missing in the CSV -> not drift, just loads NULL for that column.
    - Tags every row with LoadBatchId (one per run) and LoadDateTimeUtc, on top of whatever
      SnapshotName / ServerFolder / AuditStartDate columns the converter already added.
    - Writes one row per source file to actt.LoadHistory for auditability.

.PARAMETER CsvRoot
    Root folder containing the converted CSVs (e.g. the snapshot output folder from
    Invoke-ACTTZipToPbiCsv.ps1, or the whole PBI_READ folder to load every snapshot at once).

.PARAMETER ConnectionString
    SQL Server connection string. Defaults to the Windows-auth, encrypted connection provided.

.PARAMETER AllowNewColumns
    If set, automatically ALTER TABLE ADD any new columns found in a CSV that aren't in the
    existing target table. If not set (default), files with new columns are skipped and logged
    to actt.SchemaDrift for manual review - nothing is silently lost or misaligned.

.PARAMETER ExcludeControlFiles
    If set (default), skips the converter's own manifest/summary/parse-issue CSVs
    (files starting with "_"). Use -ExcludeControlFiles:$false to load those too.

.PARAMETER OnlyBaseNames
    Optional. Scope the load to specific source file(s)/table(s) by base name (no extension,
    case-insensitive) - e.g. -OnlyBaseNames "ACTT_Platform_MSSQL_All" loads only that file
    across every server folder under -CsvRoot, skipping everything else. Useful for re-loading
    just one table after resolving a schema-drift issue, without re-running the full snapshot.

.EXAMPLE
    .\Load-ACTTCsvToSql.ps1 -CsvRoot "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY25"

.EXAMPLE
    .\Load-ACTTCsvToSql.ps1 -CsvRoot "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX" -AllowNewColumns

.EXAMPLE
    .\Load-ACTTCsvToSql.ps1 -CsvRoot "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX\SOX_FY26_20260324" -OnlyBaseNames "ACTT_Platform_MSSQL_All" -AllowNewColumns
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvRoot,

    [Parameter(Mandatory = $false)]
    [string]$ConnectionString = "Data Source=IT-FTOMASECK2;Initial Catalog=ACTT;Encrypt=True;TrustServerCertificate=True;Integrated Security=True;",

    [Parameter(Mandatory = $false)]
    [switch]$AllowNewColumns,

    [Parameter(Mandatory = $false)]
    [bool]$ExcludeControlFiles = $true,

    [Parameter(Mandatory = $false)]
    [string[]]$OnlyBaseNames
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic

$loadBatchId = [guid]::NewGuid()
$loadDateTimeUtc = [DateTime]::UtcNow

function Get-SafeSqlIdentifier {
    param([Parameter(Mandatory = $true)][string]$Name)

    $safe = $Name -replace '[^A-Za-z0-9_]', '_'

    if ($safe -match '^[0-9]') {
        $safe = "T_$safe"
    }

    return $safe
}

function Get-CsvHeaderFields {
    param([Parameter(Mandatory = $true)][string]$Path)

    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(",")
        $parser.HasFieldsEnclosedInQuotes = $true

        if ($parser.EndOfData) {
            return @()
        }

        return @($parser.ReadFields())
    }
    finally {
        $parser.Close()
    }
}

function Test-IsEmptyPlaceholderCsv {
    param([Parameter(Mandatory = $true)][string[]]$HeaderFields)

    # Convert-ACTTFoldersToCsv_v2.ps1 writes this exact shape when a source .actt file has
    # zero data rows - it's the converter's own "nothing to report" marker, not real data,
    # and must never be mistaken for Deloitte adding/removing a column (schema drift).
    $placeholderShape = @('SnapshotName', 'ServerFolder', 'ServerName', 'SourceFile', 'SourceSHA256', 'Note')

    if ($HeaderFields.Count -ne $placeholderShape.Count) {
        return $false
    }

    $headerSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$HeaderFields, [System.StringComparer]::OrdinalIgnoreCase)

    foreach ($col in $placeholderShape) {
        if (-not $headerSet.Contains($col)) {
            return $false
        }
    }

    return $true
}

function New-SqlConnection {
    param([Parameter(Mandatory = $true)][string]$ConnString)

    $conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
    $conn.Open()
    return $conn
}

function Invoke-SqlNonQuery {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory = $true)][string]$Sql
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Sql
    [void]$cmd.ExecuteNonQuery()
}

function Get-ExistingColumns {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory = $true)][string]$TableName
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = @"
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'actt' AND TABLE_NAME = @tbl
"@
    [void]$cmd.Parameters.AddWithValue('@tbl', $TableName)

    $reader = $cmd.ExecuteReader()
    $cols = New-Object System.Collections.Generic.List[string]

    try {
        while ($reader.Read()) {
            [void]$cols.Add($reader.GetString(0))
        }
    }
    finally {
        $reader.Close()
    }

    return $cols
}

function Test-TableExists {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory = $true)][string]$TableName
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = "SELECT 1 FROM sys.tables WHERE schema_id = SCHEMA_ID('actt') AND name = @tbl"
    [void]$cmd.Parameters.AddWithValue('@tbl', $TableName)

    return ($null -ne $cmd.ExecuteScalar())
}

function New-TargetTable {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][string[]]$Columns
    )

    $colDefs = ($Columns | ForEach-Object { "    [$_] NVARCHAR(MAX) NULL" }) -join ",`n"

    $sql = @"
CREATE TABLE actt.[$TableName]
(
    LoadRowId       INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    LoadBatchId     UNIQUEIDENTIFIER  NOT NULL,
    LoadDateTimeUtc DATETIME2         NOT NULL,
$colDefs
);
"@

    Invoke-SqlNonQuery -Connection $Connection -Sql $sql
    Write-Host "  Created table actt.$TableName with $($Columns.Count) columns."
}

function Add-MissingColumns {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][string[]]$NewColumns,
        [Parameter(Mandatory = $true)][string]$SourceFile
    )

    foreach ($col in $NewColumns) {
        if ($AllowNewColumns) {
            Invoke-SqlNonQuery -Connection $Connection -Sql "ALTER TABLE actt.[$TableName] ADD [$col] NVARCHAR(MAX) NULL;"
            $action = 'AutoAltered'
            Write-Host "  Schema drift: added new column [$col] to actt.$TableName (auto)." -ForegroundColor Yellow
        }
        else {
            $action = 'BlockedPendingReview'
            Write-Warning "Schema drift: column [$col] in $SourceFile does not exist in actt.$TableName. Load of this file SKIPPED. Re-run with -AllowNewColumns once reviewed."
        }

        $driftCmd = $Connection.CreateCommand()
        $driftCmd.CommandText = @"
INSERT INTO actt.SchemaDrift (DetectedDateTimeUtc, LoadBatchId, TableName, ColumnName, ChangeType, Action, SourceFile)
VALUES (SYSUTCDATETIME(), @batch, @tbl, @col, 'ColumnAdded', @action, @src);
"@
        [void]$driftCmd.Parameters.AddWithValue('@batch', $loadBatchId)
        [void]$driftCmd.Parameters.AddWithValue('@tbl', $TableName)
        [void]$driftCmd.Parameters.AddWithValue('@col', $col)
        [void]$driftCmd.Parameters.AddWithValue('@action', $action)
        [void]$driftCmd.Parameters.AddWithValue('@src', $SourceFile)
        [void]$driftCmd.ExecuteNonQuery()
    }
}

function Import-CsvFileToTable {
    param(
        [Parameter(Mandatory = $true)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][string[]]$AllColumns
    )

    $rows = @(Import-Csv -LiteralPath $File.FullName)

    # Trust Import-Csv's own parsed property names as ground truth for what it actually read,
    # rather than the separate header-only parse (Get-CsvHeaderFields) used for schema/drift
    # checks. They normally agree, but if a file has an unusual header (e.g. a duplicate token
    # Import-Csv auto-renames to "Col_2"), building the DataTable from Import-Csv's real output
    # avoids a source/destination column mismatch at WriteToServer time.
    $actualCsvColumns = if ($rows.Count -gt 0) { @($rows[0].PSObject.Properties.Name) } else { $AllColumns }

    $onlyInHeaderParse = @($AllColumns | Where-Object { $actualCsvColumns -notcontains $_ })
    $onlyInImportCsv   = @($actualCsvColumns | Where-Object { $AllColumns -notcontains $_ })

    if ($onlyInHeaderParse.Count -gt 0 -or $onlyInImportCsv.Count -gt 0) {
        Write-Warning "  $($File.Name): header-parse and Import-Csv disagree on columns - using Import-Csv's actual columns. Header-only: [$($onlyInHeaderParse -join ', ')]  Import-Csv-only: [$($onlyInImportCsv -join ', ')]"
    }

    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('LoadBatchId', [guid])
    [void]$table.Columns.Add('LoadDateTimeUtc', [datetime])
    foreach ($col in $actualCsvColumns) {
        [void]$table.Columns.Add($col, [string])
    }

    foreach ($row in $rows) {
        $dr = $table.NewRow()
        $dr['LoadBatchId'] = $loadBatchId
        $dr['LoadDateTimeUtc'] = $loadDateTimeUtc
        foreach ($col in $actualCsvColumns) {
            $val = $row.$col
            $dr[$col] = if ($null -eq $val) { [DBNull]::Value } else { [string]$val }
        }
        [void]$table.Rows.Add($dr)
    }

    if ($table.Rows.Count -gt 0) {
        $existingDestColumns = Get-ExistingColumns -Connection $Connection -TableName $TableName
        $existingDestSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$existingDestColumns, [System.StringComparer]::OrdinalIgnoreCase)

        $bulk = New-Object System.Data.SqlClient.SqlBulkCopy($Connection)
        $bulk.DestinationTableName = "actt.[$TableName]"
        $bulk.BulkCopyTimeout = 0      # 0 = no timeout; large tables (DB_sys_DB_permissions, DB_SYS_OBJECTS) can run 300k+ rows
        $bulk.BatchSize = 5000         # commit in chunks instead of one all-or-nothing operation

        [void]$bulk.ColumnMappings.Add('LoadBatchId', 'LoadBatchId')
        [void]$bulk.ColumnMappings.Add('LoadDateTimeUtc', 'LoadDateTimeUtc')

        foreach ($col in $actualCsvColumns) {
            if ($existingDestSet.Contains($col)) {
                [void]$bulk.ColumnMappings.Add($col, $col)
            }
            else {
                Write-Warning "  $($File.Name): column [$col] has no matching destination column in actt.$TableName - values for this column will be skipped on this load. Investigate and re-run with -AllowNewColumns if this should be added."
            }
        }
        $bulk.WriteToServer($table)
        $bulk.Close()
    }

    $snapshotName = if ($rows.Count -gt 0 -and $rows[0].PSObject.Properties.Name -contains 'SnapshotName') { $rows[0].SnapshotName } else { $null }
    $serverFolder = if ($rows.Count -gt 0 -and $rows[0].PSObject.Properties.Name -contains 'ServerFolder') { $rows[0].ServerFolder } else { $null }

    $histCmd = $Connection.CreateCommand()
    $histCmd.CommandText = @"
INSERT INTO actt.LoadHistory (LoadBatchId, LoadDateTimeUtc, SnapshotName, ServerFolder, TableName, SourceFile, [RowCount], ColumnsAddedCount)
VALUES (@batch, @dt, @snap, @srvfolder, @tbl, @src, @rc, 0);
"@
    [void]$histCmd.Parameters.AddWithValue('@batch', $loadBatchId)
    [void]$histCmd.Parameters.AddWithValue('@dt', $loadDateTimeUtc)
    $snapParam = if ([string]::IsNullOrEmpty($snapshotName)) { [DBNull]::Value } else { $snapshotName }
    $srvFolderParam = if ([string]::IsNullOrEmpty($serverFolder)) { [DBNull]::Value } else { $serverFolder }
    [void]$histCmd.Parameters.AddWithValue('@snap', $snapParam)
    [void]$histCmd.Parameters.AddWithValue('@srvfolder', $srvFolderParam)
    [void]$histCmd.Parameters.AddWithValue('@tbl', $TableName)
    [void]$histCmd.Parameters.AddWithValue('@src', $File.FullName)
    [void]$histCmd.Parameters.AddWithValue('@rc', $table.Rows.Count)
    [void]$histCmd.ExecuteNonQuery()

    Write-Host ("  Loaded {0,6} rows -> actt.{1}  ({2})" -f $table.Rows.Count, $TableName, $File.Name)
}

# ---- Main ----

$resolvedCsvRoot = (Resolve-Path -LiteralPath $CsvRoot).Path

$allCsvFiles = @(Get-ChildItem -LiteralPath $resolvedCsvRoot -Filter '*.csv' -Recurse -File)

if ($ExcludeControlFiles) {
    $allCsvFiles = @($allCsvFiles | Where-Object { -not $_.Name.StartsWith('_') })
}

if ($OnlyBaseNames -and $OnlyBaseNames.Count -gt 0) {
    $allCsvFiles = @($allCsvFiles | Where-Object {
        $thisBaseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $OnlyBaseNames | Where-Object { $_ -ieq $thisBaseName }
    })
}

if ($allCsvFiles.Count -eq 0) {
    throw "No CSV files found under $resolvedCsvRoot" + $(if ($OnlyBaseNames) { " matching -OnlyBaseNames [$($OnlyBaseNames -join ', ')]" } else { "." })
}

Write-Host "CSV root:      $resolvedCsvRoot"
Write-Host "Files found:   $($allCsvFiles.Count)"
Write-Host "Load batch:    $loadBatchId"
Write-Host "Allow new cols:$AllowNewColumns"
Write-Host ""

$conn = New-SqlConnection -ConnString $ConnectionString

try {
    $grouped = $allCsvFiles | Group-Object { Get-SafeSqlIdentifier -Name ([System.IO.Path]::GetFileNameWithoutExtension($_.Name)) }

    foreach ($group in $grouped) {
        $tableName = $group.Name
        Write-Host "Table: actt.$tableName ($($group.Count) source file(s))" -ForegroundColor Cyan

        foreach ($file in $group.Group) {
            $headerFields = Get-CsvHeaderFields -Path $file.FullName

            if ($headerFields.Count -eq 0) {
                Write-Warning "  $($file.Name) has no header row - skipped."
                continue
            }

            if (Test-IsEmptyPlaceholderCsv -HeaderFields $headerFields) {
                $placeholderRows = @(Import-Csv -LiteralPath $file.FullName)
                $snap = if ($placeholderRows.Count -gt 0) { $placeholderRows[0].SnapshotName } else { [DBNull]::Value }
                $srvFolder = if ($placeholderRows.Count -gt 0) { $placeholderRows[0].ServerFolder } else { [DBNull]::Value }

                $histCmd = $conn.CreateCommand()
                $histCmd.CommandText = @"
INSERT INTO actt.LoadHistory (LoadBatchId, LoadDateTimeUtc, SnapshotName, ServerFolder, TableName, SourceFile, [RowCount], ColumnsAddedCount)
VALUES (@batch, @dt, @snap, @srvfolder, @tbl, @src, 0, 0);
"@
                [void]$histCmd.Parameters.AddWithValue('@batch', $loadBatchId)
                [void]$histCmd.Parameters.AddWithValue('@dt', $loadDateTimeUtc)
                [void]$histCmd.Parameters.AddWithValue('@snap', $snap)
                [void]$histCmd.Parameters.AddWithValue('@srvfolder', $srvFolder)
                [void]$histCmd.Parameters.AddWithValue('@tbl', $tableName)
                [void]$histCmd.Parameters.AddWithValue('@src', $file.FullName)
                [void]$histCmd.ExecuteNonQuery()

                Write-Host ("  {0,6} rows -> actt.{1}  ({2}) [source file had no data rows - logged, table untouched]" -f 0, $tableName, $file.Name)
                continue
            }

            if (-not (Test-TableExists -Connection $conn -TableName $tableName)) {
                New-TargetTable -Connection $conn -TableName $tableName -Columns $headerFields
                $effectiveColumns = $headerFields
            }
            else {
                $existingColumns = Get-ExistingColumns -Connection $conn -TableName $tableName
                $existingSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$existingColumns, [System.StringComparer]::OrdinalIgnoreCase)

                $newColumns = @($headerFields | Where-Object { -not $existingSet.Contains($_) })

                if ($newColumns.Count -gt 0) {
                    Add-MissingColumns -Connection $conn -TableName $tableName -NewColumns $newColumns -SourceFile $file.FullName

                    if (-not $AllowNewColumns) {
                        # Skip loading this file until the drift is reviewed.
                        continue
                    }
                }

                $effectiveColumns = $headerFields
            }

            Import-CsvFileToTable -Connection $conn -TableName $tableName -File $file -AllColumns $effectiveColumns
        }

        Write-Host ""
    }
}
finally {
    $conn.Close()
}

Write-Host "Done. Query actt.LoadHistory and actt.SchemaDrift for this run using LoadBatchId = '$loadBatchId'" -ForegroundColor Green
