<#
.SYNOPSIS
    Loads _ACTT_ZipIntakeManifest.csv (pre-extraction zip hashes) into actt.ZipIntake.

.DESCRIPTION
    Invoke-ACTTZipToPbiCsv.ps1 hashes every zip BEFORE extraction and writes
    _ACTT_ZipIntakeManifest.csv into both the extract root and the snapshot output folder.
    This script loads that manifest into the typed actt.ZipIntake table.

    Idempotent: (SnapshotName, ZipFileName, SHA256) is unique. Re-running against the same
    manifest is a no-op for rows already loaded - existing rows are skipped, not duplicated.

.PARAMETER ManifestPath
    Path to a single _ACTT_ZipIntakeManifest.csv file, OR a root folder to search recursively
    for every _ACTT_ZipIntakeManifest.csv under it (e.g. point it at the whole PBI_READ root
    to pick up every snapshot's manifest in one run).

.PARAMETER ConnectionString
    SQL Server connection string. Defaults to the Windows-auth, encrypted connection.

.EXAMPLE
    .\Load-ZipIntakeToSql.ps1 -ManifestPath "Q:\ftomaseck\ACTT_FILES\SOX_FY25\_Extracted\_ACTT_ZipIntakeManifest.csv"

.EXAMPLE
    .\Load-ZipIntakeToSql.ps1 -ManifestPath "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [Parameter(Mandatory = $false)]
    [string]$ConnectionString = "Data Source=IT-FTOMASECK2;Initial Catalog=ACTT;Encrypt=True;TrustServerCertificate=True;Integrated Security=True;"
)

$ErrorActionPreference = 'Stop'

$resolvedPath = (Resolve-Path -LiteralPath $ManifestPath).Path
$item = Get-Item -LiteralPath $resolvedPath

if ($item.PSIsContainer) {
    $manifestFiles = @(Get-ChildItem -LiteralPath $resolvedPath -Filter '_ACTT_ZipIntakeManifest.csv' -Recurse -File)
}
else {
    $manifestFiles = @($item)
}

if ($manifestFiles.Count -eq 0) {
    throw "No _ACTT_ZipIntakeManifest.csv files found under $resolvedPath."
}

Write-Host "Manifest files found: $($manifestFiles.Count)"

$conn = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
$conn.Open()

$totalInserted = 0
$totalSkipped = 0

try {
    foreach ($manifestFile in $manifestFiles) {
        Write-Host ""
        Write-Host "Reading $($manifestFile.FullName)" -ForegroundColor Cyan

        $rows = @(Import-Csv -LiteralPath $manifestFile.FullName)

        foreach ($row in $rows) {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = @"
IF NOT EXISTS (
    SELECT 1 FROM actt.ZipIntake
    WHERE SnapshotName = @snap AND ZipFileName = @zipname AND SHA256 = @hash
)
BEGIN
    INSERT INTO actt.ZipIntake
        (SnapshotName, ZipFileName, ZipFullPath, FileSizeBytes, ZipLastWriteTimeUtc, SHA256, HashedAtUtc)
    VALUES
        (@snap, @zipname, @zippath, @size, @lastwrite, @hash, @hashedat);
    SELECT 1 AS Inserted;
END
ELSE
BEGIN
    SELECT 0 AS Inserted;
END
"@
            [void]$cmd.Parameters.AddWithValue('@snap', [object]$row.SnapshotName)
            [void]$cmd.Parameters.AddWithValue('@zipname', $row.ZipFileName)
            [void]$cmd.Parameters.AddWithValue('@zippath', $row.ZipFullPath)
            [void]$cmd.Parameters.AddWithValue('@size', [int64]$row.FileSizeBytes)
            [void]$cmd.Parameters.AddWithValue('@lastwrite', [datetime]$row.ZipLastWriteTimeUtc)
            [void]$cmd.Parameters.AddWithValue('@hash', $row.SHA256)
            [void]$cmd.Parameters.AddWithValue('@hashedat', [datetime]$row.HashedAtUtc)

            $inserted = $cmd.ExecuteScalar()

            if ($inserted -eq 1) {
                $totalInserted++
                Write-Host "  Inserted: $($row.ZipFileName) [$($row.SHA256.Substring(0,12))...]"
            }
            else {
                $totalSkipped++
                Write-Host "  Already present, skipped: $($row.ZipFileName) [$($row.SHA256.Substring(0,12))...]"
            }
        }
    }
}
finally {
    $conn.Close()
}

Write-Host ""
Write-Host "Done. Inserted: $totalInserted   Already present (skipped): $totalSkipped" -ForegroundColor Green
