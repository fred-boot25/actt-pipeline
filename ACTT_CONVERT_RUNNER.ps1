<#
.SYNOPSIS
    End-to-end ACTT pipeline: extract zips, hash for chain-of-custody, convert to CSV,
    load zip hashes to SQL, load converted data to SQL.

.PARAMETER SnapshotName
    Required. Must match the name of the folder containing this run's zip files under
    -ZipBaseRoot (e.g. "SOX_FY26_20260324" -> Q:\ftomaseck\ACTT_FILES\SOX_FY26_20260324).
    Convention: {Base}_{yyyyMMdd} for real audit rounds, {Base}_TEST_{yyyyMMdd} for
    pipeline-validation runs that aren't real audit data.

.PARAMETER ZipBaseRoot
    Parent folder containing per-snapshot zip folders. Defaults to the standard location.

.PARAMETER OutputRoot
    Parent folder for converted CSV output. Defaults to the standard PBI_READ_SOX location.

.PARAMETER AllowNewColumns
    Passed through to Load-ACTTCsvToSql.ps1. Leave off (default) unless you've already
    reviewed a schema-drift warning from a prior run and want new columns auto-added.

.PARAMETER SkipSqlLoad
    Runs extraction/conversion only (step 1), skips steps 2-3. Useful for TEST runs where
    you just want to validate extraction/conversion without touching SQL.

.EXAMPLE
    .\ACTT_CONVERT_RUNNER.ps1 -SnapshotName "SOX_FY26_20260324"

.EXAMPLE
    .\ACTT_CONVERT_RUNNER.ps1 -SnapshotName "SOX_FY26_TEST_20260324" -SkipSqlLoad
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SnapshotName,

    [Parameter(Mandatory = $false)]
    [string]$ZipBaseRoot = "Q:\ftomaseck\ACTT_FILES",

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "Q:\ftomaseck\ACTT_FILES_SOX\PBI_READ_SOX",

    [Parameter(Mandatory = $false)]
    [switch]$AllowNewColumns,

    [Parameter(Mandatory = $false)]
    [switch]$SkipSqlLoad
)

Set-Location C:\ACTT_COMMAND
$ErrorActionPreference = 'Stop'

$ZipRoot = Join-Path $ZipBaseRoot $SnapshotName
$snapshotOutputFolder = Join-Path $OutputRoot $SnapshotName

if (-not (Test-Path -LiteralPath $ZipRoot)) {
    throw "ZipRoot does not exist: $ZipRoot`nCheck that -SnapshotName matches an actual folder under -ZipBaseRoot ($ZipBaseRoot)."
}

try {
    Write-Host "=== STEP 1/3: Extract zips, hash for chain-of-custody, convert to CSV ===" -ForegroundColor Magenta
    Write-Host "Snapshot: $SnapshotName"
    Write-Host "Zip root: $ZipRoot"
    Write-Host ""

    .\Invoke-ACTTZipToPbiCsv.ps1 `
        -ZipRoot $ZipRoot `
        -OutputRoot $OutputRoot `
        -SnapshotName $SnapshotName `
        -ConverterPath "C:\ACTT_COMMAND\Convert-ACTTFoldersToCsv_v2.ps1" `
        -CleanExtract `
        -CleanOutputSnapshot `
        -OverwriteCsv

    if ($SkipSqlLoad) {
        Write-Host ""
        Write-Host "SkipSqlLoad set - stopping after extraction/conversion. No SQL changes made." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host "=== STEP 2/3: Load zip chain-of-custody hashes to SQL ===" -ForegroundColor Magenta
    .\Load-ZipIntakeToSql.ps1 -ManifestPath $snapshotOutputFolder

    Write-Host ""
    Write-Host "=== STEP 3/3: Load converted CSVs to SQL ===" -ForegroundColor Magenta
    if ($AllowNewColumns) {
        .\Load-ACTTCsvToSql.ps1 -CsvRoot $snapshotOutputFolder -AllowNewColumns
    }
    else {
        .\Load-ACTTCsvToSql.ps1 -CsvRoot $snapshotOutputFolder
    }

    Write-Host ""
    Write-Host "ACTT pipeline complete for snapshot: $SnapshotName" -ForegroundColor Green
}
catch {
    Write-Error "ACTT pipeline FAILED for snapshot $SnapshotName : $($_.Exception.Message)"
    throw
}
