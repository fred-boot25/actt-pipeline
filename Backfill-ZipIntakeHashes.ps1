<#
.SYNOPSIS
    Backfills the pre-extraction zip hash (chain-of-custody) manifest for snapshots that were
    extracted BEFORE Invoke-ACTTZipToPbiCsv.ps1 started hashing zips automatically.

.DESCRIPTION
    Does NOT re-extract, re-convert, or touch anything except the original .zip files (read-only,
    just to hash them). Safe to run against a zip folder whose contents have already been
    extracted/converted - it only fills in the missing _ACTT_ZipIntakeManifest.csv.

    Writes the manifest into the folder's _Extracted subfolder (matching where
    Invoke-ACTTZipToPbiCsv.ps1 puts it), so it can be picked up by Load-ZipIntakeToSql.ps1
    exactly the same way as a freshly-run snapshot.

.PARAMETER ZipRoot
    Folder containing the .zip files for one snapshot (e.g. "Q:\ftomaseck\ACTT_FILES\SOX_FY25").

.PARAMETER SnapshotName
    Defaults to the ZipRoot folder's leaf name, matching Invoke-ACTTZipToPbiCsv.ps1's convention.

.EXAMPLE
    .\Backfill-ZipIntakeHashes.ps1 -ZipRoot "Q:\ftomaseck\ACTT_FILES\SOX_FY25"
    .\Backfill-ZipIntakeHashes.ps1 -ZipRoot "Q:\ftomaseck\ACTT_FILES\DV_JAN_26"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ZipRoot,

    [Parameter(Mandatory = $false)]
    [string]$SnapshotName
)

$ErrorActionPreference = 'Stop'

$resolvedZipRoot = (Resolve-Path -LiteralPath $ZipRoot).Path

if ([string]::IsNullOrWhiteSpace($SnapshotName)) {
    $SnapshotName = Split-Path -Path $resolvedZipRoot -Leaf
}

$zipFiles = @(Get-ChildItem -LiteralPath $resolvedZipRoot -File -Filter '*.zip' | Sort-Object FullName)

if ($zipFiles.Count -eq 0) {
    throw "No .zip files found directly under $resolvedZipRoot."
}

$extractRoot = Join-Path $resolvedZipRoot '_Extracted'
$manifestExists = Test-Path -LiteralPath (Join-Path $extractRoot '_ACTT_ZipIntakeManifest.csv')

if ($manifestExists) {
    Write-Warning "A _ACTT_ZipIntakeManifest.csv already exists under $extractRoot - this will overwrite it with freshly computed hashes. Existing hashes should still match if the zips haven't changed."
}

Write-Host "Snapshot:  $SnapshotName"
Write-Host "Zip root:  $resolvedZipRoot"
Write-Host "Zip count: $($zipFiles.Count)"
Write-Host ""
Write-Host "Hashing (read-only - not touching extraction or conversion output)..." -ForegroundColor Cyan

$zipIntakeManifest = New-Object System.Collections.Generic.List[object]

foreach ($zip in $zipFiles) {
    $hash = Get-FileHash -LiteralPath $zip.FullName -Algorithm SHA256

    [void]$zipIntakeManifest.Add([pscustomobject]@{
        SnapshotName        = $SnapshotName
        ZipFileName         = $zip.Name
        ZipFullPath         = $zip.FullName
        FileSizeBytes       = $zip.Length
        ZipLastWriteTimeUtc = $zip.LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm:ss')
        SHA256              = $hash.Hash
        HashedAtUtc         = [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss')
    })

    Write-Host ("  {0}  [{1}]" -f $zip.Name, $hash.Hash)
}

if (-not (Test-Path -LiteralPath $extractRoot)) {
    New-Item -ItemType Directory -Path $extractRoot -Force | Out-Null
}

$manifestPath = Join-Path $extractRoot '_ACTT_ZipIntakeManifest.csv'
$zipIntakeManifest | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Done. Manifest written to: $manifestPath" -ForegroundColor Green
Write-Host "Note: this hash reflects the zip's CURRENT state on disk, not necessarily the moment it was originally received - flag that distinction if this matters for evidentiary purposes." -ForegroundColor Yellow
Write-Host "Next: .\Load-ZipIntakeToSql.ps1 -ManifestPath `"$manifestPath`""
