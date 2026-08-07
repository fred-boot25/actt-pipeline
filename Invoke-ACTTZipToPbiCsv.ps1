<#
.SYNOPSIS
    Extracts ACTT zip files, normalizes the extracted ACTT server folders, and converts them to CSV for Power BI folder ingestion.

.DESCRIPTION
    - Takes a folder containing ACTT .zip files, for example Q:\ftomaseck\ACTT_FILES\DV_JAN_26.
    - Extracts each zip into an extraction working folder.
    - Finds ACTT server folders whether the zip contains files directly or wraps them in a subfolder.
    - Builds a normalized input root using directory junctions when possible, with Copy-Item fallback.
    - Builds a snapshot output folder under the stable Power BI parent folder.
    - Calls Convert-ACTTFoldersToCsv_v2.ps1 once so each snapshot has one Power BI-friendly CSV folder structure.

.EXAMPLE
    .\Invoke-ACTTZipToPbiCsv.ps1 `
        -ZipRoot "Q:\ftomaseck\ACTT_FILES\DV_JAN_26" `
        -OutputRoot "Q:\ftomaseck\ACTT_FILES\PBI_READ" `
        -SnapshotName "DV_JAN_26" `
        -CleanExtract `
        -CleanOutputSnapshot `
        -OverwriteCsv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ZipRoot,

    [Parameter(Mandatory = $true)]
    [string]$OutputRoot,

    [Parameter(Mandatory = $false)]
    [string]$ExtractRoot,

    [Parameter(Mandatory = $false)]
    [string]$SnapshotName,

    [Parameter(Mandatory = $false)]
    [string]$ConverterPath,

    [Parameter(Mandatory = $false)]
    [switch]$RecurseZips,

    [Parameter(Mandatory = $false)]
    [switch]$CleanExtract,

    [Parameter(Mandatory = $false)]
    [switch]$OverwriteCsv,

    [Parameter(Mandatory = $false)]
    [switch]$CleanOutputSnapshot,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLogs
)

$ErrorActionPreference = 'Stop'

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $safe = $Name
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($char, '_')
    }

    return $safe
}

function Test-IsActtFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Folder
    )

    if (Test-Path -LiteralPath (Join-Path $Folder 'ACTT_CONFIG_SETTINGS.actt')) {
        return $true
    }

    if (Test-Path -LiteralPath (Join-Path $Folder 'ACTT_CONFIG_TABLERECORDCOUNT.txt')) {
        return $true
    }

    $actt = @(Get-ChildItem -LiteralPath $Folder -File -Filter '*.actt' -ErrorAction SilentlyContinue)
    return ($actt.Count -gt 0)
}

function New-NormalizedActtFolderLink {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceFolder,

        [Parameter(Mandatory = $true)]
        [string]$NormalizedRoot,

        [Parameter(Mandatory = $true)]
        [string]$PreferredName
    )

    $baseName = Get-SafeFileName -Name $PreferredName
    $targetName = $baseName
    $i = 2

    while (Test-Path -LiteralPath (Join-Path $NormalizedRoot $targetName)) {
        $targetName = '{0}_{1}' -f $baseName, $i
        $i++
    }

    $linkPath = Join-Path $NormalizedRoot $targetName

    try {
        New-Item -ItemType Junction -Path $linkPath -Target $SourceFolder -ErrorAction Stop | Out-Null

        return [pscustomobject]@{
            FolderName   = $targetName
            SourceFolder = $SourceFolder
            LinkPath     = $linkPath
            Method       = 'Junction'
        }
    }
    catch {
        Copy-Item -LiteralPath $SourceFolder -Destination $linkPath -Recurse -Force -ErrorAction Stop

        return [pscustomobject]@{
            FolderName   = $targetName
            SourceFolder = $SourceFolder
            LinkPath     = $linkPath
            Method       = 'CopyFallback'
        }
    }
}

$resolvedZipRoot = (Resolve-Path -LiteralPath $ZipRoot).Path

if ([string]::IsNullOrWhiteSpace($ExtractRoot)) {
    $ExtractRoot = Join-Path $resolvedZipRoot '_Extracted'
}

if ([string]::IsNullOrWhiteSpace($SnapshotName)) {
    $SnapshotName = Split-Path -Path $resolvedZipRoot -Leaf
}

$safeSnapshotName = Get-SafeFileName -Name $SnapshotName

$resolvedOutputBaseRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputRoot)
$outputSnapshotRoot = Join-Path $resolvedOutputBaseRoot $safeSnapshotName

if ([string]::IsNullOrWhiteSpace($ConverterPath)) {
    $ConverterPath = Join-Path $PSScriptRoot 'Convert-ACTTFoldersToCsv_v2.ps1'
}

$resolvedConverterPath = (Resolve-Path -LiteralPath $ConverterPath).Path

if ($CleanExtract -and (Test-Path -LiteralPath $ExtractRoot)) {
    Remove-Item -LiteralPath $ExtractRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $ExtractRoot -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedOutputBaseRoot -Force | Out-Null

if ($CleanOutputSnapshot -and (Test-Path -LiteralPath $outputSnapshotRoot)) {
    Remove-Item -LiteralPath $outputSnapshotRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $outputSnapshotRoot -Force | Out-Null

$zipSearchParams = @{
    LiteralPath = $resolvedZipRoot
    File        = $true
    Filter      = '*.zip'
}

if ($RecurseZips) {
    $zipSearchParams.Recurse = $true
}

$zipFiles = @(Get-ChildItem @zipSearchParams | Sort-Object FullName)

if ($zipFiles.Count -eq 0) {
    throw "No .zip files found under $resolvedZipRoot."
}

Write-Host "Zip root:     $resolvedZipRoot" -ForegroundColor Cyan
Write-Host "Extract root: $ExtractRoot"
Write-Host "PBI root:     $resolvedOutputBaseRoot"
Write-Host "Snapshot:     $SnapshotName"
Write-Host "Snapshot out: $outputSnapshotRoot"
Write-Host "Converter:    $resolvedConverterPath"
Write-Host "Zip count:    $($zipFiles.Count)"
Write-Host ""

# --- Zip chain-of-custody: hash every zip BEFORE it is touched/expanded. ---
# This is the record of "this is exactly what we received", independent of anything
# the extraction or conversion steps do afterward.
$zipIntakeManifest = New-Object System.Collections.Generic.List[object]

Write-Host "Hashing zip files (chain of custody, pre-extraction)..." -ForegroundColor Cyan

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

$zipIntakeManifestPath = Join-Path $ExtractRoot '_ACTT_ZipIntakeManifest.csv'
$zipIntakeManifest | Export-Csv -LiteralPath $zipIntakeManifestPath -NoTypeInformation -Encoding UTF8

# Also drop a copy alongside the snapshot's other manifests, so it travels with the
# converted output and evidence for this snapshot lives in one place.
Copy-Item -LiteralPath $zipIntakeManifestPath -Destination (Join-Path $outputSnapshotRoot '_ACTT_ZipIntakeManifest.csv') -Force

Write-Host ""

foreach ($zip in $zipFiles) {
    $extractName = Get-SafeFileName -Name ([System.IO.Path]::GetFileNameWithoutExtension($zip.Name))
    $zipExtractRoot = Join-Path $ExtractRoot $extractName

    if ($CleanExtract -and (Test-Path -LiteralPath $zipExtractRoot)) {
        Remove-Item -LiteralPath $zipExtractRoot -Recurse -Force
    }

    New-Item -ItemType Directory -Path $zipExtractRoot -Force | Out-Null

    Write-Host "Extracting $($zip.Name) -> $zipExtractRoot"
    Expand-Archive -LiteralPath $zip.FullName -DestinationPath $zipExtractRoot -Force
}

$normalizedRoot = Join-Path $ExtractRoot '_ACTT_Normalized_Input'

if (Test-Path -LiteralPath $normalizedRoot) {
    Remove-Item -LiteralPath $normalizedRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $normalizedRoot -Force | Out-Null

$acttFolders = @(
    Get-ChildItem -LiteralPath $ExtractRoot -Directory -Recurse -ErrorAction Stop |
        Where-Object {
            $_.FullName -notlike "$normalizedRoot*" -and
            (Test-IsActtFolder -Folder $_.FullName)
        } |
        Sort-Object FullName
)

if ($acttFolders.Count -eq 0) {
    throw "Zip files were extracted, but no ACTT folders were discovered under $ExtractRoot."
}

Write-Host ""
Write-Host "Discovered ACTT folders: $($acttFolders.Count)" -ForegroundColor Cyan

$normalizedFolders = New-Object System.Collections.Generic.List[object]

foreach ($folder in $acttFolders) {
    $normalized = New-NormalizedActtFolderLink `
        -SourceFolder $folder.FullName `
        -NormalizedRoot $normalizedRoot `
        -PreferredName $folder.Name

    [void]$normalizedFolders.Add($normalized)

    Write-Host ("  {0} <= {1} [{2}]" -f $normalized.FolderName, $normalized.SourceFolder, $normalized.Method)
}

$normalizedManifestPath = Join-Path $ExtractRoot '_ACTT_Normalized_Input_Manifest.csv'
$normalizedFolders | Export-Csv -LiteralPath $normalizedManifestPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Running ACTT converter..." -ForegroundColor Cyan

$converterParams = @{
    Root       = $normalizedRoot
    OutputRoot = $outputSnapshotRoot
}

if (-not [string]::IsNullOrWhiteSpace($SnapshotName)) {
    $converterParams.SnapshotName = $SnapshotName
}

if ($OverwriteCsv) {
    $converterParams.Overwrite = $true
}

if ($IncludeLogs) {
    $converterParams.IncludeLogs = $true
}

& $resolvedConverterPath @converterParams

if ($LASTEXITCODE -ne $null -and $LASTEXITCODE -ne 0) {
    throw "Converter returned exit code $LASTEXITCODE."
}

Write-Host ""
Write-Host "ACTT zip-to-CSV conversion complete." -ForegroundColor Green
Write-Host "Power BI folder source: $resolvedOutputBaseRoot"
Write-Host "Snapshot output:         $outputSnapshotRoot"
Write-Host "Normalized input manifest: $normalizedManifestPath"
Write-Host "Zip intake manifest (pre-extraction hashes): $zipIntakeManifestPath"