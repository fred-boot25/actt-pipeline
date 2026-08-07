<#
.SYNOPSIS
    Converts ACTT .actt/.txt output folders to clean CSV files for Power BI ingestion.

.DESCRIPTION
    - Processes one or more ACTT server folders under a root directory.
    - Reads ACTT_CONFIG_SETTINGS.actt for audit/server metadata.
    - Reads ACTT_CONFIG_FIELDTERMINATOR.actt when available; defaults to *#*.
    - Converts each .actt/.txt data file to CSV.
    - Handles multiline ACTT records, especially MSDB_SYSJOBSTEPS command text.
    - Adds traceability columns: SnapshotName, ServerFolder, ServerName, AuditStartDate,
      AuditEndDate, DataExtractionDate, SourceFile, SourceRowNumber, SourceSHA256.
    - Renames database_scoped to DatabaseName for DB-scoped ACTT files.
    - Creates a file manifest with SHA256 hashes.
    - Creates a conversion summary and parse issue file.

.EXAMPLE
    $root = "C:\FY26-ACTT\0116 to 0119 PBC R2 ACTT"
    $out  = "$root\CsvOut"

    $folders = @(
        "BBCIDB_ACTT_MSSQL",
        "BBOMSPRODDB1_ACTT_MSSQL",
        "BBSQLPrimary_ACTT_MSSQL",
        "BBSSIS01_ACTT_MSSQL",
        "BBGPSQL_ACTT_MSSQL"
    )

    & ".\Convert-ACTTFoldersToCsv_v3.ps1" `
        -Root $root `
        -OutputRoot $out `
        -Folders $folders `
        -SnapshotName "FY26_PBC_R2" `
        -Overwrite
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Root,

    [Parameter(Mandatory = $true)]
    [string]$OutputRoot,

    [Parameter(Mandatory = $false)]
    [string[]]$Folders,

    [Parameter(Mandatory = $false)]
    [string]$SnapshotName,

    [Parameter(Mandatory = $false)]
    [string]$DefaultDelimiter = '*#*',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLogs,

    [Parameter(Mandatory = $false)]
    [switch]$Overwrite
)

$ErrorActionPreference = 'Stop'

function Get-SafeFileName {
    param([Parameter(Mandatory = $true)][string]$Name)

    $safe = $Name
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($char, '_')
    }
    return $safe
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$FullPath
    )

    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\') + '\'
    $full = [System.IO.Path]::GetFullPath($FullPath)

    if ($full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length)
    }

    return $full
}

function Convert-HeaderTokenToName {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Token,

        [Parameter(Mandatory = $true)]
        [int]$Ordinal
    )

    if ([string]::IsNullOrWhiteSpace($Token)) {
        return "Column_$Ordinal"
    }

    $name = $Token.Trim()

    # Strip common ACTT SQL data type suffixes from header definitions.
    $name = $name -replace '\s+(SYSNAME|VARCHAR|NVARCHAR|CHAR|NCHAR|INT|BIGINT|SMALLINT|TINYINT|BIT|DATETIME|DATETIME2|DATE|TIME|DECIMAL|NUMERIC|FLOAT|REAL|MONEY|SMALLMONEY|UNIQUEIDENTIFIER|VARBINARY|BINARY|TEXT|NTEXT)\s*(\(.*?\))?\s*$', ''
    $name = $name.Trim()

    if ([string]::IsNullOrWhiteSpace($name)) {
        return "Column_$Ordinal"
    }

    if ($name -ieq 'database_scoped') {
        return 'DatabaseName'
    }

    return $name
}

function Get-UniqueColumnNames {
    param([Parameter(Mandatory = $true)][string[]]$Names)

    $seen = @{}
    $output = New-Object System.Collections.Generic.List[string]

    foreach ($name in $Names) {
        $base = if ([string]::IsNullOrWhiteSpace($name)) { 'Column' } else { $name }
        $candidate = $base
        $i = 2

        while ($seen.ContainsKey($candidate.ToUpperInvariant())) {
            $candidate = "{0}_{1}" -f $base, $i
            $i++
        }

        $seen[$candidate.ToUpperInvariant()] = $true
        [void]$output.Add($candidate)
    }

    return $output.ToArray()
}

function Get-ActtDelimiter {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $false)][string]$FallbackDelimiter = '*#*'
    )

    $terminatorPath = Join-Path $Folder 'ACTT_CONFIG_FIELDTERMINATOR.actt'

    if (-not (Test-Path -LiteralPath $terminatorPath)) {
        return $FallbackDelimiter
    }

    $lines = @(Get-Content -LiteralPath $terminatorPath -ErrorAction SilentlyContinue)

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($line -match '\*#\*') {
            return '*#*'
        }
    }

    return $FallbackDelimiter
}

function Get-EffectiveDelimiterFromHeader {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HeaderLine,

        [Parameter(Mandatory = $false)]
        [string]$ConfiguredDelimiter = '*#*'
    )

    $delimiterCandidates = @(
        '*#*',
        '|',
        "`t",
        ','
    )

    foreach ($candidateDelimiter in $delimiterCandidates) {
        if ($HeaderLine.Contains($candidateDelimiter)) {
            $testParts = @(Split-ActtLine -Line $HeaderLine -Delimiter $candidateDelimiter)

            if ($testParts.Count -gt 1) {
                return $candidateDelimiter
            }
        }
    }

    return $ConfiguredDelimiter
}

function Split-ActtLine {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Line,

        [Parameter(Mandatory = $true)]
        [string]$Delimiter
    )

    if ($null -eq $Line) {
        return @()
    }

    return @($Line.Split([string[]]@($Delimiter), [System.StringSplitOptions]::None))
}

function Get-ActtConfigSettings {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $false)][string]$Delimiter = '*#*'
    )

    $settingsPath = Join-Path $Folder 'ACTT_CONFIG_SETTINGS.actt'
    $settings = @{}

    if (-not (Test-Path -LiteralPath $settingsPath)) {
        return $settings
    }

    $lines = @(Get-Content -LiteralPath $settingsPath -ErrorAction SilentlyContinue)

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $effectiveDelimiter = Get-EffectiveDelimiterFromHeader -HeaderLine $line -ConfiguredDelimiter $Delimiter
        $parts = @(Split-ActtLine -Line $line -Delimiter $effectiveDelimiter)

        if ($parts.Count -gt 2) {
            $parts = @($parts[0], (($parts[1..($parts.Count - 1)]) -join $effectiveDelimiter))
        }

        if ($parts.Count -ne 2) {
            continue
        }

        $key = Convert-HeaderTokenToName -Token $parts[0] -Ordinal 1
        $value = $parts[1].Trim()

        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        if ($key -match '^(SettingName|Column_1)$') {
            continue
        }

        $settings[$key] = $value
    }

    return $settings
}

function Get-SettingValue {
    param(
        [hashtable]$Settings,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        if ($Settings.ContainsKey($name)) {
            return $Settings[$name]
        }
    }

    return $null
}

function Convert-DateValue {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $dt = [datetime]::MinValue

    if ([datetime]::TryParse($Value, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd')
    }

    return $Value
}

function Get-ActtLogicalRecords {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string[]]$Lines,

        [Parameter(Mandatory = $true)]
        [string]$Delimiter,

        [Parameter(Mandatory = $true)]
        [int]$ExpectedColumnCount,

        [Parameter(Mandatory = $false)]
        [int]$FirstDataLineNumber = 2
    )

    if ($null -eq $Lines) {
        return @()
    }

    $Lines = @($Lines)

    $records = New-Object System.Collections.Generic.List[object]
    $buffer = $null
    $startLineNumber = $null

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        $physicalLineNumber = $FirstDataLineNumber + $i

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($null -eq $buffer) {
            $buffer = $line
            $startLineNumber = $physicalLineNumber
        }
        else {
            $buffer += "`r`n" + $line
        }

        $parts = @(Split-ActtLine -Line $buffer -Delimiter $Delimiter)

        if ($parts.Count -ge $ExpectedColumnCount) {
            [void]$records.Add([pscustomobject]@{
                SourceRowNumber = $startLineNumber
                RawRecord       = $buffer
            })

            $buffer = $null
            $startLineNumber = $null
        }
    }

    if ($null -ne $buffer) {
        [void]$records.Add([pscustomobject]@{
            SourceRowNumber = $startLineNumber
            RawRecord       = $buffer
        })
    }

    return $records.ToArray()
}

function Read-ActtFileAsObjects {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$Delimiter,

        [Parameter(Mandatory = $true)]
        [hashtable]$Metadata,

        [Parameter(Mandatory = $false)]
        [string]$SourceSHA256
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $parseIssues = New-Object System.Collections.Generic.List[object]

    $lines = @(Get-Content -LiteralPath $File.FullName -ErrorAction Stop)

    if ($lines.Count -eq 0) {
        return [pscustomobject]@{
            Rows        = @()
            ParseIssues = @()
            HeaderCount = 0
        }
    }

    # Remove trailing blank lines only. Do not remove blanks inside multiline SQL/job commands.
    while ($lines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) {
        if ($lines.Count -eq 1) {
            $lines = @()
        }
        else {
            $lines = @($lines[0..($lines.Count - 2)])
        }
    }

    if ($lines.Count -eq 0) {
        return [pscustomobject]@{
            Rows        = @()
            ParseIssues = @()
            HeaderCount = 0
        }
    }

    $sourceFile = $File.Name

    # Header is first physical line.
    # Some ACTT config files can report the wrong delimiter, so detect from the actual file header.
    $effectiveDelimiter = Get-EffectiveDelimiterFromHeader -HeaderLine $lines[0] -ConfiguredDelimiter $Delimiter
    $headerParts = @(Split-ActtLine -Line $lines[0] -Delimiter $effectiveDelimiter)

    $headerList = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $headerParts.Count; $i++) {
        $headerName = Convert-HeaderTokenToName -Token $headerParts[$i] -Ordinal ($i + 1)

        if ($headerName -ieq 'database_scoped') {
            $headerName = 'DatabaseName'
        }

        [void]$headerList.Add($headerName)
    }

    $headers = @(Get-UniqueColumnNames -Names $headerList.ToArray())
    $headerCount = $headers.Count

    if ($headerCount -eq 0) {
        return [pscustomobject]@{
            Rows        = @()
            ParseIssues = @()
            HeaderCount = 0
        }
    }

    # Data begins after header.
    $dataLines = @()

    if ($lines.Count -gt 1) {
        $dataLines = @($lines[1..($lines.Count - 1)])
    }

    if ($dataLines.Count -eq 0) {
        return [pscustomobject]@{
            Rows        = @()
            ParseIssues = @()
            HeaderCount = $headerCount
        }
    }

    # Rebuild logical records. This handles MSDB_SYSJOBSTEPS multiline command text.
    $logicalRecords = @(Get-ActtLogicalRecords `
        -Lines $dataLines `
        -Delimiter $effectiveDelimiter `
        -ExpectedColumnCount $headerCount `
        -FirstDataLineNumber 2)

    foreach ($record in $logicalRecords) {
        if ($null -eq $record -or [string]::IsNullOrWhiteSpace($record.RawRecord)) {
            continue
        }

        $values = @(Split-ActtLine -Line $record.RawRecord -Delimiter $effectiveDelimiter)

        # Skip blank/delimiter-only logical records.
        $nonBlankValues = @($values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        if ($nonBlankValues.Count -eq 0) {
            continue
        }

        $row = [ordered]@{
            SnapshotName       = $Metadata.SnapshotName
            ServerFolder       = $Metadata.ServerFolder
            ServerName         = $Metadata.ServerName
            AuditStartDate     = $Metadata.AuditStartDate
            AuditEndDate       = $Metadata.AuditEndDate
            DataExtractionDate = $Metadata.DataExtractionDate
            SourceFile         = $sourceFile
            SourceRowNumber    = $record.SourceRowNumber
            SourceSHA256       = $SourceSHA256
        }

        for ($i = 0; $i -lt $headerCount; $i++) {
            $value = $null

            if ($i -lt $values.Count) {
                $value = $values[$i]
            }

            $row[$headers[$i]] = $value
        }

        if ($values.Count -ne $headerCount) {
            $issueText = "Expected $headerCount columns; found $($values.Count)."
            $row['_ParseIssue'] = $issueText

            [void]$parseIssues.Add([pscustomobject]@{
                SnapshotName    = $Metadata.SnapshotName
                ServerFolder    = $Metadata.ServerFolder
                ServerName      = $Metadata.ServerName
                SourceFile      = $sourceFile
                SourceRowNumber = $record.SourceRowNumber
                Issue           = $issueText
                RawLine         = $record.RawRecord
            })
        }
        else {
            $row['_ParseIssue'] = $null
        }

        [void]$rows.Add([pscustomobject]$row)
    }

    return [pscustomobject]@{
        Rows        = $rows.ToArray()
        ParseIssues = $parseIssues.ToArray()
        HeaderCount = $headerCount
    }
}

function Test-IsActtFolder {
    param([Parameter(Mandatory = $true)][string]$Folder)

    if (Test-Path -LiteralPath (Join-Path $Folder 'ACTT_CONFIG_SETTINGS.actt')) {
        return $true
    }

    if (Test-Path -LiteralPath (Join-Path $Folder 'ACTT_CONFIG_TABLERECORDCOUNT.txt')) {
        return $true
    }

    $actt = @(Get-ChildItem -LiteralPath $Folder -File -Filter '*.actt' -ErrorAction SilentlyContinue)
    return ($actt.Count -gt 0)
}

$resolvedRoot = (Resolve-Path -LiteralPath $Root).Path

New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$resolvedOutputRoot = (Resolve-Path -LiteralPath $OutputRoot).Path

if ([string]::IsNullOrWhiteSpace($SnapshotName)) {
    $SnapshotName = Split-Path -Path $resolvedRoot -Leaf
}

if (-not $Folders -or $Folders.Count -eq 0) {
    $Folders = @(Get-ChildItem -LiteralPath $resolvedRoot -Directory -ErrorAction Stop |
        Where-Object { Test-IsActtFolder -Folder $_.FullName } |
        Sort-Object Name |
        Select-Object -ExpandProperty Name)
}

if (-not $Folders -or $Folders.Count -eq 0) {
    throw "No ACTT folders were supplied or discovered under $resolvedRoot."
}

$manifest = New-Object System.Collections.Generic.List[object]
$summary = New-Object System.Collections.Generic.List[object]
$allParseIssues = New-Object System.Collections.Generic.List[object]

foreach ($folderName in $Folders) {
    $inputFolder = Join-Path $resolvedRoot $folderName

    if (-not (Test-Path -LiteralPath $inputFolder)) {
        Write-Warning "Input folder not found: $inputFolder"
        continue
    }

    $serverOut = Join-Path $resolvedOutputRoot (Get-SafeFileName -Name $folderName)
    New-Item -ItemType Directory -Path $serverOut -Force | Out-Null

    $delimiter = Get-ActtDelimiter -Folder $inputFolder -FallbackDelimiter $DefaultDelimiter
    $settings = Get-ActtConfigSettings -Folder $inputFolder -Delimiter $delimiter

    $serverName = Get-SettingValue -Settings $settings -Names @(
        'SQLServer Instance',
        'SQL Server Instance',
        'ServerName',
        'Server Name'
    )

    if ([string]::IsNullOrWhiteSpace($serverName)) {
        $serverName = $folderName -replace '_ACTT_MSSQL$', ''
    }

    $metadata = @{
        SnapshotName       = $SnapshotName
        ServerFolder       = $folderName
        ServerName         = $serverName
        AuditStartDate     = Convert-DateValue (Get-SettingValue -Settings $settings -Names @('Audit Start Date', 'AuditStartDate'))
        AuditEndDate       = Convert-DateValue (Get-SettingValue -Settings $settings -Names @('Audit End Date', 'AuditEndDate'))
        DataExtractionDate = Convert-DateValue (Get-SettingValue -Settings $settings -Names @('Data Extraction Date', 'DataExtractionDate'))
    }

    Write-Host ""
    Write-Host "Processing $folderName" -ForegroundColor Cyan
    Write-Host "  Server:    $($metadata.ServerName)"
    Write-Host "  Audit:     $($metadata.AuditStartDate) to $($metadata.AuditEndDate)"
    Write-Host "  Delimiter: $delimiter"

    $files = @(Get-ChildItem -LiteralPath $inputFolder -File -ErrorAction Stop |
        Where-Object {
            $_.Extension -ieq '.actt' -or
            $_.Name -ieq 'ACTT_CONFIG_TABLERECORDCOUNT.txt' -or
            ($IncludeLogs -and $_.Extension -ieq '.log')
        } |
        Sort-Object Name)

    foreach ($file in $files) {
        $hash = Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256
        $relativePath = Get-RelativePath -BasePath $resolvedRoot -FullPath $file.FullName

        [void]$manifest.Add([pscustomobject]@{
            SnapshotName = $SnapshotName
            ServerFolder = $folderName
            ServerName   = $metadata.ServerName
            RelativePath  = $relativePath
            FileName      = $file.Name
            Length        = $file.Length
            LastWriteTime = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            SHA256        = $hash.Hash
        })

        # LOG files are evidence/control only for now. Put them in manifest, do not parse as tabular data.
        if ($file.Extension -ieq '.log') {
            [void]$summary.Add([pscustomobject]@{
                SnapshotName    = $SnapshotName
                ServerFolder    = $folderName
                ServerName      = $metadata.ServerName
                SourceFile      = $file.Name
                CsvFile         = $null
                RowCount        = $null
                ColumnCount     = $null
                ParseIssueCount = $null
                Status          = 'ManifestOnlyLogFile'
                Error           = $null
            })
            continue
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $csvName = (Get-SafeFileName -Name $baseName) + '.csv'
        $csvPath = Join-Path $serverOut $csvName

        if ((Test-Path -LiteralPath $csvPath) -and -not $Overwrite) {
            Write-Host "  Skipping existing $csvName"
            continue
        }

        try {
            $parsed = Read-ActtFileAsObjects `
                -File $file `
                -Delimiter $delimiter `
                -Metadata $metadata `
                -SourceSHA256 $hash.Hash

            $rows = @($parsed.Rows)
            $issues = @($parsed.ParseIssues)

            if ($rows.Count -gt 0) {
                $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
            }
            else {
                [pscustomobject]@{
                    SnapshotName = $SnapshotName
                    ServerFolder = $folderName
                    ServerName   = $metadata.ServerName
                    SourceFile   = $file.Name
                    SourceSHA256 = $hash.Hash
                    Note         = 'No data rows found'
                } | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
            }

            foreach ($issue in $issues) {
                [void]$allParseIssues.Add($issue)
            }

            [void]$summary.Add([pscustomobject]@{
                SnapshotName    = $SnapshotName
                ServerFolder    = $folderName
                ServerName      = $metadata.ServerName
                SourceFile      = $file.Name
                CsvFile         = $csvPath
                RowCount        = $rows.Count
                ColumnCount     = $parsed.HeaderCount
                ParseIssueCount = $issues.Count
                Status          = 'Converted'
                Error           = $null
            })

            Write-Host ("  {0} -> {1} rows, {2} parse issues" -f $file.Name, $rows.Count, $issues.Count)
        }
        catch {
            Write-Warning "Failed converting $($file.FullName): $($_.Exception.Message)"

            [void]$summary.Add([pscustomobject]@{
                SnapshotName    = $SnapshotName
                ServerFolder    = $folderName
                ServerName      = $metadata.ServerName
                SourceFile      = $file.Name
                CsvFile         = $csvPath
                RowCount        = 0
                ColumnCount     = $null
                ParseIssueCount = -1
                Status          = 'Failed'
                Error           = $_.Exception.Message
            })
        }
    }
}

$manifestPath = Join-Path $resolvedOutputRoot '_ACTT_FileManifest.csv'
$summaryPath = Join-Path $resolvedOutputRoot '_ACTT_Csv_Conversion_Summary.csv'
$parseIssuePath = Join-Path $resolvedOutputRoot '_ACTT_ParseIssues.csv'

$manifest | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding UTF8
$summary | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

if ($allParseIssues.Count -gt 0) {
    $allParseIssues | Export-Csv -LiteralPath $parseIssuePath -NoTypeInformation -Encoding UTF8
}
else {
    [pscustomobject]@{
        SnapshotName = $SnapshotName
        Note         = 'No parse issues found'
    } | Export-Csv -LiteralPath $parseIssuePath -NoTypeInformation -Encoding UTF8
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Manifest:     $manifestPath"
Write-Host "Summary:      $summaryPath"
Write-Host "Parse issues: $parseIssuePath"