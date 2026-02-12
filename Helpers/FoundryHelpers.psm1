#region File Operations

# Write text file as UTF-8 without BOM.
function Write-Utf8NoBomFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $encoding = [System.Text.UTF8Encoding]::new($false)
    $normalized = ($Content -replace "`r?`n", "`r`n").TrimEnd("`r", "`n") + "`r`n"
    [System.IO.File]::WriteAllText($Path, $normalized, $encoding)
}

#endregion File Operations

#region Path Resolution

# Resolve a writable temporary root directory in a cross-platform way.
function Get-TemporaryRootPath {
    $tempPath = [System.IO.Path]::GetTempPath()
    if (-not $tempPath) {
        return '/tmp'
    }

    return $tempPath
}

# Resolve 7-Zip executable path for cross-platform CAB extraction.
function Get-SevenZipCommandPath {
    $candidates = @('7zz', '7z')
    foreach ($candidate in $candidates) {
        $command = Get-Command -Name $candidate -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) {
            return [string]$command.Source
        }
    }

    throw 'Required tool not found: 7-Zip CLI (7zz or 7z). Install 7-Zip/p7zip before running this script.'
}

#endregion Path Resolution

#region Archive Operations

# Extract file patterns from an archive using 7-Zip.
function Invoke-SevenZipExtract {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SevenZipPath,

        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )

    $arguments = @('e', '-y', "-o$OutputDirectory", $ArchivePath)
    $arguments += $Patterns
    & $SevenZipPath @arguments | Out-Null
}

#endregion Archive Operations

#region Type Conversions

# Return safe integer conversion for XML numeric attributes.
function ConvertTo-IntOrNull {
    param(
        [Parameter()]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if (-not $Value) {
        return $null
    }

    [int]$parsed = 0
    if ([int]::TryParse($Value, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

# Return safe int64 conversion for XML numeric attributes.
function ConvertTo-Int64OrNull {
    param(
        [Parameter()]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if (-not $Value) {
        return $null
    }

    [long]$parsed = 0
    if ([long]::TryParse($Value, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

#endregion Type Conversions

#region XML Utilities

# Return text content from XML elements including CDATA payloads.
function Get-XmlInnerText {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Node
    )

    if ($null -eq $Node) {
        return $null
    }

    if ($Node -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Node)) {
            return $null
        }

        return $Node.Trim()
    }

    if ($Node.PSObject.Properties.Name -contains 'InnerText') {
        $innerText = [string]$Node.InnerText
        if (-not [string]::IsNullOrWhiteSpace($innerText)) {
            return $innerText.Trim()
        }
    }

    return $null
}

#endregion XML Utilities

#region Catalog Output Helpers

# Create a standardized README.md summary for catalog outputs.
function Write-CatalogReadme {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Manufacturer,

        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$GeneratedAtUtc,

        [Parameter(Mandatory = $true)]
        [int]$ItemCount,

        [Parameter(Mandatory = $true)]
        [string]$CatalogVersion,

        [Parameter(Mandatory = $true)]
        [int]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [string]$XmlHash,

        [Parameter()]
        [hashtable]$AdditionalMetrics = @{}
    )

    $status = if ($ItemCount -gt 0) { 'SUCCESS' } else { 'WARNING' }
    $generatedAtFormatted = $GeneratedAtUtc -replace 'T', ' ' -replace 'Z', ' UTC'

    $summaryLines = @(
        "# $Manufacturer Summary",
        '',
        '| Metric | Value |',
        '| --- | --- |',
        "| Executed At (UTC) | $generatedAtFormatted |",
        "| Category | $Category |",
        "| Status | $status |",
        "| Items | $ItemCount |",
        "| Catalog Version | $CatalogVersion |",
        "| Duration (Seconds) | $DurationSeconds |",
        "| SHA256 XML | $XmlHash |"
    )

    foreach ($metric in $AdditionalMetrics.GetEnumerator()) {
        $summaryLines += "| $($metric.Key) | $($metric.Value) |"
    }

    Write-Utf8NoBomFile -Path $Path -Content ($summaryLines -join "`r`n")
}

# Create XmlWriter with standard settings for catalog files.
function New-CatalogXmlWriter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $true
    $settings.IndentChars = '  '
    $settings.NewLineChars = "`r`n"
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)

    return [System.Xml.XmlWriter]::Create($Path, $settings)
}

#endregion Catalog Output Helpers
