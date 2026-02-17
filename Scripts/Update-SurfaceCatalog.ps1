#region Parameters
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\DriverPack\Surface'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackArticleUri = 'https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120',

    [Parameter()]
    [ValidateRange(0, 1000000)]
    [int]$MinimumDriverPackCount = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
#endregion Parameters

#region Utility Functions

function ConvertTo-DeterministicJson {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object
    )

    $json = ConvertTo-Json -InputObject $Object -Depth 10
    return ($json -replace "`r?`n", "`r`n").TrimEnd("`r", "`n")
}

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

function Get-StringSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    return [System.Convert]::ToHexString(
        [System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($Text))
    ).ToLowerInvariant()
}

function Invoke-WebRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$OperationName = 'web request',

        [Parameter()]
        [ValidateRange(0, 10)]
        [int]$RetryCount = 3,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]$RetryDelaySeconds = 2,

        [Parameter()]
        [ValidateRange(5, 300)]
        [int]$TimeoutSeconds = 45
    )

    $attempt = 0
    $maxAttempts = $RetryCount + 1
    while ($attempt -lt $maxAttempts) {
        $attempt++
        try {
            return Invoke-WebRequest -Uri $Uri -TimeoutSec $TimeoutSeconds -ErrorAction Stop
        }
        catch {
            if ($attempt -ge $maxAttempts) {
                throw
            }

            Write-Warning ("{0} failed for '{1}' (attempt {2}/{3}): {4}" -f $OperationName, $Uri, $attempt, $maxAttempts, $_.Exception.Message)
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Get-HtmlMetaContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $pattern = '(?is)<meta\s+name="' + [regex]::Escape($Name) + '"\s+content="([^"]*)"'
    $match = [regex]::Match($Html, $pattern)
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return $null
}

function Get-DriverPackItemModel {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Title
    )

    if (-not $Title) {
        return $null
    }

    $clean = $Title
    $clean = [System.Net.WebUtility]::HtmlDecode($clean)
    $clean = $clean -replace '(?i)^\s*Download\s+', ''
    $clean = $clean -replace '(?i)\s+from\s+Official\s+Microsoft\s+Download\s+Center\s*$', ''
    $clean = $clean -replace '(?i)\s*-\s*Official Microsoft Download Center\s*$', ''
    $clean = $clean -replace '(?i)\s*Download\s*$', ''
    $clean = $clean -replace '(?i)^Drivers?\s+and\s+Firmware\s+for\s+.+?\s+on\s+(Surface .+)$', '$1'
    $clean = $clean -replace '(?i)^(Surface .+?)\s+Drivers?\s+and\s+Firmware$', '$1'
    $clean = $clean -replace '(?i)^(Surface .+?)\s+Firmware\s+and\s+Drivers$', '$1'
    $clean = $clean -replace '(?i)^(Surface .+?)\s+Driver\s+and\s+Firmware$', '$1'
    $clean = $clean.Trim()
    if (-not $clean) {
        return $null
    }

    return $clean
}

function ConvertTo-Int64OrNull {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    [int64]$parsed = 0
    if ([int64]::TryParse(([string]$Value), [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function ConvertTo-SurfaceDateOrNull {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Value
    )

    if (-not $Value) {
        return $null
    }

    [datetimeoffset]$parsedOffset = [datetimeoffset]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces
    if ([datetimeoffset]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsedOffset)) {
        return $parsedOffset.ToUniversalTime().ToString('yyyy-MM-dd')
    }
    if ([datetimeoffset]::TryParse($Value, [System.Globalization.CultureInfo]::CurrentCulture, $styles, [ref]$parsedOffset)) {
        return $parsedOffset.ToUniversalTime().ToString('yyyy-MM-dd')
    }

    [datetime]$parsedDate = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsedDate)) {
        return $parsedDate.ToString('yyyy-MM-dd')
    }
    if ([datetime]::TryParse($Value, [System.Globalization.CultureInfo]::CurrentCulture, $styles, [ref]$parsedDate)) {
        return $parsedDate.ToString('yyyy-MM-dd')
    }

    return $null
}

function Get-UriFileName {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Uri
    )

    if (-not $Uri) {
        return $null
    }

    try {
        return [System.IO.Path]::GetFileName(([System.Uri]$Uri).LocalPath)
    }
    catch {
        return [System.IO.Path]::GetFileName($Uri)
    }
}

function Get-SurfaceDlcDetails {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html
    )

    $marker = 'window.__DLCDetails__'
    $markerIndex = $Html.IndexOf($marker, [System.StringComparison]::Ordinal)
    if ($markerIndex -lt 0) {
        return $null
    }

    $equalsIndex = $Html.IndexOf('=', $markerIndex)
    if ($equalsIndex -lt 0) {
        return $null
    }

    $jsonStart = $Html.IndexOf('{', $equalsIndex)
    if ($jsonStart -lt 0) {
        return $null
    }

    $depth = 0
    $inString = $false
    $escape = $false
    $jsonEnd = -1

    for ($position = $jsonStart; $position -lt $Html.Length; $position++) {
        $char = $Html[$position]

        if ($inString) {
            if ($escape) {
                $escape = $false
                continue
            }

            if ($char -eq '\') {
                $escape = $true
                continue
            }

            if ($char -eq '"') {
                $inString = $false
            }

            continue
        }

        if ($char -eq '"') {
            $inString = $true
            continue
        }

        if ($char -eq '{') {
            $depth++
            continue
        }

        if ($char -eq '}') {
            $depth--
            if ($depth -eq 0) {
                $jsonEnd = $position
                break
            }
        }
    }

    if ($jsonEnd -lt $jsonStart) {
        return $null
    }

    $json = $Html.Substring($jsonStart, ($jsonEnd - $jsonStart + 1))
    try {
        return ($json | ConvertFrom-Json -Depth 20 -ErrorAction Stop)
    }
    catch {
        Write-Verbose ("Unable to parse Surface DLC details JSON: {0}" -f $_.Exception.Message)
        return $null
    }
}

function Get-SurfaceDriverPackItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArticleUri
    )

    $articleResponse = Invoke-WebRequestWithRetry -Uri $ArticleUri -OperationName 'Surface support article request'
    $articleHtml = $articleResponse.Content

    $downloadCenterMatches = [regex]::Matches(
        $articleHtml,
        'https://www\.microsoft\.com/(?:[a-z]{2}-[a-z]{2}/)?download/details\.aspx\?id=(\d+)',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $downloadCenterLinks = @($downloadCenterMatches | ForEach-Object { $_.Value } | Sort-Object -Unique)

    if ($downloadCenterLinks.Count -eq 0) {
        throw ("No Microsoft Download Center links were extracted from Surface support article '{0}'." -f $ArticleUri)
    }

    $items = @()
    foreach ($downloadCenterUrl in $downloadCenterLinks) {
        $detailHtml = $null
        try {
            $detailResponse = Invoke-WebRequestWithRetry -Uri $downloadCenterUrl -OperationName 'Surface package detail request'
            $detailHtml = $detailResponse.Content
        }
        catch {
            Write-Warning ("Skipping Surface package detail page due to repeated request failure: {0}. Error: {1}" -f $downloadCenterUrl, $_.Exception.Message)
            continue
        }

        $titleMatch = [regex]::Match($detailHtml, '(?is)<title>(.*?)</title>')
        $title = if ($titleMatch.Success) { [System.Net.WebUtility]::HtmlDecode($titleMatch.Groups[1].Value).Trim() } else { $null }

        $packageIdMatch = [regex]::Match($downloadCenterUrl, '(?i)id=(\d+)')
        $packageId = if ($packageIdMatch.Success) { $packageIdMatch.Groups[1].Value } else { $null }

        $dlcDetails = Get-SurfaceDlcDetails -Html $detailHtml
        $detailView = if ($dlcDetails -and ($dlcDetails.PSObject.Properties.Name -contains 'dlcDetailsView')) { $dlcDetails.dlcDetailsView } else { $null }

        $downloadTitle = if ($detailView -and [string]$detailView.downloadTitle) { [string]$detailView.downloadTitle } else { $null }
        $model = $null
        if ($downloadTitle) {
            $model = Get-DriverPackItemModel -Title $downloadTitle
        }
        if (-not $model) {
            $model = Get-DriverPackItemModel -Title $title
        }

        $supportedOperatingSystems = if ($detailView -and [string]$detailView.systemRequirementsSection_supportedOS) { [string]$detailView.systemRequirementsSection_supportedOS } else { $null }

        $downloadFiles = @()
        if ($detailView -and ($detailView.PSObject.Properties.Name -contains 'downloadFile')) {
            foreach ($downloadFile in @($detailView.downloadFile)) {
                $downloadUrl = if ([string]$downloadFile.url) { [string]$downloadFile.url } else { $null }
                if (-not $downloadUrl) {
                    continue
                }

                $fileName = if ([string]$downloadFile.name) { [string]$downloadFile.name } else { Get-UriFileName -Uri $downloadUrl }
                $extension = if ($fileName) { [System.IO.Path]::GetExtension($fileName) } else { [System.IO.Path]::GetExtension((Get-UriFileName -Uri $downloadUrl)) }
                $format = if ($extension) { $extension.TrimStart('.').ToLowerInvariant() } else { $null }

                $downloadFiles += [pscustomobject]([ordered]@{
                        downloadUrl = $downloadUrl
                        fileName = $fileName
                        format = $format
                        version = if ([string]$downloadFile.version) { [string]$downloadFile.version } else { $null }
                        sizeBytes = ConvertTo-Int64OrNull -Value $downloadFile.size
                        datePublished = ConvertTo-SurfaceDateOrNull -Value ([string]$downloadFile.datePublished)
                    })
            }
        }

        if ($downloadFiles.Count -lt 1) {
            $msiMatches = [regex]::Matches(
                $detailHtml,
                'https://download\.microsoft\.com/download/[^"\s<>]+?\.msi',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
            foreach ($msiUrl in @($msiMatches | ForEach-Object { $_.Value } | Sort-Object -Unique)) {
                $downloadFiles += [pscustomobject]([ordered]@{
                        downloadUrl = $msiUrl
                        fileName = Get-UriFileName -Uri $msiUrl
                        format = 'msi'
                        version = $null
                        sizeBytes = $null
                        datePublished = $null
                    })
            }
        }

        if ($downloadFiles.Count -gt 0) {
            foreach ($downloadFile in $downloadFiles) {
                $itemFileName = if ([string]$downloadFile.fileName) { [string]$downloadFile.fileName } else { Get-UriFileName -Uri ([string]$downloadFile.downloadUrl) }
                $id = if ($packageId -and $itemFileName) { ($packageId + '|' + $itemFileName) } elseif ($packageId) { $packageId } else { [string]$downloadFile.downloadUrl }

                $items += [pscustomobject]([ordered]@{
                        id = $id
                        model = $model
                        downloadTitle = $downloadTitle
                        packageId = $packageId
                        downloadCenterUrl = $downloadCenterUrl
                        downloadUrl = [string]$downloadFile.downloadUrl
                        msiUrl = if ([string]$downloadFile.format -eq 'msi') { [string]$downloadFile.downloadUrl } else { $null }
                        fileName = $itemFileName
                        format = [string]$downloadFile.format
                        version = if ([string]$downloadFile.version) { [string]$downloadFile.version } else { $null }
                        sizeBytes = ConvertTo-Int64OrNull -Value $downloadFile.sizeBytes
                        datePublished = if ([string]$downloadFile.datePublished) { [string]$downloadFile.datePublished } else { $null }
                        supportedOperatingSystems = $supportedOperatingSystems
                        device = $null
                        importFolderCount = $null
                        importFolders = $null
                        hidMiniZipUrl = $null
                        guidanceUrl = $null
                    })
            }
            continue
        }

        $id = if ($packageId) { $packageId } else { $downloadCenterUrl }
        $items += [pscustomobject]([ordered]@{
                id = $id
                model = $model
                downloadTitle = $downloadTitle
                packageId = $packageId
                downloadCenterUrl = $downloadCenterUrl
                downloadUrl = $null
                msiUrl = $null
                fileName = $null
                format = $null
                version = $null
                sizeBytes = $null
                datePublished = $null
                supportedOperatingSystems = $supportedOperatingSystems
                device = $null
                importFolderCount = $null
                importFolders = $null
                hidMiniZipUrl = $null
                guidanceUrl = $null
            })
    }

    $metadata = [ordered]@{
        articleGuid = Get-HtmlMetaContent -Html $articleHtml -Name 'awa-articleGuid'
        firstPublishedDate = Get-HtmlMetaContent -Html $articleHtml -Name 'firstPublishedDate'
        lastPublishedDate = Get-HtmlMetaContent -Html $articleHtml -Name 'lastPublishedDate'
    }

    return [pscustomobject]@{
        SourceHash = Get-StringSha256 -Text $articleHtml
        Metadata = [pscustomobject]$metadata
        Items = @($items | Sort-Object -Property @(
                @{ Expression = { $_.model } },
                @{ Expression = { $_.packageId } },
                @{ Expression = { $_.fileName } }
            ))
    }
}

function New-SurfaceCatalog {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('DriverPack')]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [hashtable]$Source,

        [Parameter(Mandatory = $true)]
        [hashtable]$CatalogMetadata,

        [Parameter(Mandatory = $true)]
        [string]$GeneratedAtUtc,

        [Parameter(Mandatory = $true)]
        [object[]]$Items
    )

    return [ordered]@{
        schemaVersion = 1
        generatedAtUtc = $GeneratedAtUtc
        category = $Category
        source = $Source
        catalog = $CatalogMetadata
        itemCount = $Items.Count
        items = $Items
    }
}

function Write-SurfaceCatalogXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [hashtable]$Catalog
    )

    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $true
    $settings.IndentChars = '  '
    $settings.NewLineChars = "`r`n"
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)

    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('SurfaceCatalog')
        $writer.WriteAttributeString('schemaVersion', [string]$Catalog.schemaVersion)
        $writer.WriteAttributeString('generatedAtUtc', [string]$Catalog.generatedAtUtc)
        $writer.WriteAttributeString('category', [string]$Catalog.category)

        $writer.WriteStartElement('Source')
        foreach ($property in $Catalog.source.GetEnumerator()) {
            if ($null -eq $property.Value) {
                continue
            }

            $writer.WriteAttributeString([string]$property.Key, [string]$property.Value)
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('CatalogMetadata')
        foreach ($property in $Catalog.catalog.GetEnumerator()) {
            if ($null -eq $property.Value) {
                continue
            }

            $writer.WriteAttributeString([string]$property.Key, [string]$property.Value)
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('Items')
        foreach ($item in $Catalog.items) {
            $writer.WriteStartElement('Item')
            foreach ($property in $item.PSObject.Properties) {
                if ($null -eq $property.Value) {
                    continue
                }

                $writer.WriteElementString([string]$property.Name, [string]$property.Value)
            }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-SurfaceCategoryOutputs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [hashtable]$Catalog,

        [Parameter(Mandatory = $true)]
        [datetime]$StartedAt
    )

    if (-not (Test-Path -Path $OutputDirectory)) {
        $null = New-Item -Path $OutputDirectory -ItemType Directory -Force
    }

    $filePrefix = 'DriverPack_Surface'
    $xmlPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.xml')
    $mdPath = Join-Path -Path $OutputDirectory -ChildPath 'README.md'

        Write-SurfaceCatalogXml -Path $xmlPath -Catalog $Catalog

    $xmlHash = (Get-FileHash -Path $xmlPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $status = if ($Catalog.itemCount -gt 0) { 'SUCCESS' } else { 'WARNING' }
    $durationSeconds = [int][Math]::Round(((Get-Date) - $StartedAt).TotalSeconds)

    $summaryLines = @(
        '# Surface Summary',
        '',
        '| Metric | Value |',
        '| --- | --- |',
        "| Executed At (UTC) | $($Catalog.generatedAtUtc -replace 'T', ' ' -replace 'Z', ' UTC') |",
        "| Category | $($Catalog.category) |",
        "| Status | $status |",
        "| Items | $($Catalog.itemCount) |",
        "| Last Published | $($Catalog.catalog.lastPublishedDate) |",
        "| Duration (Seconds) | $durationSeconds |",
        "| SHA256 XML | $xmlHash |"
    )

    Write-Utf8NoBomFile -Path $mdPath -Content ($summaryLines -join "`r`n")

    return [pscustomobject]@{
        Category = $Catalog.category
        XmlPath = $xmlPath
        MarkdownPath = $mdPath
        Items = $Catalog.itemCount
        Status = $status
    }
}

#endregion Surface-Specific Functions

#region Main Execution

$startedAt = Get-Date
$generatedAtUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

$driverPackData = Get-SurfaceDriverPackItems -ArticleUri $DriverPackArticleUri
if (@($driverPackData.Items).Count -lt $MinimumDriverPackCount) {
    throw ("Surface DriverPack item count below expected threshold (actual={0}, minimum={1})." -f @($driverPackData.Items).Count, $MinimumDriverPackCount)
}

$driverPackCatalog = New-SurfaceCatalog -Category 'DriverPack' -Source ([ordered]@{
        name = 'Microsoft Surface Drivers and Firmware'
        catalogType = 'SupportArticle'
        uri = $DriverPackArticleUri
        catalogSha256 = $driverPackData.SourceHash
    }) -CatalogMetadata ([ordered]@{
        articleGuid = $driverPackData.Metadata.articleGuid
        firstPublishedDate = $driverPackData.Metadata.firstPublishedDate
        lastPublishedDate = $driverPackData.Metadata.lastPublishedDate
    }) -GeneratedAtUtc $generatedAtUtc -Items @($driverPackData.Items)
$driverPackOutput = Write-SurfaceCategoryOutputs -OutputDirectory $DriverPackOutputDirectory -Catalog $driverPackCatalog -StartedAt $startedAt
$driverPackOutput

#endregion Main Execution
