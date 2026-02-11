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
    $clean = $clean -replace '(?i)\s*-\s*Official Microsoft Download Center\s*$', ''
    $clean = $clean -replace '(?i)\s*Download\s*$', ''
    $clean = $clean.Trim()
    if (-not $clean) {
        return $null
    }

    return $clean
}

function Get-SurfaceDriverPackItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArticleUri
    )

    $articleResponse = Invoke-WebRequest -Uri $ArticleUri -ErrorAction Stop
    $articleHtml = $articleResponse.Content

    $downloadCenterMatches = [regex]::Matches(
        $articleHtml,
        'https://www\.microsoft\.com/download/details\.aspx\?id=(\d+)',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $downloadCenterLinks = @($downloadCenterMatches | ForEach-Object { $_.Value } | Sort-Object -Unique)

    $items = @()
    foreach ($downloadCenterUrl in $downloadCenterLinks) {
        $detailResponse = Invoke-WebRequest -Uri $downloadCenterUrl -ErrorAction Stop
        $detailHtml = $detailResponse.Content

        $titleMatch = [regex]::Match($detailHtml, '(?is)<title>(.*?)</title>')
        $title = if ($titleMatch.Success) { [System.Net.WebUtility]::HtmlDecode($titleMatch.Groups[1].Value).Trim() } else { $null }
        $model = Get-DriverPackItemModel -Title $title

        $packageIdMatch = [regex]::Match($downloadCenterUrl, '(?i)id=(\d+)')
        $packageId = if ($packageIdMatch.Success) { $packageIdMatch.Groups[1].Value } else { $null }

        $msiMatches = [regex]::Matches(
            $detailHtml,
            'https://download\.microsoft\.com/download/[^"\s<>]+?\.msi',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        $msiUrls = @($msiMatches | ForEach-Object { $_.Value } | Sort-Object -Unique)

        if ($msiUrls.Count -gt 0) {
            foreach ($msiUrl in $msiUrls) {
                $fileName = [System.IO.Path]::GetFileName($msiUrl)
                $id = ($packageId + '|' + $fileName)

                $items += [pscustomobject]([ordered]@{
                        id = $id
                        model = $model
                        packageId = $packageId
                        downloadCenterUrl = $downloadCenterUrl
                        msiUrl = $msiUrl
                        fileName = $fileName
                        device = $null
                        importFolderCount = $null
                        importFolders = $null
                        hidMiniZipUrl = $null
                        guidanceUrl = $null
                    })
            }
        }
        else {
            $id = if ($packageId) { $packageId } else { $downloadCenterUrl }
            $items += [pscustomobject]([ordered]@{
                    id = $id
                    model = $model
                    packageId = $packageId
                    downloadCenterUrl = $downloadCenterUrl
                    msiUrl = $null
                    fileName = $null
                    device = $null
                    importFolderCount = $null
                    importFolders = $null
                    hidMiniZipUrl = $null
                    guidanceUrl = $null
                })
        }
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
    $jsonPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.json')
    $xmlPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.xml')
    $mdPath = Join-Path -Path $OutputDirectory -ChildPath 'README.md'

    $json = ConvertTo-DeterministicJson -Object $Catalog
    Write-Utf8NoBomFile -Path $jsonPath -Content $json
    Write-SurfaceCatalogXml -Path $xmlPath -Catalog $Catalog

    $jsonHash = (Get-FileHash -Path $jsonPath -Algorithm SHA256).Hash.ToLowerInvariant()
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
        "| SHA256 JSON | $jsonHash |",
        "| SHA256 XML | $xmlHash |"
    )

    Write-Utf8NoBomFile -Path $mdPath -Content ($summaryLines -join "`r`n")

    return [pscustomobject]@{
        Category = $Catalog.category
        JsonPath = $jsonPath
        XmlPath = $xmlPath
        MarkdownPath = $mdPath
        Items = $Catalog.itemCount
        Status = $status
    }
}

#endregion Utility Functions

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
