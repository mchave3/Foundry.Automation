#region Parameters
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\DriverPack\Lenovo'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CatalogXmlUri = 'https://download.lenovo.com/cdrt/td/catalogv2.xml',

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

function Get-RegexAttributeValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AttributesText,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $pattern = '(?i)\b' + [regex]::Escape($Name) + '\s*=\s*"([^"]*)"'
    $match = [regex]::Match($AttributesText, $pattern)
    if ($match.Success) {
        $value = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    return $null
}

function Get-FileNameFromUrl {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Url
    )

    if (-not $Url) {
        return $null
    }

    $withoutQuery = $Url.Split('?', 2)[0]
    $fileName = [System.IO.Path]::GetFileName($withoutQuery)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        return $null
    }

    return $fileName
}

function Get-LenovoDriverPackData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CatalogXmlUri
    )

    $response = Invoke-WebRequest -Uri $CatalogXmlUri -ErrorAction Stop
    $catalogXmlText = $response.Content

    $catalogVersionMatch = [regex]::Match($catalogXmlText, '(?is)<ModelList\b[^>]*version="([^"]*)"')
    $catalogVersion = if ($catalogVersionMatch.Success) {
        [System.Net.WebUtility]::HtmlDecode($catalogVersionMatch.Groups[1].Value).Trim()
    }
    else {
        $null
    }

    $modelNodes = [regex]::Matches($catalogXmlText, '(?is)<Model\b[^>]*name="([^"]*)"[^>]*>(.*?)</Model>')
    if ($modelNodes.Count -lt 1) {
        throw 'Invalid Lenovo catalog XML: Model nodes not found.'
    }

    $items = @()

    foreach ($model in $modelNodes) {
        $modelName = [System.Net.WebUtility]::HtmlDecode($model.Groups[1].Value).Trim()
        if (-not $modelName) {
            $modelName = $null
        }

        $modelBody = $model.Groups[2].Value
        $machineTypeMatches = [regex]::Matches($modelBody, '(?is)<Type>(.*?)</Type>')

        $machineTypes = @($machineTypeMatches |
            ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value).Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique)

        $machineTypesValue = if ($machineTypes.Count -gt 0) { $machineTypes -join ',' } else { $null }
        $machineTypeCount = if ($machineTypes.Count -gt 0) { $machineTypes.Count } else { $null }

        $sccmMatches = [regex]::Matches($modelBody, '(?is)<SCCM\b([^>]*)>(.*?)</SCCM>')
        foreach ($sccm in $sccmMatches) {
            $attributesText = $sccm.Groups[1].Value
            $downloadUrl = [System.Net.WebUtility]::HtmlDecode($sccm.Groups[2].Value)
            if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
                continue
            }

            $downloadUrl = $downloadUrl.Trim()

            if ($downloadUrl -match '(?i)winpe') {
                continue
            }

            $os = Get-RegexAttributeValue -AttributesText $attributesText -Name 'os'
            $osVersion = Get-RegexAttributeValue -AttributesText $attributesText -Name 'version'
            $releaseDate = Get-RegexAttributeValue -AttributesText $attributesText -Name 'date'
            $crc = Get-RegexAttributeValue -AttributesText $attributesText -Name 'crc'
            $md5 = Get-RegexAttributeValue -AttributesText $attributesText -Name 'md5'
            $fileName = Get-FileNameFromUrl -Url $downloadUrl

            $idModel = if ($modelName) { $modelName } else { '' }
            $idOs = if ($os) { $os } else { '' }
            $idOsVersion = if ($osVersion) { $osVersion } else { '' }
            $idFileName = if ($fileName) { $fileName } else { '' }
            $id = @($idModel, $idOs, $idOsVersion, $idFileName) -join '|'

            if ($id -eq '|||') {
                $id = Get-StringSha256 -Text $downloadUrl
            }

            $items += [pscustomobject]([ordered]@{
                    id = $id
                    model = $modelName
                    machineTypes = $machineTypesValue
                    machineTypeCount = $machineTypeCount
                    os = $os
                    osVersion = $osVersion
                    releaseDate = $releaseDate
                    downloadUrl = $downloadUrl
                    fileName = $fileName
                    crc = $crc
                    md5 = $md5
                })
        }
    }

    $metadata = [ordered]@{
        catalogVersion = $catalogVersion
        modelCount = $modelNodes.Count
    }

    return [pscustomobject]@{
        SourceHash = Get-StringSha256 -Text $catalogXmlText
        Metadata = [pscustomobject]$metadata
        Items = @($items | Sort-Object -Property @(
                @{ Expression = { $_.model } },
                @{ Expression = { $_.os } },
                @{ Expression = { $_.osVersion } },
                @{ Expression = { $_.fileName } }
            ))
    }
}

function New-LenovoCatalog {
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

function Write-LenovoCatalogXml {
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
        $writer.WriteStartElement('LenovoCatalog')
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

function Write-LenovoOutputs {
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

    $filePrefix = 'DriverPack_Lenovo'
    $jsonPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.json')
    $xmlPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.xml')
    $mdPath = Join-Path -Path $OutputDirectory -ChildPath 'README.md'

    $json = ConvertTo-DeterministicJson -Object $Catalog
    Write-Utf8NoBomFile -Path $jsonPath -Content $json
    Write-LenovoCatalogXml -Path $xmlPath -Catalog $Catalog

    $jsonHash = (Get-FileHash -Path $jsonPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $xmlHash = (Get-FileHash -Path $xmlPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $status = if ($Catalog.itemCount -gt 0) { 'SUCCESS' } else { 'WARNING' }
    $durationSeconds = [int][Math]::Round(((Get-Date) - $StartedAt).TotalSeconds)

    $summaryLines = @(
        '# Lenovo Summary',
        '',
        '| Metric | Value |',
        '| --- | --- |',
        "| Executed At (UTC) | $($Catalog.generatedAtUtc -replace 'T', ' ' -replace 'Z', ' UTC') |",
        "| Category | $($Catalog.category) |",
        "| Status | $status |",
        "| Items | $($Catalog.itemCount) |",
        "| Catalog Version | $($Catalog.catalog.catalogVersion) |",
        "| Models | $($Catalog.catalog.modelCount) |",
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

$driverPackData = Get-LenovoDriverPackData -CatalogXmlUri $CatalogXmlUri
if (@($driverPackData.Items).Count -lt $MinimumDriverPackCount) {
    throw ("Lenovo DriverPack item count below expected threshold (actual={0}, minimum={1})." -f @($driverPackData.Items).Count, $MinimumDriverPackCount)
}

$driverPackCatalog = New-LenovoCatalog -Category 'DriverPack' -Source ([ordered]@{
        name = 'Lenovo DriverPack Catalog V2'
        catalogType = 'DriverPackCatalogXml'
        uri = $CatalogXmlUri
        catalogSha256 = $driverPackData.SourceHash
    }) -CatalogMetadata ([ordered]@{
        catalogVersion = $driverPackData.Metadata.catalogVersion
        modelCount = $driverPackData.Metadata.modelCount
    }) -GeneratedAtUtc $generatedAtUtc -Items @($driverPackData.Items)

$output = Write-LenovoOutputs -OutputDirectory $DriverPackOutputDirectory -Catalog $driverPackCatalog -StartedAt $startedAt
$output

#endregion Main Execution
