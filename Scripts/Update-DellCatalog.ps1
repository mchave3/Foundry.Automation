#region Parameters
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\DriverPack\Dell'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$WinPEOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\WinPE\Dell'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CatalogCabUri = 'https://downloads.dell.com/catalog/DriverPackCatalog.cab',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CatalogXmlFileName = 'DriverPackCatalog.xml',

    [Parameter()]
    [ValidateRange(0, 200000)]
    [int]$MinimumDriverPackCount = 1,

    [Parameter()]
    [ValidateRange(0, 200000)]
    [int]$MinimumWinPECount = 1

)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
#endregion Parameters

#region Import Helpers

$helpersPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath @('Helpers', 'FoundryHelpers.psm1')
if (Test-Path -Path $helpersPath) {
    Import-Module -Name $helpersPath -Force -ErrorAction Stop
}
else {
    throw "Helpers module not found at: $helpersPath"
}

#endregion Import Helpers

#region Dell-Specific Functions

# Build absolute download URL from base location and package relative path.
function Join-DellDownloadUrl {
    param(
        [Parameter()]
        [AllowEmptyString()]
        [string]$BaseLocation,

        [Parameter()]
        [AllowEmptyString()]
        [string]$RelativePath
    )

    if (-not $RelativePath) {
        return $null
    }

    if ($RelativePath -match '^(?i)(?:https?|ftps?)://') {
        return $RelativePath
    }

    if (-not $BaseLocation) {
        return $null
    }

    $normalizedBase = $BaseLocation
    if ($normalizedBase -notmatch '^(?i)[a-z][a-z0-9+.-]*://') {
        $normalizedBase = 'https://' + $normalizedBase
    }

    return $normalizedBase.TrimEnd('/') + '/' + $RelativePath.TrimStart('/')
}

# Parse DriverPackCatalog.xml into normalized package objects.
function ConvertFrom-DellDriverPackCatalogXml {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$CatalogXml
    )

    $manifest = $CatalogXml.DriverPackManifest
    if (-not $manifest) {
        throw 'Invalid Dell catalog XML: DriverPackManifest node not found.'
    }

    $metadata = [ordered]@{
        version = if ([string]$manifest.version) { [string]$manifest.version } else { $null }
        schemaVersion = if ([string]$manifest.schemaVersion) { [string]$manifest.schemaVersion } else { $null }
        baseLocation = if ([string]$manifest.baseLocation) { [string]$manifest.baseLocation } else { $null }
        baseLocationAccessProtocols = if ([string]$manifest.baseLocationAccessProtocols) { [string]$manifest.baseLocationAccessProtocols } else { $null }
        dateTime = if ([string]$manifest.dateTime) { [string]$manifest.dateTime } else { $null }
    }

    $items = foreach ($package in @($manifest.DriverPackage)) {
        if (-not $package) {
            continue
        }

        $rawType = [string]$package.type
        $normalizedType = switch (($rawType).ToLowerInvariant()) {
            'win' { 'Win'; break }
            'winpe' { 'WinPE'; break }
            default { $rawType }
        }

        $relativePath = [string]$package.path
        if (-not $relativePath) {
            continue
        }

        $supportedSystems = @()
        $brandNodes = @()
        if ($package.SupportedSystems -and ($package.SupportedSystems.PSObject.Properties.Name -contains 'Brand')) {
            $brandNodes = @($package.SupportedSystems.Brand)
        }

        foreach ($brand in $brandNodes) {
            $brandKey = if ([string]$brand.key) { [string]$brand.key } else { $null }
            $brandPrefix = if ([string]$brand.prefix) { [string]$brand.prefix } else { $null }

            $models = @($brand.Model)
            if ($models.Count -lt 1) {
                $supportedSystems += [pscustomobject]([ordered]@{
                        brandKey = $brandKey
                        brandPrefix = $brandPrefix
                        systemId = $null
                        systemName = $null
                        generation = $null
                        rtsDate = $null
                    })
                continue
            }

            foreach ($model in $models) {
                $supportedSystems += [pscustomobject]([ordered]@{
                        brandKey = $brandKey
                        brandPrefix = $brandPrefix
                        systemId = if ([string]$model.systemID) { [string]$model.systemID } else { $null }
                        systemName = if ([string]$model.name) { [string]$model.name } else { $null }
                        generation = if ([string]$model.generation) { [string]$model.generation } else { $null }
                        rtsDate = if ([string]$model.rtsDate) { [string]$model.rtsDate } else { $null }
                    })
            }
        }

        $osNodes = @()
        if ($package.SupportedOperatingSystems -and ($package.SupportedOperatingSystems.PSObject.Properties.Name -contains 'OperatingSystem')) {
            $osNodes = @($package.SupportedOperatingSystems.OperatingSystem)
        }

        $supportedOperatingSystems = @(
            foreach ($os in $osNodes) {
                [pscustomobject]([ordered]@{
                        osCode = if ([string]$os.osCode) { [string]$os.osCode } else { $null }
                        osVendor = if ([string]$os.osVendor) { [string]$os.osVendor } else { $null }
                        osArch = if ([string]$os.osArch) { [string]$os.osArch } else { $null }
                        osType = if ([string]$os.osType) { [string]$os.osType } else { $null }
                        majorVersion = ConvertTo-IntOrNull -Value ([string]$os.majorVersion)
                        minorVersion = ConvertTo-IntOrNull -Value ([string]$os.minorVersion)
                        spMajorVersion = ConvertTo-IntOrNull -Value ([string]$os.spMajorVersion)
                        spMinorVersion = ConvertTo-IntOrNull -Value ([string]$os.spMinorVersion)
                    })
            }
        )

        [pscustomobject]([ordered]@{
                releaseId = if ([string]$package.releaseID) { [string]$package.releaseID } else { [System.IO.Path]::GetFileNameWithoutExtension($relativePath) }
                type = $normalizedType
                name = Get-XmlInnerText -Node $package.Name
                description = Get-XmlInnerText -Node $package.Description
                importantInfo = Get-XmlInnerText -Node $package.ImportantInfo
                format = if ([string]$package.format) { [string]$package.format } else { $null }
                vendorVersion = if ([string]$package.vendorVersion) { [string]$package.vendorVersion } else { $null }
                dellVersion = if ([string]$package.dellVersion) { [string]$package.dellVersion } else { $null }
                dateTime = if ([string]$package.dateTime) { [string]$package.dateTime } else { $null }
                hashMD5 = if ([string]$package.hashMD5) { [string]$package.hashMD5 } else { $null }
                sizeBytes = ConvertTo-Int64OrNull -Value ([string]$package.size)
                path = $relativePath
                downloadUrl = Join-DellDownloadUrl -BaseLocation ([string]$manifest.baseLocation) -RelativePath $relativePath
                supportedSystems = $supportedSystems
                supportedOperatingSystems = $supportedOperatingSystems
            })
    }

    $itemsSorted = @($items | Sort-Object -Property @(
            @{ Expression = { $_.type } },
            @{ Expression = { $_.releaseId } },
            @{ Expression = { $_.dellVersion } },
            @{ Expression = { $_.path } }
        ))

    return [pscustomobject]@{
        Metadata = [pscustomobject]$metadata
        Items = $itemsSorted
    }
}

# Create catalog object for one category output.
function New-DellCatalog {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('DriverPack', 'WinPE')]
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

# Emit XML output aligned with the JSON catalog shape.
function Write-DellCatalogXml {
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
        $writer.WriteStartElement('DellCatalog')
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

            foreach ($field in @('releaseId', 'type', 'name', 'description', 'importantInfo', 'format', 'vendorVersion', 'dellVersion', 'dateTime', 'hashMD5', 'sizeBytes', 'path', 'downloadUrl')) {
                $value = $item.$field
                if ($null -eq $value) {
                    continue
                }

                $writer.WriteElementString($field, [string]$value)
            }

            $writer.WriteStartElement('SupportedSystems')
            foreach ($system in @($item.supportedSystems)) {
                $writer.WriteStartElement('System')
                foreach ($field in @('brandKey', 'brandPrefix', 'systemId', 'systemName', 'generation', 'rtsDate')) {
                    $value = $system.$field
                    if ($null -eq $value) {
                        continue
                    }

                    $writer.WriteAttributeString($field, [string]$value)
                }
                $writer.WriteEndElement()
            }
            $writer.WriteEndElement()

            $writer.WriteStartElement('SupportedOperatingSystems')
            foreach ($operatingSystem in @($item.supportedOperatingSystems)) {
                $writer.WriteStartElement('OperatingSystem')
                foreach ($field in @('osCode', 'osVendor', 'osArch', 'osType', 'majorVersion', 'minorVersion', 'spMajorVersion', 'spMinorVersion')) {
                    $value = $operatingSystem.$field
                    if ($null -eq $value) {
                        continue
                    }

                    $writer.WriteAttributeString($field, [string]$value)
                }
                $writer.WriteEndElement()
            }
            $writer.WriteEndElement()

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

# Write XML/README outputs for one category folder.
function Write-DellCategoryOutputs {
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

    $filePrefix = if ($Catalog.category -eq 'DriverPack') { 'DriverPack_Dell' } else { 'WinPE_Dell' }
    $xmlPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.xml')
    $mdPath = Join-Path -Path $OutputDirectory -ChildPath 'README.md'

    Write-DellCatalogXml -Path $xmlPath -Catalog $Catalog
    $xmlHash = (Get-FileHash -Path $xmlPath -Algorithm SHA256).Hash.ToLowerInvariant()

    $status = if ($Catalog.itemCount -gt 0) { 'SUCCESS' } else { 'WARNING' }
    $durationSeconds = [int][Math]::Round(((Get-Date) - $StartedAt).TotalSeconds)

    Write-CatalogReadme -Path $mdPath `
        -Manufacturer 'Dell' `
        -Category $Catalog.category `
        -GeneratedAtUtc $Catalog.generatedAtUtc `
        -ItemCount $Catalog.itemCount `
        -CatalogVersion $Catalog.catalog.version `
        -DurationSeconds $durationSeconds `
        -XmlHash $xmlHash

    return [pscustomobject]@{
        Category = $Catalog.category
        XmlPath = $xmlPath
        MarkdownPath = $mdPath
        Items = $Catalog.itemCount
        Status = $status
    }
}

#endregion Dell-Specific Functions

#region Main Execution

$startedAt = Get-Date
$generatedAtUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

$sevenZipPath = Get-SevenZipCommandPath
$tempRoot = Join-Path -Path (Get-TemporaryRootPath) -ChildPath ('foundry-dell-catalog-' + [guid]::NewGuid())
$null = New-Item -Path $tempRoot -ItemType Directory -Force

try {
    $catalogCabPath = Join-Path -Path $tempRoot -ChildPath 'DriverPackCatalog.cab'
    $catalogXmlPath = Join-Path -Path $tempRoot -ChildPath $CatalogXmlFileName

    Invoke-WebRequest -Uri $CatalogCabUri -OutFile $catalogCabPath -ErrorAction Stop
    $catalogCabSha256 = (Get-FileHash -Path $catalogCabPath -Algorithm SHA256).Hash.ToLowerInvariant()

    try {
        Invoke-SevenZipExtract -SevenZipPath $sevenZipPath -ArchivePath $catalogCabPath -OutputDirectory $tempRoot -Patterns @($CatalogXmlFileName)
    }
    catch {
        Write-Verbose -Message ("7-Zip direct extraction failed for '{0}': {1}" -f $CatalogXmlFileName, $_.Exception.Message)
    }

    if (-not (Test-Path -Path $catalogXmlPath)) {
        try {
            Invoke-SevenZipExtract -SevenZipPath $sevenZipPath -ArchivePath $catalogCabPath -OutputDirectory $tempRoot -Patterns @('*.xml')
        }
        catch {
            Write-Verbose -Message ("7-Zip wildcard extraction failed for Dell catalog: {0}" -f $_.Exception.Message)
        }

        $xmlCandidates = @(Get-ChildItem -Path $tempRoot -Filter '*.xml' -File | Sort-Object -Descending -Property LastWriteTimeUtc, Name)
        if ($xmlCandidates.Count -ge 1) {
            Copy-Item -Path $xmlCandidates[0].FullName -Destination $catalogXmlPath -Force
        }
    }

    if (-not (Test-Path -Path $catalogXmlPath)) {
        throw "Catalog XML '$CatalogXmlFileName' not found after CAB extraction."
    }

    $catalogXmlSha256 = (Get-FileHash -Path $catalogXmlPath -Algorithm SHA256).Hash.ToLowerInvariant()
    [xml]$catalogXml = Get-Content -Path $catalogXmlPath -Raw

    $parsedCatalog = ConvertFrom-DellDriverPackCatalogXml -CatalogXml $catalogXml
    $driverPackItems = @($parsedCatalog.Items | Where-Object { $_.type -eq 'Win' })
    $winPEItems = @($parsedCatalog.Items | Where-Object { $_.type -eq 'WinPE' })

    if ($driverPackItems.Count -lt $MinimumDriverPackCount) {
        throw ("DriverPack item count below expected threshold (actual={0}, minimum={1})." -f $driverPackItems.Count, $MinimumDriverPackCount)
    }

    if ($winPEItems.Count -lt $MinimumWinPECount) {
        throw ("WinPE item count below expected threshold (actual={0}, minimum={1})." -f $winPEItems.Count, $MinimumWinPECount)
    }

    $source = [ordered]@{
        name = 'Dell Driver Pack Catalog'
        catalogCabUri = $CatalogCabUri
        catalogXmlFile = [System.IO.Path]::GetFileName($catalogXmlPath)
        catalogCabSha256 = $catalogCabSha256
        catalogXmlSha256 = $catalogXmlSha256
    }

    $catalogMetadata = [ordered]@{
        version = $parsedCatalog.Metadata.version
        schemaVersion = $parsedCatalog.Metadata.schemaVersion
        baseLocation = $parsedCatalog.Metadata.baseLocation
        baseLocationAccessProtocols = $parsedCatalog.Metadata.baseLocationAccessProtocols
        dateTime = $parsedCatalog.Metadata.dateTime
    }

    $driverPackCatalog = New-DellCatalog -Category 'DriverPack' -Source $source -CatalogMetadata $catalogMetadata -GeneratedAtUtc $generatedAtUtc -Items $driverPackItems
    $winPECatalog = New-DellCatalog -Category 'WinPE' -Source $source -CatalogMetadata $catalogMetadata -GeneratedAtUtc $generatedAtUtc -Items $winPEItems

    $driverPackOutput = Write-DellCategoryOutputs -OutputDirectory $DriverPackOutputDirectory -Catalog $driverPackCatalog -StartedAt $startedAt
    $winPEOutput = Write-DellCategoryOutputs -OutputDirectory $WinPEOutputDirectory -Catalog $winPECatalog -StartedAt $startedAt

    $result = [pscustomobject]@{
        DriverPack = $driverPackOutput
        WinPE = $winPEOutput
    }

    return $result
}
finally {
    Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

#endregion Main Execution
