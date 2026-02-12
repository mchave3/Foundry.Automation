#region Parameters
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackOutputPath = (Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath @('Cache', 'DriverPack', 'DriverPack_Unified.xml')),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$WinPEOutputPath = (Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath @('Cache', 'WinPE', 'WinPE_Unified.xml')),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CacheDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath @('Cache')),

    [Parameter()]
    [switch]$PassThru
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

#region Configuration

$ManufacturerConfigs = @{
    Dell = @{
        DriverPackPath = Join-Path -Path $CacheDirectory -ChildPath 'DriverPack' -AdditionalChildPath @('Dell', 'DriverPack_Dell.xml')
        WinPEPath = Join-Path -Path $CacheDirectory -ChildPath 'WinPE' -AdditionalChildPath @('Dell', 'WinPE_Dell.xml')
        CatalogUrl = 'https://downloads.dell.com/catalog/DriverPackCatalog.cab'
    }
    HP = @{
        DriverPackPath = Join-Path -Path $CacheDirectory -ChildPath 'DriverPack' -AdditionalChildPath @('HP', 'DriverPack_HP.xml')
        WinPEPath = Join-Path -Path $CacheDirectory -ChildPath 'WinPE' -AdditionalChildPath @('HP', 'WinPE_HP.xml')
        CatalogUrl = 'https://hpia.hpcloud.hp.com/downloads/driverpackcatalog/HPClientDriverPackCatalog.cab'
    }
    Lenovo = @{
        DriverPackPath = Join-Path -Path $CacheDirectory -ChildPath 'DriverPack' -AdditionalChildPath @('Lenovo', 'DriverPack_Lenovo.xml')
        CatalogUrl = 'https://download.lenovo.com/cdrt/td/catalogv2.xml'
    }
    Microsoft = @{
        DriverPackPath = Join-Path -Path $CacheDirectory -ChildPath 'DriverPack' -AdditionalChildPath @('Surface', 'DriverPack_Surface.xml')
        CatalogUrl = 'https://support.microsoft.com/en-us/surface/download-drivers-and-firmware-for-surface-09bb2e09-2a4b-cb69-0951-078a7739e120'
    }
}

#endregion Configuration

#region Import Functions

function Get-XmlElementText {
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Node,

        [Parameter(Mandatory = $true)]
        [string]$ElementName
    )

    if ($null -eq $Node) {
        return $null
    }

    if ($Node.PSObject.Properties.Name -contains $ElementName) {
        return [string]$Node.$ElementName
    }

    return $null
}

function Import-DellDriverPacks {
    param([string]$DriverPackPath, [string]$WinPEPath)

    $items = @()

    if (Test-Path -Path $DriverPackPath) {
        [xml]$xml = Get-Content -Path $DriverPackPath -Raw

        foreach ($item in $xml.DellCatalog.Items.Item) {
            $models = @()
            foreach ($system in $item.SupportedSystems.System) {
                $models += [pscustomobject]([ordered]@{
                        name = [string]$system.systemName
                        systemId = [string]$system.systemId
                    })
            }

            $osName = $null
            $osArch = 'x64'
            $osNodes = @($item.SupportedOperatingSystems.OperatingSystem)
            if ($osNodes.Count -gt 0 -and $osNodes[0]) {
                $osName = [string]$osNodes[0].osCode
                $osArch = [string]$osNodes[0].osArch
            }

            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.releaseId
                    packageId = [string]$item.releaseId
                    manufacturer = 'Dell'
                    name = [string]$item.name
                    version = [string]$item.dellVersion
                    fileName = [string]$item.name
                    downloadUrl = [string]$item.downloadUrl
                    sizeBytes = ConvertTo-Int64OrNull -Value (Get-XmlElementText -Node $item -ElementName 'sizeBytes')
                    format = [string]$item.format
                    type = 'Win'
                    releaseDate = [string]$item.dateTime
                    legacy = $null
                    models = $models
                    osName = $osName
                    osReleaseId = $null
                    osBuild = $null
                    osArchitecture = $osArch
                    hashMD5 = Get-XmlElementText -Node $item -ElementName 'hashMD5'
                    hashSHA256 = $null
                    hashCRC = $null
                })
        }
    }

    if (Test-Path -Path $WinPEPath) {
        [xml]$xml = Get-Content -Path $WinPEPath -Raw

        foreach ($item in $xml.DellCatalog.Items.Item) {
            $osName = 'WinPE'
            $osArch = 'x64'
            $osNodes = @($item.SupportedOperatingSystems.OperatingSystem)
            if ($osNodes.Count -gt 0 -and $osNodes[0]) {
                $osArch = [string]$osNodes[0].osArch
            }

            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.releaseId
                    packageId = [string]$item.releaseId
                    manufacturer = 'Dell'
                    name = [string]$item.name
                    version = [string]$item.dellVersion
                    fileName = [string]$item.name
                    downloadUrl = [string]$item.downloadUrl
                    sizeBytes = ConvertTo-Int64OrNull -Value (Get-XmlElementText -Node $item -ElementName 'sizeBytes')
                    format = [string]$item.format
                    type = 'WinPE'
                    releaseDate = [string]$item.dateTime
                    legacy = $null
                    models = @()
                    osName = $osName
                    osReleaseId = $null
                    osBuild = $null
                    osArchitecture = $osArch
                    hashMD5 = Get-XmlElementText -Node $item -ElementName 'hashMD5'
                    hashSHA256 = $null
                    hashCRC = $null
                })
        }
    }

    return $items
}

function Import-HPDriverPacks {
    param([string]$DriverPackPath, [string]$WinPEPath)

    $items = @()

    if (Test-Path -Path $DriverPackPath) {
        [xml]$xml = Get-Content -Path $DriverPackPath -Raw

        foreach ($item in $xml.HPCatalog.Items.Item) {
            $models = @([pscustomobject]([ordered]@{
                        name = [string]$item.systemName
                        systemId = [string]$item.systemId
                    }))

            $osName = [string]$item.osName
            $osVersion = $null
            $osReleaseId = $null
            if ($osName -match 'Windows (\d+)') {
                $osVersion = "Windows $($matches[1])"
            }
            if ($osName -match '(\d{4})$') {
                $osReleaseId = $matches[1]
            }

            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.id
                    packageId = [string]$item.softPaqId
                    manufacturer = 'HP'
                    name = [string]$item.name
                    version = [string]$item.softPaqVersion
                    fileName = "$($item.softPaqId).exe"
                    downloadUrl = [string]$item.downloadUrl
                    sizeBytes = ConvertTo-Int64OrNull -Value (Get-XmlElementText -Node $item -ElementName 'sizeBytes')
                    format = 'exe'
                    type = 'Win'
                    releaseDate = [string]$item.dateReleased
                    legacy = $null
                    models = $models
                    osName = $osVersion
                    osReleaseId = $osReleaseId
                    osBuild = $null
                    osArchitecture = if ([string]$item.architecture -eq '64-bit') { 'x64' } else { 'x86' }
                    hashMD5 = Get-XmlElementText -Node $item -ElementName 'md5'
                    hashSHA256 = Get-XmlElementText -Node $item -ElementName 'sha256'
                    hashCRC = $null
                })
        }
    }

    if (Test-Path -Path $WinPEPath) {
        [xml]$xml = Get-Content -Path $WinPEPath -Raw

        foreach ($item in $xml.HPCatalog.Items.Item) {
            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.id
                    packageId = [string]$item.softPaqId
                    manufacturer = 'HP'
                    name = [string]$item.name
                    version = [string]$item.softPaqVersion
                    fileName = "$($item.softPaqId).exe"
                    downloadUrl = [string]$item.downloadUrl
                    sizeBytes = ConvertTo-Int64OrNull -Value (Get-XmlElementText -Node $item -ElementName 'sizeBytes')
                    format = 'exe'
                    type = 'WinPE'
                    releaseDate = [string]$item.dateReleased
                    legacy = $null
                    models = @()
                    osName = 'WinPE'
                    osReleaseId = $null
                    osBuild = $null
                    osArchitecture = 'x64'
                    hashMD5 = Get-XmlElementText -Node $item -ElementName 'md5'
                    hashSHA256 = Get-XmlElementText -Node $item -ElementName 'sha256'
                    hashCRC = $null
                })
        }
    }

    return $items
}

function Import-LenovoDriverPacks {
    param([string]$DriverPackPath)

    $items = @()

    if (Test-Path -Path $DriverPackPath) {
        [xml]$xml = Get-Content -Path $DriverPackPath -Raw

        foreach ($item in $xml.LenovoCatalog.Items.Item) {
            $models = @([pscustomobject]([ordered]@{
                        name = [string]$item.model
                        systemId = [string]$item.machineTypes
                    }))

            $osValue = Get-XmlElementText -Node $item -ElementName 'os'
            $osVersion = if ($osValue) { "Windows $osValue" } else { $null }
            $osReleaseId = Get-XmlElementText -Node $item -ElementName 'osVersion'

            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.id
                    packageId = Get-XmlElementText -Node $item -ElementName 'fileName'
                    manufacturer = 'Lenovo'
                    name = Get-XmlElementText -Node $item -ElementName 'fileName'
                    version = $null
                    fileName = Get-XmlElementText -Node $item -ElementName 'fileName'
                    downloadUrl = Get-XmlElementText -Node $item -ElementName 'downloadUrl'
                    sizeBytes = $null
                    format = 'exe'
                    type = 'Win'
                    releaseDate = Get-XmlElementText -Node $item -ElementName 'releaseDate'
                    legacy = $null
                    models = $models
                    osName = $osVersion
                    osReleaseId = $osReleaseId
                    osBuild = $null
                    osArchitecture = 'x64'
                    hashMD5 = Get-XmlElementText -Node $item -ElementName 'md5'
                    hashSHA256 = $null
                    hashCRC = Get-XmlElementText -Node $item -ElementName 'crc'
                })
        }
    }

    return $items
}

function Import-MicrosoftDriverPacks {
    param([string]$DriverPackPath)

    $items = @()

    if (Test-Path -Path $DriverPackPath) {
        [xml]$xml = Get-Content -Path $DriverPackPath -Raw

        foreach ($item in $xml.SurfaceCatalog.Items.Item) {
            $models = @([pscustomobject]([ordered]@{
                        name = [string]$item.model
                        systemId = $null
                    }))

            $fileName = Get-XmlElementText -Node $item -ElementName 'fileName'
            $osVersion = $null
            $osReleaseId = $null
            $osBuild = $null

            if ($fileName -match 'Win(10|11)') {
                $osVersion = "Windows $($matches[1])"
            }
            if ($fileName -match '(\d{5})') {
                $osBuild = $matches[1]
                $buildToRelease = @{
                    '18362' = '1903'
                    '18363' = '1909'
                    '19041' = '2004'
                    '19042' = '20H2'
                    '19043' = '21H1'
                    '19044' = '21H2'
                    '19045' = '22H2'
                    '22621' = '22H2'
                    '22631' = '23H2'
                }
                $osReleaseId = $buildToRelease[$osBuild]
            }

            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.id
                    packageId = Get-XmlElementText -Node $item -ElementName 'packageId'
                    manufacturer = 'Microsoft'
                    name = Get-XmlElementText -Node $item -ElementName 'fileName'
                    version = $null
                    fileName = Get-XmlElementText -Node $item -ElementName 'fileName'
                    downloadUrl = Get-XmlElementText -Node $item -ElementName 'msiUrl'
                    sizeBytes = $null
                    format = 'msi'
                    type = 'Win'
                    releaseDate = $null
                    legacy = $null
                    models = $models
                    osName = $osVersion
                    osReleaseId = $osReleaseId
                    osBuild = $osBuild
                    osArchitecture = 'x64'
                    hashMD5 = $null
                    hashSHA256 = $null
                    hashCRC = $null
                })
        }
    }

    return $items
}

#endregion Import Functions

#region XML Generation

function Write-UnifiedDriverPackXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [array]$DriverPacks,

        [Parameter(Mandatory = $true)]
        [hashtable]$Sources,

        [Parameter(Mandatory = $true)]
        [ValidateSet('DriverPack', 'WinPE')]
        [string]$Category
    )

    $generatedAtUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $writer = New-CatalogXmlWriter -Path $Path

    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('DriverPackCatalog')
        $writer.WriteAttributeString('schemaVersion', '1')
        $writer.WriteAttributeString('generatedAtUtc', $generatedAtUtc)
        $writer.WriteAttributeString('totalItems', [string]$DriverPacks.Count)
        $writer.WriteAttributeString('category', $Category)

        $writer.WriteStartElement('Metadata')
        $writer.WriteAttributeString('name', "Foundry Unified $Category Catalog")
        $writer.WriteAttributeString('description', "Normalized $Category catalog containing driver packs from Dell, HP, Lenovo, and Microsoft")
        $writer.WriteEndElement()

        $writer.WriteStartElement('Sources')
        foreach ($source in ($Sources.GetEnumerator() | Sort-Object Key)) {
            $sourceItemCount = $source.Value.ItemCount
            if ($Category -eq 'WinPE') {
                $sourceItemCount = $source.Value.WinPEItemCount
            }
            if ($sourceItemCount -gt 0) {
                $writer.WriteStartElement('Source')
                $writer.WriteAttributeString('manufacturer', $source.Key)
                $writer.WriteAttributeString('catalogUrl', $source.Value.Url)
                if ($source.Value.Version) {
                    $writer.WriteAttributeString('catalogVersion', $source.Value.Version)
                }
                $writer.WriteAttributeString('lastUpdated', $source.Value.LastUpdated)
                $writer.WriteAttributeString('itemCount', [string]$sourceItemCount)
                $writer.WriteEndElement()
            }
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('DriverPacks')
        foreach ($pack in ($DriverPacks | Sort-Object manufacturer, name)) {
            Write-DriverPackElement -Writer $writer -DriverPack $pack
        }
        $writer.WriteEndElement()

        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    }
    finally {
        $writer.Dispose()
    }
}

function Write-DriverPackElement {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlWriter]$Writer,

        [Parameter(Mandatory = $true)]
        [object]$DriverPack
    )

    $Writer.WriteStartElement('DriverPack')
    $Writer.WriteAttributeString('id', $DriverPack.id)
    $Writer.WriteAttributeString('packageId', $DriverPack.packageId)
    $Writer.WriteAttributeString('manufacturer', $DriverPack.manufacturer)
    $Writer.WriteAttributeString('name', $DriverPack.name)
    if ($DriverPack.version) {
        $Writer.WriteAttributeString('version', $DriverPack.version)
    }
    $Writer.WriteAttributeString('fileName', $DriverPack.fileName)
    $Writer.WriteAttributeString('downloadUrl', $DriverPack.downloadUrl)
    if ($DriverPack.sizeBytes) {
        $Writer.WriteAttributeString('sizeBytes', [string]$DriverPack.sizeBytes)
    }
    $Writer.WriteAttributeString('format', $DriverPack.format)
    $Writer.WriteAttributeString('type', $DriverPack.type)
    if ($DriverPack.releaseDate) {
        $Writer.WriteAttributeString('releaseDate', $DriverPack.releaseDate)
    }
    if ($DriverPack.legacy -ne $null) {
        $Writer.WriteAttributeString('legacy', [string]$DriverPack.legacy.ToString().ToLower())
    }

    if ($DriverPack.models -and $DriverPack.models.Count -gt 0) {
        $Writer.WriteStartElement('Models')
        foreach ($model in $DriverPack.models) {
            $Writer.WriteStartElement('Model')
            $Writer.WriteAttributeString('name', $model.name)
            if ($model.systemId) {
                $Writer.WriteAttributeString('systemId', $model.systemId)
            }
            $Writer.WriteEndElement()
        }
        $Writer.WriteEndElement()
    }

    $Writer.WriteStartElement('OsInfo')
    $Writer.WriteAttributeString('name', $DriverPack.osName)
    if ($DriverPack.osReleaseId) {
        $Writer.WriteAttributeString('releaseId', $DriverPack.osReleaseId)
    }
    if ($DriverPack.osBuild) {
        $Writer.WriteAttributeString('build', $DriverPack.osBuild)
    }
    $Writer.WriteAttributeString('architecture', $DriverPack.osArchitecture)
    $Writer.WriteEndElement()

    if ($DriverPack.hashMD5 -or $DriverPack.hashSHA256 -or $DriverPack.hashCRC) {
        $Writer.WriteStartElement('Hashes')
        if ($DriverPack.hashMD5) {
            $Writer.WriteAttributeString('md5', $DriverPack.hashMD5)
        }
        if ($DriverPack.hashSHA256) {
            $Writer.WriteAttributeString('sha256', $DriverPack.hashSHA256)
        }
        if ($DriverPack.hashCRC) {
            $Writer.WriteAttributeString('crc', $DriverPack.hashCRC)
        }
        $Writer.WriteEndElement()
    }

    $Writer.WriteEndElement()
}

#endregion XML Generation

#region Main Execution

$startedAt = Get-Date

Write-Verbose "Importing Dell driver packs..."
$dellPacks = Import-DellDriverPacks -DriverPackPath $ManufacturerConfigs.Dell.DriverPackPath -WinPEPath $ManufacturerConfigs.Dell.WinPEPath

Write-Verbose "Importing HP driver packs..."
$hpPacks = Import-HPDriverPacks -DriverPackPath $ManufacturerConfigs.HP.DriverPackPath -WinPEPath $ManufacturerConfigs.HP.WinPEPath

Write-Verbose "Importing Lenovo driver packs..."
$lenovoPacks = Import-LenovoDriverPacks -DriverPackPath $ManufacturerConfigs.Lenovo.DriverPackPath

Write-Verbose "Importing Microsoft driver packs..."
$microsoftPacks = Import-MicrosoftDriverPacks -DriverPackPath $ManufacturerConfigs.Microsoft.DriverPackPath

$allPacks = @($dellPacks) + @($hpPacks) + @($lenovoPacks) + @($microsoftPacks)

# Separate DriverPack and WinPE
$driverPackItems = @($allPacks | Where-Object { $_.type -eq 'Win' })
$winPEItems = @($allPacks | Where-Object { $_.type -eq 'WinPE' })

Write-Verbose "Total driver packs imported: $($allPacks.Count)"
Write-Verbose "  - DriverPack items: $($driverPackItems.Count)"
Write-Verbose "  - WinPE items: $($winPEItems.Count)"

# Calculate counts per manufacturer
$dellDriverPackCount = @($dellPacks | Where-Object { $_.type -eq 'Win' }).Count
$dellWinPECount = @($dellPacks | Where-Object { $_.type -eq 'WinPE' }).Count
$hpDriverPackCount = @($hpPacks | Where-Object { $_.type -eq 'Win' }).Count
$hpWinPECount = @($hpPacks | Where-Object { $_.type -eq 'WinPE' }).Count
$lenovoDriverPackCount = @($lenovoPacks | Where-Object { $_.type -eq 'Win' }).Count
$microsoftDriverPackCount = @($microsoftPacks | Where-Object { $_.type -eq 'Win' }).Count

$sources = @{
    Dell = @{
        Url = $ManufacturerConfigs.Dell.CatalogUrl
        Version = $null
        LastUpdated = if ($dellPacks.Count -gt 0) { $dellPacks[0].releaseDate } else { (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
        ItemCount = $dellDriverPackCount
        WinPEItemCount = $dellWinPECount
    }
    HP = @{
        Url = $ManufacturerConfigs.HP.CatalogUrl
        Version = $null
        LastUpdated = if ($hpPacks.Count -gt 0) { $hpPacks[0].releaseDate } else { (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
        ItemCount = $hpDriverPackCount
        WinPEItemCount = $hpWinPECount
    }
    Lenovo = @{
        Url = $ManufacturerConfigs.Lenovo.CatalogUrl
        Version = $null
        LastUpdated = if ($lenovoPacks.Count -gt 0) { $lenovoPacks[0].releaseDate } else { (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
        ItemCount = $lenovoDriverPackCount
        WinPEItemCount = 0
    }
    Microsoft = @{
        Url = $ManufacturerConfigs.Microsoft.CatalogUrl
        Version = $null
        LastUpdated = if ($microsoftPacks.Count -gt 0) { $microsoftPacks[0].releaseDate } else { (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
        ItemCount = $microsoftDriverPackCount
        WinPEItemCount = 0
    }
}

# Create output directories
$driverPackOutputDir = Split-Path -Path $DriverPackOutputPath -Parent
$winPEOutputDir = Split-Path -Path $WinPEOutputPath -Parent

if (-not (Test-Path -Path $driverPackOutputDir)) {
    $null = New-Item -Path $driverPackOutputDir -ItemType Directory -Force
}
if (-not (Test-Path -Path $winPEOutputDir)) {
    $null = New-Item -Path $winPEOutputDir -ItemType Directory -Force
}

# Generate DriverPack unified catalog
Write-Verbose "Generating unified DriverPack XML at: $DriverPackOutputPath"
Write-UnifiedDriverPackXml -Path $DriverPackOutputPath -DriverPacks $driverPackItems -Sources $sources -Category 'DriverPack'
$driverPackXmlHash = (Get-FileHash -Path $DriverPackOutputPath -Algorithm SHA256).Hash.ToLowerInvariant()

# Generate WinPE unified catalog
Write-Verbose "Generating unified WinPE XML at: $WinPEOutputPath"
Write-UnifiedDriverPackXml -Path $WinPEOutputPath -DriverPacks $winPEItems -Sources $sources -Category 'WinPE'
$winPEXmlHash = (Get-FileHash -Path $WinPEOutputPath -Algorithm SHA256).Hash.ToLowerInvariant()

$durationSeconds = [int][Math]::Round(((Get-Date) - $startedAt).TotalSeconds)

Write-Verbose "Unified catalogs generated successfully"
Write-Verbose "  - DriverPack SHA256: $driverPackXmlHash"
Write-Verbose "  - WinPE SHA256: $winPEXmlHash"
Write-Verbose "Duration: $durationSeconds seconds"

$result = [pscustomobject]([ordered]@{
        DriverPackPath = $DriverPackOutputPath
        DriverPackItems = $driverPackItems.Count
        DriverPackSHA256 = $driverPackXmlHash
        WinPEPath = $WinPEOutputPath
        WinPEItems = $winPEItems.Count
        WinPESHA256 = $winPEXmlHash
        DellDriverPackCount = $dellDriverPackCount
        DellWinPECount = $dellWinPECount
        HPDriverPackCount = $hpDriverPackCount
        HPWinPECount = $hpWinPECount
        LenovoDriverPackCount = $lenovoDriverPackCount
        MicrosoftDriverPackCount = $microsoftDriverPackCount
        DurationSeconds = $durationSeconds
    })

if ($PassThru) {
    return $result
}

#endregion Main Execution
