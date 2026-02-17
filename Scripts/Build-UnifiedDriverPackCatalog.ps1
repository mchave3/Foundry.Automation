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

function ConvertTo-ParsedDateTimeOffset {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Value
    )

    if (-not $Value) {
        return $null
    }

    $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces
    [datetimeoffset]$parsedOffset = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsedOffset)) {
        return $parsedOffset
    }
    if ([datetimeoffset]::TryParse($Value, [System.Globalization.CultureInfo]::CurrentCulture, $styles, [ref]$parsedOffset)) {
        return $parsedOffset
    }

    [datetime]$parsedDate = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsedDate)) {
        return [datetimeoffset]::new($parsedDate)
    }
    if ([datetime]::TryParse($Value, [System.Globalization.CultureInfo]::CurrentCulture, $styles, [ref]$parsedDate)) {
        return [datetimeoffset]::new($parsedDate)
    }

    return $null
}

function ConvertTo-IsoDateOrNull {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Value
    )

    if (-not $Value) {
        return $null
    }

    $normalized = $Value.Trim()
    if (-not $normalized) {
        return $null
    }

    if ($normalized -match '^\d{4}-\d{2}-\d{2}$') {
        return $normalized
    }

    $hasExplicitTime = ($normalized -match 'T\d{1,2}:\d{2}') -or ($normalized -match '\d{1,2}:\d{2}')
    if (-not $hasExplicitTime) {
        [datetime]$dateOnly = [datetime]::MinValue
        if ([datetime]::TryParse($normalized, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dateOnly)) {
            return $dateOnly.ToString('yyyy-MM-dd')
        }
        if ([datetime]::TryParse($normalized, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dateOnly)) {
            return $dateOnly.ToString('yyyy-MM-dd')
        }
    }

    $parsed = ConvertTo-ParsedDateTimeOffset -Value $normalized
    if ($null -eq $parsed) {
        return $null
    }

    return $parsed.ToUniversalTime().ToString('yyyy-MM-dd')
}

function Normalize-OsArchitecture {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Architecture,

        [Parameter()]
        [ValidateSet('x86', 'x64', 'arm64')]
        [string]$Default = 'x64'
    )

    if (-not $Architecture) {
        return $Default
    }

    switch -Regex ($Architecture.Trim().ToLowerInvariant()) {
        '^(x64|amd64|64-bit|64)$' { return 'x64' }
        '^(x86|86|32-bit|32|ia32)$' { return 'x86' }
        '^(arm64|aarch64)$' { return 'arm64' }
        default { return $Default }
    }
}

function Get-ReleaseIdFromBuild {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Build
    )

    [int]$buildNumber = 0
    if (-not [int]::TryParse($Build, [ref]$buildNumber)) {
        return $null
    }

    switch ($buildNumber) {
        { $_ -ge 26200 } { return '25H2' }
        { $_ -ge 26100 } { return '24H2' }
        { $_ -ge 22631 } { return '23H2' }
        { $_ -ge 22621 } { return '22H2' }
        { $_ -ge 22000 } { return '21H2' }
        { $_ -ge 19045 } { return '22H2' }
        { $_ -ge 19044 } { return '21H2' }
        { $_ -ge 19043 } { return '21H1' }
        { $_ -ge 19042 } { return '20H2' }
        { $_ -ge 19041 } { return '2004' }
        { $_ -ge 18363 } { return '1909' }
        { $_ -ge 18362 } { return '1903' }
        { $_ -ge 17763 } { return '1809' }
        { $_ -ge 17134 } { return '1803' }
        { $_ -ge 16299 } { return '1709' }
        { $_ -ge 15063 } { return '1703' }
        { $_ -ge 14393 } { return '1607' }
        { $_ -ge 10586 } { return '1511' }
        { $_ -ge 10240 } { return '1507' }
        default { return $null }
    }
}

function Normalize-LenovoOsName {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$OsValue
    )

    if (-not $OsValue) {
        return $null
    }

    $value = $OsValue.Trim()
    if ($value -match '^(?i)win(?:dows)?\s*(10|11)$') {
        return "Windows $($matches[1])"
    }
    if ($value -match '^(10|11)$') {
        return "Windows $value"
    }
    if ($value -match '^(?i)windows\s+') {
        return ('Windows ' + ($value -replace '^(?i)windows\s*', '').Trim())
    }

    return "Windows $value"
}

function Normalize-DellOsName {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$OsCode
    )

    if (-not $OsCode) {
        return $null
    }

    $value = $OsCode.Trim()
    switch -Regex ($value) {
        '^(?i)windows\s*11$|^Windows11$' { return 'Windows 11' }
        '^(?i)windows\s*10$|^Windows10$' { return 'Windows 10' }
        '^(?i)windows\s*8\.1$|^Windows8\.1$' { return 'Windows 8.1' }
        '^(?i)windows\s*8$|^Windows8$' { return 'Windows 8' }
        '^(?i)windows\s*7$|^Windows7$' { return 'Windows 7' }
        '^(?i)vista$' { return 'Windows Vista' }
        '^(?i)xp$' { return 'Windows XP' }
        default { return $value }
    }
}

function Get-WinPEReleaseId {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Text
    )

    if (-not $Text) {
        return $null
    }

    $normalized = $Text.Trim().ToLowerInvariant()
    if (-not $normalized) {
        return $null
    }

    if ($normalized -match '10\s*/\s*11') {
        return '10/11'
    }

    $codeMatch = [regex]::Match($normalized, 'winpe([0-9]+)x')
    if ($codeMatch.Success) {
        return $codeMatch.Groups[1].Value
    }

    $familyMatch = [regex]::Match($normalized, 'winpe\s*([0-9]+(?:\.[0-9]+)?)')
    if ($familyMatch.Success) {
        $majorMatch = [regex]::Match($familyMatch.Groups[1].Value, '^([0-9]+)')
        if ($majorMatch.Success) {
            return $majorMatch.Groups[1].Value
        }
    }

    $fallbackMatch = [regex]::Match($normalized, '(^|[^0-9])(11|10|5|4|3)([^0-9]|$)')
    if ($fallbackMatch.Success) {
        return $fallbackMatch.Groups[2].Value
    }

    return $null
}

function Get-ReleaseIdFromText {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Text
    )

    if (-not $Text) {
        return $null
    }

    $normalized = $Text.ToUpperInvariant()
    $match = [regex]::Match($normalized, '(25H2|24H2|23H2|22H2|21H2|21H1|20H2|2004|1909|1903|1809|1803|1709|1703|1607|1511|1507)')
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return $null
}

function Get-PreferredReleaseIdFromCsv {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$CsvValue
    )

    if (-not $CsvValue) {
        return $null
    }

    $candidates = @($CsvValue -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($candidates.Count -lt 1) {
        return $null
    }

    function Get-ReleaseScore {
        param([string]$Release)

        if ($Release -match '^(\d{2})H([12])$') {
            return ([int]$matches[1] * 10) + [int]$matches[2]
        }
        if ($Release -match '^\d{4}$') {
            return [int]$Release
        }

        return -1
    }

    $best = $null
    $bestScore = -1
    foreach ($candidate in $candidates) {
        $score = Get-ReleaseScore -Release $candidate
        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $candidate
        }
    }

    if ($best) {
        return $best
    }

    return $candidates[0]
}

function Normalize-DriverPackItems {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Items
    )

    $normalized = @()
    foreach ($item in @($Items)) {
        if ($null -eq $item) {
            continue
        }

        $downloadUrl = [string]$item.downloadUrl
        if (-not $downloadUrl) {
            Write-Warning ("Skipping unified item '{0}' because downloadUrl is empty." -f [string]$item.id)
            continue
        }

        $format = [string]$item.format
        if ($format) {
            $format = $format.ToLowerInvariant()
        }
        if (@('cab', 'exe', 'msi', 'zip') -notcontains $format) {
            if ($downloadUrl -match '(?i)\.cab($|[?&])') {
                $format = 'cab'
            }
            elseif ($downloadUrl -match '(?i)\.zip($|[?&])') {
                $format = 'zip'
            }
            elseif ($downloadUrl -match '(?i)\.msi($|[?&])') {
                $format = 'msi'
            }
            else {
                $format = 'exe'
            }
        }

        $normalized += [pscustomobject]([ordered]@{
                id = [string]$item.id
                packageId = [string]$item.packageId
                manufacturer = [string]$item.manufacturer
                name = [string]$item.name
                version = if ([string]$item.version) { [string]$item.version } else { $null }
                fileName = [string]$item.fileName
                downloadUrl = $downloadUrl
                sizeBytes = ConvertTo-Int64OrNull -Value $item.sizeBytes
                format = $format
                type = [string]$item.type
                releaseDate = ConvertTo-IsoDateOrNull -Value ([string]$item.releaseDate)
                legacy = $item.legacy
                models = @($item.models)
                osName = if ([string]$item.osName) { [string]$item.osName } else { 'Windows' }
                osReleaseId = if ([string]$item.osReleaseId) { [string]$item.osReleaseId } else { $null }
                osBuild = if ([string]$item.osBuild) { [string]$item.osBuild } else { $null }
                osArchitecture = Normalize-OsArchitecture -Architecture ([string]$item.osArchitecture) -Default 'x64'
                hashMD5 = if ([string]$item.hashMD5) { [string]$item.hashMD5 } else { $null }
                hashSHA256 = if ([string]$item.hashSHA256) { [string]$item.hashSHA256 } else { $null }
                hashCRC = if ([string]$item.hashCRC) { [string]$item.hashCRC } else { $null }
            })
    }

    return $normalized
}

function Get-LastUpdatedUtcFromItems {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Items
    )

    $latest = $null
    foreach ($item in @($Items)) {
        $releaseDateValue = [string]$item.releaseDate
        if (-not $releaseDateValue) {
            continue
        }

        $normalizedReleaseDate = $releaseDateValue.Trim()
        if (-not $normalizedReleaseDate) {
            continue
        }

        $utc = $null
        if ($normalizedReleaseDate -match '^\d{4}-\d{2}-\d{2}$') {
            [datetime]$exactDate = [datetime]::MinValue
            if ([datetime]::TryParseExact($normalizedReleaseDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$exactDate)) {
                $utc = [datetimeoffset]::new($exactDate, [timespan]::Zero)
            }
        }
        else {
            $parsed = ConvertTo-ParsedDateTimeOffset -Value $normalizedReleaseDate
            if ($null -ne $parsed) {
                $utc = $parsed.ToUniversalTime()
            }
        }

        if ($null -eq $utc) {
            continue
        }

        if (($null -eq $latest) -or ($utc -gt $latest)) {
            $latest = $utc
        }
    }

    if ($null -eq $latest) {
        return (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }

    return $latest.ToString('yyyy-MM-ddTHH:mm:ssZ')
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

            $osNodes = @($item.SupportedOperatingSystems.OperatingSystem)
            if ($osNodes.Count -lt 1) {
                $osNodes = @($null)
            }

            $baseReleaseId = [string]$item.releaseId
            $osSequence = 0
            foreach ($osNode in $osNodes) {
                $osSequence++
                $osCode = if ($osNode) { [string]$osNode.osCode } else { $null }
                $osArchRaw = if ($osNode) { [string]$osNode.osArch } else { $null }
                $osArch = Normalize-OsArchitecture -Architecture $osArchRaw -Default 'x64'
                $osName = Normalize-DellOsName -OsCode $osCode

                $id = $baseReleaseId
                if ($osNodes.Count -gt 1) {
                    $id = "{0}|{1}" -f $baseReleaseId, $osArch
                    if ($osSequence -gt 1 -and ($items | Where-Object { $_.id -eq $id })) {
                        $id = "{0}|{1}|{2}" -f $baseReleaseId, $osArch, $osSequence
                    }
                }

                $items += [pscustomobject]([ordered]@{
                        id = $id
                        packageId = $baseReleaseId
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
                        osReleaseId = Get-ReleaseIdFromText -Text $osCode
                        osBuild = $null
                        osArchitecture = $osArch
                        hashMD5 = Get-XmlElementText -Node $item -ElementName 'hashMD5'
                        hashSHA256 = Get-XmlElementText -Node $item -ElementName 'hashSHA256'
                        hashCRC = $null
                    })
            }
        }
    }

    if (Test-Path -Path $WinPEPath) {
        [xml]$xml = Get-Content -Path $WinPEPath -Raw

        foreach ($item in $xml.DellCatalog.Items.Item) {
            $osName = 'WinPE'
            $osNodes = @($item.SupportedOperatingSystems.OperatingSystem)
            if ($osNodes.Count -lt 1) {
                $osNodes = @($null)
            }

            $baseReleaseId = [string]$item.releaseId
            $osSequence = 0
            foreach ($osNode in $osNodes) {
                $osSequence++
                $osCode = if ($osNode) { [string]$osNode.osCode } else { $null }
                $osArchRaw = if ($osNode) { [string]$osNode.osArch } else { $null }
                $osArch = Normalize-OsArchitecture -Architecture $osArchRaw -Default 'x64'

                $id = $baseReleaseId
                if ($osNodes.Count -gt 1) {
                    $id = "{0}|{1}" -f $baseReleaseId, $osArch
                    if ($osSequence -gt 1 -and ($items | Where-Object { $_.id -eq $id })) {
                        $id = "{0}|{1}|{2}" -f $baseReleaseId, $osArch, $osSequence
                    }
                }

                $items += [pscustomobject]([ordered]@{
                        id = $id
                        packageId = $baseReleaseId
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
                        osReleaseId = Get-WinPEReleaseId -Text $osCode
                        osBuild = $null
                        osArchitecture = $osArch
                        hashMD5 = Get-XmlElementText -Node $item -ElementName 'hashMD5'
                        hashSHA256 = Get-XmlElementText -Node $item -ElementName 'hashSHA256'
                        hashCRC = $null
                    })
            }
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
            if ($osName -match '(?i)Windows\s*(10|11)') {
                $osVersion = "Windows $($matches[1])"
            }
            $osReleaseId = Get-ReleaseIdFromText -Text $osName
            if (-not $osReleaseId) {
                $preferredWin11 = Get-PreferredReleaseIdFromCsv -CsvValue (Get-XmlElementText -Node $item -ElementName 'platformWin11Releases')
                $preferredWin10 = Get-PreferredReleaseIdFromCsv -CsvValue (Get-XmlElementText -Node $item -ElementName 'platformWin10Releases')

                if ($osVersion -eq 'Windows 11' -and $preferredWin11) {
                    $osReleaseId = $preferredWin11
                }
                elseif ($osVersion -eq 'Windows 10' -and $preferredWin10) {
                    $osReleaseId = $preferredWin10
                }
                elseif ($preferredWin11) {
                    $osReleaseId = $preferredWin11
                }
                elseif ($preferredWin10) {
                    $osReleaseId = $preferredWin10
                }
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
                    osName = if ($osVersion) { $osVersion } else { 'Windows' }
                    osReleaseId = $osReleaseId
                    osBuild = $null
                    osArchitecture = Normalize-OsArchitecture -Architecture ([string]$item.architecture) -Default 'x64'
                    hashMD5 = Get-XmlElementText -Node $item -ElementName 'md5'
                    hashSHA256 = Get-XmlElementText -Node $item -ElementName 'sha256'
                    hashCRC = $null
                })
        }
    }

    if (Test-Path -Path $WinPEPath) {
        [xml]$xml = Get-Content -Path $WinPEPath -Raw

        foreach ($item in $xml.HPCatalog.Items.Item) {
            $winpeFamily = Get-XmlElementText -Node $item -ElementName 'winpeFamily'
            if (-not $winpeFamily) {
                $winpeFamily = Get-XmlElementText -Node $item -ElementName 'osName'
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
                    type = 'WinPE'
                    releaseDate = [string]$item.dateReleased
                    legacy = $null
                    models = @()
                    osName = 'WinPE'
                    osReleaseId = Get-WinPEReleaseId -Text $winpeFamily
                    osBuild = $null
                    osArchitecture = Normalize-OsArchitecture -Architecture ([string]$item.architecture) -Default 'x64'
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
            $osVersion = Normalize-LenovoOsName -OsValue $osValue
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
            $models = @()
            $modelName = Get-XmlElementText -Node $item -ElementName 'model'
            if (-not $modelName) {
                $modelName = Get-XmlElementText -Node $item -ElementName 'downloadTitle'
            }
            if ($modelName) {
                $models += [pscustomobject]([ordered]@{
                        name = $modelName
                        systemId = $null
                    })
            }

            $downloadUrl = Get-XmlElementText -Node $item -ElementName 'downloadUrl'
            if (-not $downloadUrl) {
                $downloadUrl = Get-XmlElementText -Node $item -ElementName 'msiUrl'
            }
            if (-not $downloadUrl) {
                $downloadUrl = Get-XmlElementText -Node $item -ElementName 'downloadCenterUrl'
            }
            if (-not $downloadUrl) {
                Write-Warning ("Skipping Microsoft item '{0}' because no download URL was found." -f [string]$item.id)
                continue
            }

            $fileName = Get-XmlElementText -Node $item -ElementName 'fileName'
            if (-not $fileName) {
                try {
                    $fileName = [System.IO.Path]::GetFileName(([System.Uri]$downloadUrl).LocalPath)
                }
                catch {
                    $fileName = [System.IO.Path]::GetFileName($downloadUrl)
                }
            }

            $format = Get-XmlElementText -Node $item -ElementName 'format'
            if (-not $format) {
                if ($fileName) {
                    $extension = [System.IO.Path]::GetExtension($fileName)
                    if ($extension) {
                        $format = $extension.TrimStart('.').ToLowerInvariant()
                    }
                }
                if (-not $format) {
                    if ($downloadUrl -match '(?i)\.msi($|[?&])') {
                        $format = 'msi'
                    }
                    elseif ($downloadUrl -match '(?i)\.zip($|[?&])') {
                        $format = 'zip'
                    }
                    elseif ($downloadUrl -match '(?i)\.cab($|[?&])') {
                        $format = 'cab'
                    }
                    else {
                        $format = 'exe'
                    }
                }
            }

            $osVersion = Get-XmlElementText -Node $item -ElementName 'osName'
            $osReleaseId = Get-XmlElementText -Node $item -ElementName 'osReleaseId'
            $osBuild = Get-XmlElementText -Node $item -ElementName 'osBuild'
            $supportedOperatingSystems = Get-XmlElementText -Node $item -ElementName 'supportedOperatingSystems'

            if (-not $osVersion -and $fileName) {
                if ($fileName -match '(?i)Win(?:dows)?[_\-]?(10|11)(?:[_\-.]|$)') {
                    $osVersion = "Windows $($matches[1])"
                }
                elseif ($fileName -match '(?i)Win(?:dows)?[_\-]?(8x|8\.1|81)(?:[_\-.]|$)') {
                    $osVersion = 'Windows 8.1'
                }
                elseif ($fileName -match '(?i)Win(?:dows)?[_\-]?8(?:[_\-.]|$)') {
                    $osVersion = 'Windows 8'
                }
                elseif ($fileName -match '(?i)Win(?:dows)?[_\-]?7(?:[_\-.]|$)') {
                    $osVersion = 'Windows 7'
                }
            }

            if (-not $osVersion -and $supportedOperatingSystems) {
                if ($supportedOperatingSystems -match '(?i)Windows\s*11') {
                    $osVersion = 'Windows 11'
                }
                elseif ($supportedOperatingSystems -match '(?i)Windows\s*10') {
                    $osVersion = 'Windows 10'
                }
                elseif ($supportedOperatingSystems -match '(?i)Windows\s*8\.1') {
                    $osVersion = 'Windows 8.1'
                }
                elseif ($supportedOperatingSystems -match '(?i)Windows\s*8') {
                    $osVersion = 'Windows 8'
                }
                elseif ($supportedOperatingSystems -match '(?i)Windows\s*7') {
                    $osVersion = 'Windows 7'
                }
            }

            if (-not $osBuild -and $fileName -match '(\d{5})') {
                $osBuild = $matches[1]
            }

            if (-not $osReleaseId -and $osBuild) {
                $osReleaseId = Get-ReleaseIdFromBuild -Build $osBuild
            }

            if (-not $osReleaseId -and $fileName) {
                $osReleaseId = Get-ReleaseIdFromText -Text $fileName
            }

            if (-not $osReleaseId -and $fileName) {
                $legacyReleaseMatch = [regex]::Match($fileName, '(?i)(1507|1511|1607|1703|1709|1803|1809|1903|1909|2004)')
                if ($legacyReleaseMatch.Success) {
                    $osReleaseId = $legacyReleaseMatch.Groups[1].Value.ToUpperInvariant()
                }
            }

            if ((-not $osVersion) -and $osBuild) {
                [int]$parsedBuild = 0
                if ([int]::TryParse($osBuild, [ref]$parsedBuild)) {
                    $osVersion = if ($parsedBuild -ge 22000) { 'Windows 11' } else { 'Windows 10' }
                }
            }

            $osArchitecture = 'x64'
            if ($fileName -match '(?i)(arm64|aarch64)') {
                $osArchitecture = 'arm64'
            }
            elseif ($fileName -match '(?i)x86') {
                $osArchitecture = 'x86'
            }

            $items += [pscustomobject]([ordered]@{
                    id = [string]$item.id
                    packageId = if (Get-XmlElementText -Node $item -ElementName 'packageId') { Get-XmlElementText -Node $item -ElementName 'packageId' } else { [string]$item.id }
                    manufacturer = 'Microsoft'
                    name = if ($fileName) { $fileName } else { [string]$item.id }
                    version = Get-XmlElementText -Node $item -ElementName 'version'
                    fileName = if ($fileName) { $fileName } else { [string]$item.id }
                    downloadUrl = $downloadUrl
                    sizeBytes = ConvertTo-Int64OrNull -Value (Get-XmlElementText -Node $item -ElementName 'sizeBytes')
                    format = $format
                    type = 'Win'
                    releaseDate = Get-XmlElementText -Node $item -ElementName 'datePublished'
                    legacy = $null
                    models = $models
                    osName = if ($osVersion) { $osVersion } else { 'Windows' }
                    osReleaseId = $osReleaseId
                    osBuild = $osBuild
                    osArchitecture = $osArchitecture
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

$allPacks = Normalize-DriverPackItems -Items (@($dellPacks) + @($hpPacks) + @($lenovoPacks) + @($microsoftPacks))

# Separate DriverPack and WinPE
$driverPackItems = @($allPacks | Where-Object { $_.type -eq 'Win' })
$winPEItems = @($allPacks | Where-Object { $_.type -eq 'WinPE' })

Write-Verbose "Total driver packs imported: $($allPacks.Count)"
Write-Verbose "  - DriverPack items: $($driverPackItems.Count)"
Write-Verbose "  - WinPE items: $($winPEItems.Count)"

# Calculate counts per manufacturer
$dellDriverPackItems = @($driverPackItems | Where-Object { $_.manufacturer -eq 'Dell' })
$dellWinPEItems = @($winPEItems | Where-Object { $_.manufacturer -eq 'Dell' })
$hpDriverPackItems = @($driverPackItems | Where-Object { $_.manufacturer -eq 'HP' })
$hpWinPEItems = @($winPEItems | Where-Object { $_.manufacturer -eq 'HP' })
$lenovoDriverPackItems = @($driverPackItems | Where-Object { $_.manufacturer -eq 'Lenovo' })
$microsoftDriverPackItems = @($driverPackItems | Where-Object { $_.manufacturer -eq 'Microsoft' })

$dellDriverPackCount = $dellDriverPackItems.Count
$dellWinPECount = $dellWinPEItems.Count
$hpDriverPackCount = $hpDriverPackItems.Count
$hpWinPECount = $hpWinPEItems.Count
$lenovoDriverPackCount = $lenovoDriverPackItems.Count
$microsoftDriverPackCount = $microsoftDriverPackItems.Count

$driverPackSources = @{
    Dell = @{
        Url = $ManufacturerConfigs.Dell.CatalogUrl
        Version = $null
        LastUpdated = Get-LastUpdatedUtcFromItems -Items $dellDriverPackItems
        ItemCount = $dellDriverPackCount
    }
    HP = @{
        Url = $ManufacturerConfigs.HP.CatalogUrl
        Version = $null
        LastUpdated = Get-LastUpdatedUtcFromItems -Items $hpDriverPackItems
        ItemCount = $hpDriverPackCount
    }
    Lenovo = @{
        Url = $ManufacturerConfigs.Lenovo.CatalogUrl
        Version = $null
        LastUpdated = Get-LastUpdatedUtcFromItems -Items $lenovoDriverPackItems
        ItemCount = $lenovoDriverPackCount
    }
    Microsoft = @{
        Url = $ManufacturerConfigs.Microsoft.CatalogUrl
        Version = $null
        LastUpdated = Get-LastUpdatedUtcFromItems -Items $microsoftDriverPackItems
        ItemCount = $microsoftDriverPackCount
    }
}

$winPESources = @{
    Dell = @{
        Url = $ManufacturerConfigs.Dell.CatalogUrl
        Version = $null
        LastUpdated = Get-LastUpdatedUtcFromItems -Items $dellWinPEItems
        ItemCount = $dellWinPECount
    }
    HP = @{
        Url = $ManufacturerConfigs.HP.CatalogUrl
        Version = $null
        LastUpdated = Get-LastUpdatedUtcFromItems -Items $hpWinPEItems
        ItemCount = $hpWinPECount
    }
    Lenovo = @{
        Url = $ManufacturerConfigs.Lenovo.CatalogUrl
        Version = $null
        LastUpdated = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        ItemCount = 0
    }
    Microsoft = @{
        Url = $ManufacturerConfigs.Microsoft.CatalogUrl
        Version = $null
        LastUpdated = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        ItemCount = 0
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
Write-UnifiedDriverPackXml -Path $DriverPackOutputPath -DriverPacks $driverPackItems -Sources $driverPackSources -Category 'DriverPack'
$driverPackXmlHash = (Get-FileHash -Path $DriverPackOutputPath -Algorithm SHA256).Hash.ToLowerInvariant()

# Generate WinPE unified catalog
Write-Verbose "Generating unified WinPE XML at: $WinPEOutputPath"
Write-UnifiedDriverPackXml -Path $WinPEOutputPath -DriverPacks $winPEItems -Sources $winPESources -Category 'WinPE'
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
