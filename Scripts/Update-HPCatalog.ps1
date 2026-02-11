#region Parameters
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\DriverPack\HP'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$WinPEOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\WinPE\HP'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackCatalogCabUri = 'https://hpia.hpcloud.hp.com/downloads/driverpackcatalog/HPClientDriverPackCatalog.cab',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DriverPackCatalogXmlFileName = 'HPClientDriverPackCatalog.xml',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$PlatformListCabUri = 'https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/platformList.cab',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$PlatformListXmlFileName = 'platformList.xml',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$WinPECatalogHtmlUri = 'https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HP_WinPE_DriverPack.html',

    [Parameter()]
    [ValidateRange(0, 1000000)]
    [int]$MinimumDriverPackCount = 1,

    [Parameter()]
    [ValidateRange(0, 1000000)]
    [int]$MinimumWinPECount = 1
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

    $json = ConvertTo-Json -InputObject $Object -Depth 12
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

function Get-TemporaryRootPath {
    $tempPath = [System.IO.Path]::GetTempPath()
    if (-not $tempPath) {
        return '/tmp'
    }

    return $tempPath
}

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

function Get-XmlNodeText {
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
        $value = [string]$Node.InnerText
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    return $null
}

function Get-HPSchemaMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$CatalogXml
    )

    $root = $CatalogXml.NewDataSet.HPClientDriverPackCatalog
    if (-not $root) {
        throw 'Invalid HP driver pack catalog XML: HPClientDriverPackCatalog node not found.'
    }

    return [ordered]@{
        schemaVersion = if ([string]$root.SchemaVersion) { [string]$root.SchemaVersion } else { $null }
        toolVersion = if ([string]$root.ToolVersion) { [string]$root.ToolVersion } else { $null }
        dateReleased = if ([string]$root.DateReleased) { [string]$root.DateReleased } else { $null }
    }
}

function Get-HPPlatformListIndex {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$PlatformListXml
    )

    $platformNodes = @($PlatformListXml.ImagePal.Platform)
    $index = @{}

    foreach ($platform in $platformNodes) {
        if (-not $platform) {
            continue
        }

        $systemId = if ([string]$platform.SystemID) { ([string]$platform.SystemID).Trim() } else { $null }
        if (-not $systemId) {
            continue
        }

        $osNodes = @($platform.OS)
        $win10Releases = @()
        $win11Releases = @()

        foreach ($os in $osNodes) {
            $isWindows11 = ([string]$os.IsWindows11).ToLowerInvariant() -eq 'true'
            $release = if ([string]$os.OSReleaseIdDisplay) { [string]$os.OSReleaseIdDisplay } else { [string]$os.OSReleaseId }

            if (-not $release) {
                continue
            }

            if ($isWindows11) {
                $win11Releases += $release
            }
            else {
                $win10Releases += $release
            }
        }

        $index[$systemId] = [ordered]@{
            platformListMatched = $true
            platformProductName = Get-XmlNodeText -Node $platform.ProductName
            platformSystemFamily = if ([string]$platform.SystemFamily) { [string]$platform.SystemFamily } else { $null }
            platformDateAdded = if ($platform.Attributes['DateAdded']) { [string]$platform.Attributes['DateAdded'].Value } else { $null }
            platformDateLastModified = if ($platform.Attributes['DateLastModified']) { [string]$platform.Attributes['DateLastModified'].Value } else { $null }
            platformSupportsWin10 = $win10Releases.Count -gt 0
            platformSupportsWin11 = $win11Releases.Count -gt 0
            platformWin10Releases = if ($win10Releases.Count -gt 0) { (@($win10Releases | Sort-Object -Unique) -join ',') } else { $null }
            platformWin11Releases = if ($win11Releases.Count -gt 0) { (@($win11Releases | Sort-Object -Unique) -join ',') } else { $null }
        }
    }

    return $index
}

function ConvertFrom-HPDriverPackCatalogXml {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$CatalogXml,

        [Parameter(Mandatory = $true)]
        [hashtable]$PlatformListIndex
    )

    $root = $CatalogXml.NewDataSet.HPClientDriverPackCatalog
    if (-not $root) {
        throw 'Invalid HP driver pack catalog XML: HPClientDriverPackCatalog node not found.'
    }

    $softPaqById = @{}
    foreach ($softPaq in @($root.SoftPaqList.SoftPaq)) {
        $id = if ([string]$softPaq.Id) { [string]$softPaq.Id } else { $null }
        if (-not $id) {
            continue
        }

        $softPaqById[$id] = [ordered]@{
            id = $id
            name = Get-XmlNodeText -Node $softPaq.Name
            version = if ([string]$softPaq.Version) { [string]$softPaq.Version } else { $null }
            category = if ([string]$softPaq.Category) { [string]$softPaq.Category } else { $null }
            dateReleased = if ([string]$softPaq.DateReleased) { [string]$softPaq.DateReleased } else { $null }
            url = if ([string]$softPaq.Url) { [string]$softPaq.Url } else { $null }
            sizeBytes = ConvertTo-Int64OrNull -Value ([string]$softPaq.Size)
            md5 = if ([string]$softPaq.MD5) { [string]$softPaq.MD5 } else { $null }
            sha256 = if ([string]$softPaq.SHA256) { [string]$softPaq.SHA256 } else { $null }
            cvaFileUrl = if ([string]$softPaq.CvaFileUrl) { [string]$softPaq.CvaFileUrl } else { $null }
            releaseNotesUrl = if ([string]$softPaq.ReleaseNotesUrl) { [string]$softPaq.ReleaseNotesUrl } else { $null }
            hsaCompliant = if ([string]$softPaq.HSACompliant) { [string]$softPaq.HSACompliant } else { $null }
        }
    }

    $items = foreach ($mapping in @($root.ProductOSDriverPackList.ProductOSDriverPack)) {
        $softPaqId = if ([string]$mapping.SoftPaqId) { [string]$mapping.SoftPaqId } else { $null }
        $softPaq = $null
        if ($softPaqId -and $softPaqById.ContainsKey($softPaqId)) {
            $softPaq = $softPaqById[$softPaqId]
        }

        $idProduct = if ([string]$mapping.ProductId) { [string]$mapping.ProductId } else { '' }
        $idSystem = if ([string]$mapping.SystemId) { [string]$mapping.SystemId } else { '' }
        $idOs = if ([string]$mapping.OSId) { [string]$mapping.OSId } else { '' }
        $idSoftPaq = if ($softPaqId) { $softPaqId } else { '' }
        $id = @($idProduct, $idSystem, $idOs, $idSoftPaq) -join '|'

        $platformKey = if ([string]$mapping.SystemId) { ([string]$mapping.SystemId).Trim() } else { $null }
        $platformInfo = $null
        if ($platformKey -and $PlatformListIndex.ContainsKey($platformKey)) {
            $platformInfo = $PlatformListIndex[$platformKey]
        }

        [pscustomobject]([ordered]@{
                id = $id
                productId = if ([string]$mapping.ProductId) { [string]$mapping.ProductId } else { $null }
                productType = if ([string]$mapping.ProductType) { [string]$mapping.ProductType } else { $null }
                systemId = if ([string]$mapping.SystemId) { [string]$mapping.SystemId } else { $null }
                systemName = if ([string]$mapping.SystemName) { [string]$mapping.SystemName } else { $null }
                architecture = if ([string]$mapping.Architecture) { [string]$mapping.Architecture } else { $null }
                osId = if ([string]$mapping.OSId) { [string]$mapping.OSId } else { $null }
                osName = if ([string]$mapping.OSName) { [string]$mapping.OSName } else { $null }
                softPaqId = $softPaqId
                softPaqVersion = if ($softPaq) { [string]$softPaq.version } else { $null }
                name = if ($softPaq) { [string]$softPaq.name } else { Get-XmlNodeText -Node $mapping.Name }
                category = if ($softPaq) { [string]$softPaq.category } else { $null }
                dateReleased = if ($softPaq) { [string]$softPaq.dateReleased } else { $null }
                downloadUrl = if ($softPaq) { [string]$softPaq.url } else { $null }
                releaseNotesUrl = if ($softPaq) { [string]$softPaq.releaseNotesUrl } else { $null }
                sizeBytes = if ($softPaq) { $softPaq.sizeBytes } else { $null }
                md5 = if ($softPaq) { [string]$softPaq.md5 } else { $null }
                sha256 = if ($softPaq) { [string]$softPaq.sha256 } else { $null }
                cvaFileUrl = if ($softPaq) { [string]$softPaq.cvaFileUrl } else { $null }
                hsaCompliant = if ($softPaq) { [string]$softPaq.hsaCompliant } else { $null }
                winpeFamily = $null
                version = $null
                platformListMatched = if ($platformInfo) { [bool]$platformInfo.platformListMatched } else { $false }
                platformProductName = if ($platformInfo) { [string]$platformInfo.platformProductName } else { $null }
                platformSystemFamily = if ($platformInfo) { [string]$platformInfo.platformSystemFamily } else { $null }
                platformDateAdded = if ($platformInfo) { [string]$platformInfo.platformDateAdded } else { $null }
                platformDateLastModified = if ($platformInfo) { [string]$platformInfo.platformDateLastModified } else { $null }
                platformSupportsWin10 = if ($platformInfo) { [bool]$platformInfo.platformSupportsWin10 } else { $false }
                platformSupportsWin11 = if ($platformInfo) { [bool]$platformInfo.platformSupportsWin11 } else { $false }
                platformWin10Releases = if ($platformInfo) { [string]$platformInfo.platformWin10Releases } else { $null }
                platformWin11Releases = if ($platformInfo) { [string]$platformInfo.platformWin11Releases } else { $null }
            })
    }

    return @($items | Sort-Object -Property @(
            @{ Expression = { $_.productType } },
            @{ Expression = { $_.systemName } },
            @{ Expression = { $_.osName } },
            @{ Expression = { $_.softPaqId } }
        ))
}

function ConvertFrom-HPWinPEHtml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Html
    )

    $rows = [regex]::Matches($Html, '(?is)<tr>\s*<td>\s*WinPE.*?</tr>')
    $items = foreach ($row in $rows) {
        $cellMatches = [regex]::Matches($row.Value, '(?is)<td[^>]*>\s*(.*?)\s*</td>')
        if ($cellMatches.Count -lt 6) {
            continue
        }

        $winpeFamily = [regex]::Replace($cellMatches[0].Groups[1].Value, '<[^>]+>', '').Trim()
        $version = [regex]::Replace($cellMatches[1].Groups[1].Value, '<[^>]+>', '').Trim()
        $softPaqId = [regex]::Replace($cellMatches[2].Groups[1].Value, '<[^>]+>', '').Trim()
        $dateReleased = [regex]::Replace($cellMatches[3].Groups[1].Value, '<[^>]+>', '').Trim()

        $exeLinkMatch = [regex]::Match($cellMatches[4].Groups[1].Value, '(?is)href="([^"]+)"')
        $releaseNotesLinkMatch = [regex]::Match($cellMatches[5].Groups[1].Value, '(?is)href="([^"]+)"')

        $exeUrl = if ($exeLinkMatch.Success) { $exeLinkMatch.Groups[1].Value.Trim() } else { $null }
        $releaseNotesUrl = if ($releaseNotesLinkMatch.Success) { $releaseNotesLinkMatch.Groups[1].Value.Trim() } else { $null }

        if (-not $softPaqId) {
            continue
        }

        [pscustomobject]([ordered]@{
                id = $softPaqId
                productId = $null
                productType = $null
                systemId = $null
                systemName = $null
                architecture = 'x64'
                osId = $null
                osName = $winpeFamily
                softPaqId = $softPaqId
                softPaqVersion = $version
                name = $softPaqId
                category = 'WinPE'
                dateReleased = $dateReleased
                downloadUrl = $exeUrl
                releaseNotesUrl = $releaseNotesUrl
                sizeBytes = $null
                md5 = $null
                sha256 = $null
                cvaFileUrl = $null
                hsaCompliant = $null
                winpeFamily = $winpeFamily
                version = $version
            })
    }

    return @($items | Sort-Object -Descending -Property @(
            @{ Expression = { $_.dateReleased } },
            @{ Expression = { $_.softPaqId } }
        ))
}

function New-HPCatalog {
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

function Write-HPCatalogXml {
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
        $writer.WriteStartElement('HPCatalog')
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

function Write-HPCategoryOutputs {
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

    $filePrefix = if ($Catalog.category -eq 'DriverPack') { 'DriverPack_HP' } else { 'WinPE_HP' }
    $jsonPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.json')
    $xmlPath = Join-Path -Path $OutputDirectory -ChildPath ($filePrefix + '.xml')
    $mdPath = Join-Path -Path $OutputDirectory -ChildPath 'README.md'

    $json = ConvertTo-DeterministicJson -Object $Catalog
    Write-Utf8NoBomFile -Path $jsonPath -Content $json
    Write-HPCatalogXml -Path $xmlPath -Catalog $Catalog

    $jsonHash = (Get-FileHash -Path $jsonPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $xmlHash = (Get-FileHash -Path $xmlPath -Algorithm SHA256).Hash.ToLowerInvariant()

    $status = if ($Catalog.itemCount -gt 0) { 'SUCCESS' } else { 'WARNING' }
    $durationSeconds = [int][Math]::Round(((Get-Date) - $StartedAt).TotalSeconds)

    $summaryLines = @(
        '# HP Summary',
        '',
        '| Metric | Value |',
        '| --- | --- |',
        "| Executed At (UTC) | $($Catalog.generatedAtUtc -replace 'T', ' ' -replace 'Z', ' UTC') |",
        "| Category | $($Catalog.category) |",
        "| Status | $status |",
        "| Items | $($Catalog.itemCount) |",
        "| Catalog Version | $($Catalog.catalog.schemaVersion) |",
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

$sevenZipPath = Get-SevenZipCommandPath
$tempRoot = Join-Path -Path (Get-TemporaryRootPath) -ChildPath ('foundry-hp-catalog-' + [guid]::NewGuid())
$null = New-Item -Path $tempRoot -ItemType Directory -Force

try {
    $catalogCabPath = Join-Path -Path $tempRoot -ChildPath 'HPClientDriverPackCatalog.cab'
    $catalogXmlPath = Join-Path -Path $tempRoot -ChildPath $DriverPackCatalogXmlFileName
    $platformListCabPath = Join-Path -Path $tempRoot -ChildPath 'platformList.cab'
    $platformListXmlPath = Join-Path -Path $tempRoot -ChildPath $PlatformListXmlFileName

    Invoke-WebRequest -Uri $DriverPackCatalogCabUri -OutFile $catalogCabPath -ErrorAction Stop
    $driverPackCatalogCabSha256 = (Get-FileHash -Path $catalogCabPath -Algorithm SHA256).Hash.ToLowerInvariant()

    try {
        Invoke-SevenZipExtract -SevenZipPath $sevenZipPath -ArchivePath $catalogCabPath -OutputDirectory $tempRoot -Patterns @($DriverPackCatalogXmlFileName)
    }
    catch {
        Write-Verbose -Message ("7-Zip direct extraction failed for '{0}': {1}" -f $DriverPackCatalogXmlFileName, $_.Exception.Message)
    }

    if (-not (Test-Path -Path $catalogXmlPath)) {
        try {
            Invoke-SevenZipExtract -SevenZipPath $sevenZipPath -ArchivePath $catalogCabPath -OutputDirectory $tempRoot -Patterns @('*.xml')
        }
        catch {
            Write-Verbose -Message ("7-Zip wildcard extraction failed for HP driver pack catalog: {0}" -f $_.Exception.Message)
        }

        $xmlCandidates = @(Get-ChildItem -Path $tempRoot -Filter '*.xml' -File | Sort-Object -Descending -Property LastWriteTimeUtc, Name)
        if ($xmlCandidates.Count -ge 1) {
            Copy-Item -Path $xmlCandidates[0].FullName -Destination $catalogXmlPath -Force
        }
    }

    if (-not (Test-Path -Path $catalogXmlPath)) {
        throw "Catalog XML '$DriverPackCatalogXmlFileName' not found after CAB extraction."
    }

    Invoke-WebRequest -Uri $PlatformListCabUri -OutFile $platformListCabPath -ErrorAction Stop
    $platformListCabSha256 = (Get-FileHash -Path $platformListCabPath -Algorithm SHA256).Hash.ToLowerInvariant()

    try {
        Invoke-SevenZipExtract -SevenZipPath $sevenZipPath -ArchivePath $platformListCabPath -OutputDirectory $tempRoot -Patterns @($PlatformListXmlFileName)
    }
    catch {
        Write-Verbose -Message ("7-Zip direct extraction failed for '{0}': {1}" -f $PlatformListXmlFileName, $_.Exception.Message)
    }

    if (-not (Test-Path -Path $platformListXmlPath)) {
        try {
            Invoke-SevenZipExtract -SevenZipPath $sevenZipPath -ArchivePath $platformListCabPath -OutputDirectory $tempRoot -Patterns @('*.xml')
        }
        catch {
            Write-Verbose -Message ("7-Zip wildcard extraction failed for HP platform list: {0}" -f $_.Exception.Message)
        }

        $platformXmlCandidates = @(Get-ChildItem -Path $tempRoot -Filter '*.xml' -File | Where-Object { $_.FullName -ne $catalogXmlPath } | Sort-Object -Descending -Property LastWriteTimeUtc, Name)
        if ($platformXmlCandidates.Count -ge 1) {
            Copy-Item -Path $platformXmlCandidates[0].FullName -Destination $platformListXmlPath -Force
        }
    }

    if (-not (Test-Path -Path $platformListXmlPath)) {
        throw "Platform list XML '$PlatformListXmlFileName' not found after CAB extraction."
    }

    $driverPackCatalogXmlSha256 = (Get-FileHash -Path $catalogXmlPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $platformListXmlSha256 = (Get-FileHash -Path $platformListXmlPath -Algorithm SHA256).Hash.ToLowerInvariant()
    [xml]$catalogXml = Get-Content -Path $catalogXmlPath -Raw
    [xml]$platformListXml = Get-Content -Path $platformListXmlPath -Raw

    $platformListIndex = Get-HPPlatformListIndex -PlatformListXml $platformListXml

    $driverPackItems = ConvertFrom-HPDriverPackCatalogXml -CatalogXml $catalogXml -PlatformListIndex $platformListIndex
    if ($driverPackItems.Count -lt $MinimumDriverPackCount) {
        throw ("DriverPack item count below expected threshold (actual={0}, minimum={1})." -f $driverPackItems.Count, $MinimumDriverPackCount)
    }

    $winPEResponse = Invoke-WebRequest -Uri $WinPECatalogHtmlUri -ErrorAction Stop
    $winPEHtml = $winPEResponse.Content
    $winPEHtmlHash = [System.Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($winPEHtml))).ToLowerInvariant()
    $winPEItems = ConvertFrom-HPWinPEHtml -Html $winPEHtml
    if ($winPEItems.Count -lt $MinimumWinPECount) {
        throw ("WinPE item count below expected threshold (actual={0}, minimum={1})." -f $winPEItems.Count, $MinimumWinPECount)
    }

    $metadata = Get-HPSchemaMetadata -CatalogXml $catalogXml

    $driverPackCatalog = New-HPCatalog -Category 'DriverPack' -Source ([ordered]@{
            name = 'HP Client DriverPack Catalog'
            catalogType = 'DriverPackCatalogCab'
            uri = $DriverPackCatalogCabUri
            extractedFile = [System.IO.Path]::GetFileName($catalogXmlPath)
            catalogSha256 = $driverPackCatalogCabSha256
            extractedSha256 = $driverPackCatalogXmlSha256
            platformListUri = $PlatformListCabUri
            platformListCatalogSha256 = $platformListCabSha256
            platformListExtractedFile = [System.IO.Path]::GetFileName($platformListXmlPath)
            platformListExtractedSha256 = $platformListXmlSha256
        }) -CatalogMetadata ([ordered]@{
            schemaVersion = $metadata.schemaVersion
            toolVersion = $metadata.toolVersion
            dateReleased = $metadata.dateReleased
        }) -GeneratedAtUtc $generatedAtUtc -Items $driverPackItems

    $winPECatalog = New-HPCatalog -Category 'WinPE' -Source ([ordered]@{
            name = 'HP WinPE DriverPack Page'
            catalogType = 'WinPEHtmlPage'
            uri = $WinPECatalogHtmlUri
            extractedFile = $null
            catalogSha256 = $winPEHtmlHash
            extractedSha256 = $null
        }) -CatalogMetadata ([ordered]@{
            schemaVersion = $metadata.schemaVersion
            toolVersion = $metadata.toolVersion
            dateReleased = $metadata.dateReleased
        }) -GeneratedAtUtc $generatedAtUtc -Items $winPEItems

    $driverPackOutput = Write-HPCategoryOutputs -OutputDirectory $DriverPackOutputDirectory -Catalog $driverPackCatalog -StartedAt $startedAt
    $winPEOutput = Write-HPCategoryOutputs -OutputDirectory $WinPEOutputDirectory -Catalog $winPECatalog -StartedAt $startedAt

    [pscustomobject]@{
        DriverPack = $driverPackOutput
        WinPE = $winPEOutput
    }
}
finally {
    Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

#endregion Main Execution
