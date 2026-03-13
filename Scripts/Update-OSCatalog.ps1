#region Parameters
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\OS'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SourceOutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\OS\Microsoft'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('23H2', '24H2', '25H2')]
    [string[]]$TargetReleases = @('23H2', '24H2', '25H2'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]]$ClientTypes = @('CLIENTCONSUMER', 'CLIENTBUSINESS'),

    [Parameter()]
    [System.Management.Automation.SwitchParameter]$IncludeKey,

    [Parameter()]
    [ValidateRange(1, 1000000)]
    [int]$MinimumItemCount = 50
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

#region Utility Functions

$script:windows11ReleaseSources = @{
    '23H2' = [pscustomobject]@{
        ReleaseId = '23H2'
        SourceType = 'StaticCab'
        CabUrl = 'https://download.microsoft.com/download/6/2/b/62b47bc5-1b28-4bfa-9422-e7a098d326d4/products_win11_20231208.cab'
        ExpectedBuildMajor = 22631
    }
    '24H2' = [pscustomobject]@{
        ReleaseId = '24H2'
        SourceType = 'StaticCab'
        CabUrl = 'https://download.microsoft.com/download/8e0c23e7-ddc2-45c4-b7e1-85a808b408ee/Products-Win11-24H2-6B.cab'
        ExpectedBuildMajor = 26100
    }
    '25H2' = [pscustomobject]@{
        ReleaseId = '25H2'
        SourceType = 'DynamicWindowsUpdate'
        Products = 'PN=Windows.Products.Cab.amd64&V=0.0.0.0'
        DeviceAttributes = 'DUScan=1;OSVersion=10.0.26100.1'
        ExpectedBuildMajor = 26200
    }
}

function Resolve-DirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -Force
    }

    return (Resolve-Path -Path $Path).Path
}

function Normalize-OsArchitecture {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Architecture
    )

    if (-not $Architecture) {
        return $null
    }

    $normalized = $Architecture.Trim().ToLowerInvariant()
    switch -Regex ($normalized) {
        '^(x64|amd64|x86_64|64)$' { return 'x64' }
        '^(x86|86|ia32|32)$' { return 'x86' }
        '^(arm64|aarch64|a64)$' { return 'arm64' }
        default { return $Architecture.Trim() }
    }
}

function Normalize-DownloadUrl {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Url
    )

    if (-not $Url) {
        return $null
    }

    $normalized = $Url.Trim()
    if (-not $normalized) {
        return $null
    }

    if ($normalized -match '^http://') {
        return 'https://' + $normalized.Substring(7)
    }

    return $normalized
}

function Get-ClientTypeFromText {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Text
    )

    if (-not $Text) {
        return $null
    }

    $value = $Text.ToUpperInvariant()
    if ($value -match 'CLIENTBUSINESS|_VOL_|(^|[^A-Z])VOL([^A-Z]|$)') {
        return 'CLIENTBUSINESS'
    }
    if ($value -match 'CLIENTCONSUMER|_RET_|(^|[^A-Z])RET([^A-Z]|$)') {
        return 'CLIENTCONSUMER'
    }

    return $null
}

function Get-LicenseChannelFromText {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Text
    )

    if (-not $Text) {
        return $null
    }

    $value = $Text.ToUpperInvariant()
    if ($value -match '_VOL_|(^|[^A-Z])VOL([^A-Z]|$)') {
        return 'VOL'
    }
    if ($value -match '_RET_|(^|[^A-Z])RET([^A-Z]|$)') {
        return 'RET'
    }

    return $null
}

function Get-WindowsReleaseIdFromBuildMajor {
    param(
        [Parameter()]
        [AllowNull()]
        [int]$BuildMajor
    )

    if (-not $BuildMajor) {
        return $null
    }

    switch ($BuildMajor) {
        { $_ -ge 26200 } { return '25H2' }
        { $_ -ge 26100 } { return '24H2' }
        { $_ -ge 22631 } { return '23H2' }
        default { return $null }
    }
}

function Get-SourceDefinitionFromFileName {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $match = [regex]::Match($baseName, '^(?i)Win(?<windowsMajor>\d+)_(?<releaseId>\d{2}H\d)_(?<buildMajor>\d{5})$')
    if (-not $match.Success) {
        throw ("Source XML name '{0}' must match 'Win<major>_<releaseId>_<buildMajor>.xml'." -f $FileName)
    }

    [int]$parsedBuildMajor = 0
    if (-not [int]::TryParse($match.Groups['buildMajor'].Value, [ref]$parsedBuildMajor)) {
        throw ("Source XML name '{0}' contains an invalid build major." -f $FileName)
    }

    return [pscustomobject]([ordered]@{
            id = $baseName
            windowsMajor = $match.Groups['windowsMajor'].Value
            releaseId = $match.Groups['releaseId'].Value.ToUpperInvariant()
            buildMajor = $parsedBuildMajor
        })
}

function Get-WindowsUpdateProductsCabMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Products,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DeviceAttributes
    )

    $body = [ordered]@{
        Products         = $Products
        DeviceAttributes = $DeviceAttributes
    } | ConvertTo-Json -Compress

    $response = Invoke-RestMethod -Uri 'https://fe3.delivery.mp.microsoft.com/UpdateMetadataService/updates/search/v1/bydeviceinfo' `
        -Method Post `
        -ContentType 'application/json' `
        -Headers @{ Accept = '*/*' } `
        -Body $body `
        -ErrorAction Stop

    if ($response -is [System.Array]) {
        $response = $response[0]
    }

    if (-not $response -or -not $response.FileLocations) {
        throw "Windows Update metadata response did not include file locations."
    }

    $fileRecord = $response.FileLocations |
        Where-Object { $_.FileName -eq 'products.cab' } |
        Select-Object -First 1

    if (-not $fileRecord) {
        throw "Windows Update metadata response did not include products.cab."
    }

    return [pscustomobject]@{
        Products = $Products
        DeviceAttributes = $DeviceAttributes
        DownloadUrl = [string]$fileRecord.Url
        DigestBase64 = [string]$fileRecord.Digest
        SizeBytes = [long]$fileRecord.Size
    }
}

function Save-VerifiedProductsCab {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$Metadata,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationPath
    )

    $destinationDirectory = Split-Path -Path $DestinationPath -Parent
    if ($destinationDirectory -and -not (Test-Path -Path $destinationDirectory)) {
        $null = New-Item -Path $destinationDirectory -ItemType Directory -Force
    }

    Invoke-WebRequest -Uri $Metadata.DownloadUrl -OutFile $DestinationPath -Headers @{ Accept = '*/*' } -ErrorAction Stop

    $downloadedSize = (Get-Item -Path $DestinationPath -ErrorAction Stop).Length
    if ($downloadedSize -ne $Metadata.SizeBytes) {
        throw "Downloaded products.cab size mismatch. Expected $($Metadata.SizeBytes) bytes, got $downloadedSize bytes."
    }

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $stream = [System.IO.File]::OpenRead($DestinationPath)
    try {
        $hashBytes = $sha256.ComputeHash($stream)
    }
    finally {
        $stream.Dispose()
        $sha256.Dispose()
    }

    $digestBase64 = [Convert]::ToBase64String($hashBytes)
    if ($digestBase64 -ne $Metadata.DigestBase64) {
        throw "Downloaded products.cab digest mismatch for products query '$($Metadata.Products)'."
    }
}

function Get-ProductsXmlContentFromCab {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CabPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SevenZipPath
    )

    $tempExtractDirectory = Join-Path -Path (Get-TemporaryRootPath) -ChildPath ('foundry-os-products-' + [guid]::NewGuid())
    try {
        $null = New-Item -Path $tempExtractDirectory -ItemType Directory -Force
        Invoke-SevenZipExtract -SevenZipPath $SevenZipPath -ArchivePath $CabPath -OutputDirectory $tempExtractDirectory -Patterns @('*.xml')

        $xmlFile = Get-ChildItem -Path $tempExtractDirectory -Filter '*.xml' -File -ErrorAction Stop | Select-Object -First 1
        if (-not $xmlFile) {
            throw "products.cab did not contain an XML file."
        }

        return Get-Content -Path $xmlFile.FullName -Raw -ErrorAction Stop
    }
    finally {
        if (Test-Path -Path $tempExtractDirectory) {
            Remove-Item -Path $tempExtractDirectory -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-Windows11ReleaseSourceDefinition {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('23H2', '24H2', '25H2')]
        [string]$ReleaseId
    )

    $normalizedReleaseId = $ReleaseId.ToUpperInvariant()
    if (-not $script:windows11ReleaseSources.ContainsKey($normalizedReleaseId)) {
        throw "Unsupported Windows 11 release: $ReleaseId"
    }

    return $script:windows11ReleaseSources[$normalizedReleaseId]
}

function Save-ProductsCabForRelease {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$ReleaseDefinition,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationPath
    )

    switch ($ReleaseDefinition.SourceType) {
        'StaticCab' {
            Invoke-WebRequest -Uri $ReleaseDefinition.CabUrl -OutFile $DestinationPath -Headers @{ Accept = '*/*' } -ErrorAction Stop
        }
        'DynamicWindowsUpdate' {
            $cabMetadata = Get-WindowsUpdateProductsCabMetadata -Products $ReleaseDefinition.Products -DeviceAttributes $ReleaseDefinition.DeviceAttributes
            Save-VerifiedProductsCab -Metadata $cabMetadata -DestinationPath $DestinationPath
        }
        default {
            throw "Unsupported source type '$($ReleaseDefinition.SourceType)' for release $($ReleaseDefinition.ReleaseId)."
        }
    }
}

function Get-RepresentativeReleaseItem {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject[]]$Items,

        [Parameter(Mandatory = $true)]
        [ValidateSet('23H2', '24H2', '25H2')]
        [string]$ReleaseId,

        [Parameter()]
        [AllowNull()]
        [int]$ExpectedBuildMajor
    )

    $releaseItems = @(
        $Items |
        Where-Object { $_.windowsRelease -eq 11 -and $_.releaseId -eq $ReleaseId }
    )

    if ($releaseItems.Count -lt 1) {
        throw "Downloaded products.xml did not contain any Windows 11 $ReleaseId ESD entries after filtering."
    }

    $representativeItem = $releaseItems |
        Sort-Object -Descending -Property @(
            @{ Expression = { if ($null -eq $_.buildUbr) { -1 } else { $_.buildUbr } } },
            @{ Expression = { $_.buildMajor } },
            @{ Expression = { $_.build } },
            @{ Expression = { $_.fileName } }
        ) |
        Select-Object -First 1

    if ($ExpectedBuildMajor -and $representativeItem.buildMajor -ne $ExpectedBuildMajor) {
        throw "Windows 11 $ReleaseId source returned build major $($representativeItem.buildMajor), expected $ExpectedBuildMajor."
    }

    return $representativeItem
}

function Get-ProductsSourceFiles {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceOutputDirectory,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string[]]$TargetReleases,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string[]]$ClientTypes,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SevenZipPath,

        [Parameter()]
        [System.Management.Automation.SwitchParameter]$IncludeKey
    )

    $resolvedSourceOutputDirectory = Resolve-DirectoryPath -Path $SourceOutputDirectory
    $normalizedTargetReleases = @(
        $TargetReleases |
        Where-Object { $_ } |
        ForEach-Object { $_.Trim().ToUpperInvariant() } |
        Select-Object -Unique
    )

    $generatedFilePaths = New-Object System.Collections.Generic.List[string]
    foreach ($releaseId in $normalizedTargetReleases) {
        $releaseDefinition = Get-Windows11ReleaseSourceDefinition -ReleaseId $releaseId
        $tempRoot = Join-Path -Path (Get-TemporaryRootPath) -ChildPath ('foundry-os-wu-' + [guid]::NewGuid())
        try {
            $null = New-Item -Path $tempRoot -ItemType Directory -Force

            $productsCabPath = Join-Path -Path $tempRoot -ChildPath 'products.cab'
            Save-ProductsCabForRelease -ReleaseDefinition $releaseDefinition -DestinationPath $productsCabPath

            $productsXmlContent = Get-ProductsXmlContentFromCab -CabPath $productsCabPath -SevenZipPath $SevenZipPath
            [xml]$productsXml = $productsXmlContent
            $sourceItems = @(
                ConvertFrom-ProductsXml -ProductsXml $productsXml -SourceId ("Win11Dynamic_{0}" -f $releaseId) -ClientTypes $ClientTypes -IncludeKey:$IncludeKey
            )

            if ($sourceItems.Count -lt 1) {
                throw "Downloaded products.xml for Windows 11 $releaseId did not yield any matching ESD items."
            }

            $representativeItem = Get-RepresentativeReleaseItem -Items $sourceItems -ReleaseId $releaseId -ExpectedBuildMajor $releaseDefinition.ExpectedBuildMajor
            if (-not $representativeItem.buildMajor) {
                throw "Unable to determine build major for Windows 11 $releaseId."
            }

            $sourceFileName = "Win11_{0}_{1}.xml" -f $releaseId, $representativeItem.buildMajor
            $sourceFilePath = Join-Path -Path $resolvedSourceOutputDirectory -ChildPath $sourceFileName
            Write-Utf8NoBomFile -Path $sourceFilePath -Content $productsXmlContent
            $generatedFilePaths.Add($sourceFilePath)
        }
        finally {
            if (Test-Path -Path $tempRoot) {
                Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $generatedFileNames = @($generatedFilePaths | ForEach-Object { [System.IO.Path]::GetFileName($_) })
    $staleFiles = @(
        Get-ChildItem -Path $resolvedSourceOutputDirectory -Filter 'Win11_*.xml' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notin $generatedFileNames }
    )
    foreach ($staleFile in $staleFiles) {
        Remove-Item -Path $staleFile.FullName -Force -ErrorAction Stop
    }

    return @(
        $generatedFilePaths |
        Sort-Object -Unique |
        ForEach-Object { Get-Item -Path $_ -ErrorAction Stop } |
        Sort-Object -Property Name
    )
}

#endregion Utility Functions

#region OS-Specific Functions

function ConvertFrom-ProductsXml {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$ProductsXml,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string[]]$ClientTypes,

        [Parameter()]
        [System.Management.Automation.SwitchParameter]$IncludeKey
    )

    function Get-XmlNodePropertyValue {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Node,

            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        if ($Node.PSObject.Properties.Name -contains $Name) {
            return [string]$Node.$Name
        }

        return $null
    }

    $fileNodes = @()
    try {
        $fileNodes = @($ProductsXml.MCT.Catalogs.Catalog.PublishedMedia.Files.File)
    }
    catch {
        return @()
    }

    if (-not $fileNodes -or $fileNodes.Count -lt 1) {
        return @()
    }

    $allowedClientTypes = @(
        $ClientTypes |
        Where-Object { $_ } |
        ForEach-Object { $_.Trim().ToUpperInvariant() } |
        Where-Object { $_ } |
        Select-Object -Unique
    )
    $clientTypeRegex = ($allowedClientTypes | ForEach-Object { [regex]::Escape($_) }) -join '|'

    $candidates = foreach ($node in $fileNodes) {
        if (-not $node) {
            continue
        }

        $fileName = Get-XmlNodePropertyValue -Node $node -Name 'FileName'
        if (-not $fileName -or -not ($fileName -like '*.esd')) {
            continue
        }

        $filePathRaw = Get-XmlNodePropertyValue -Node $node -Name 'FilePath'
        if (-not $filePathRaw) {
            continue
        }
        $filePath = Normalize-DownloadUrl -Url $filePathRaw

        $sizeBytes = $null
        $sizeText = Get-XmlNodePropertyValue -Node $node -Name 'Size'
        if ($sizeText) {
            [long]$parsedSize = 0
            if ([long]::TryParse($sizeText, [ref]$parsedSize)) {
                $sizeBytes = $parsedSize
            }
        }

        $build = $null
        $buildMajor = $null
        $buildUbr = $null
        $buildMatch = [regex]::Match($fileName, '(\d{5})\.(\d+)')
        if ($buildMatch.Success) {
            $build = $buildMatch.Value
            [int]$parsedMajor = 0
            if ([int]::TryParse($buildMatch.Groups[1].Value, [ref]$parsedMajor)) {
                $buildMajor = $parsedMajor
            }

            [int]$parsedUbr = 0
            if ([int]::TryParse($buildMatch.Groups[2].Value, [ref]$parsedUbr)) {
                $buildUbr = $parsedUbr
            }
        }

        $windowsRelease = $null
        if ($buildMajor) {
            $windowsRelease = if ($buildMajor -ge 22000) { 11 } else { 10 }
        }
        $releaseId = Get-WindowsReleaseIdFromBuildMajor -BuildMajor $buildMajor

        $clientType = $null
        if ($clientTypeRegex) {
            $clientTypeMatch = [regex]::Match([string]$filePath, $clientTypeRegex)
            if ($clientTypeMatch.Success) {
                $clientType = $clientTypeMatch.Value.ToUpperInvariant()
            }
        }

        if (-not $clientType -and $clientTypeRegex) {
            $clientTypeMatch = [regex]::Match([string]$fileName, $clientTypeRegex)
            if ($clientTypeMatch.Success) {
                $clientType = $clientTypeMatch.Value.ToUpperInvariant()
            }
        }

        if (-not $clientType) {
            $clientType = Get-ClientTypeFromText -Text ("{0}|{1}" -f [string]$filePath, [string]$fileName)
        }

        $licenseChannel = Get-LicenseChannelFromText -Text ("{0}|{1}" -f [string]$filePath, [string]$fileName)

        $mctId = $null
        if ($node.PSObject.Properties.Name -contains 'id') {
            $parsedMctId = [string]$node.id
            if (-not [string]::IsNullOrWhiteSpace($parsedMctId)) {
                $mctId = $parsedMctId.Trim()
            }
        }

        $isRetailOnly = $null
        $isRetailOnlyText = Get-XmlNodePropertyValue -Node $node -Name 'IsRetailOnly'
        if ($isRetailOnlyText) {
            $isRetailOnly = ($isRetailOnlyText -match '^(?i:true)$')
        }

        $architecture = Normalize-OsArchitecture -Architecture (Get-XmlNodePropertyValue -Node $node -Name 'Architecture')
        $languageCode = Get-XmlNodePropertyValue -Node $node -Name 'LanguageCode'
        $language = Get-XmlNodePropertyValue -Node $node -Name 'Language'
        $edition = Get-XmlNodePropertyValue -Node $node -Name 'Edition'
        $sha1 = Get-XmlNodePropertyValue -Node $node -Name 'Sha1'
        $sha256 = Get-XmlNodePropertyValue -Node $node -Name 'Sha256'

        $item = [ordered]@{
            sourceId = $SourceId
            mctId = $mctId
            clientType = $clientType
            windowsRelease = $windowsRelease
            releaseId = $releaseId
            build = $build
            buildMajor = $buildMajor
            buildUbr = $buildUbr
            architecture = $architecture
            languageCode = $languageCode
            language = $language
            edition = $edition
            fileName = $fileName
            sizeBytes = $sizeBytes
            sha1 = $sha1
            sha256 = $sha256
            isRetailOnly = $isRetailOnly
            licenseChannel = $licenseChannel
            url = $filePath
        }

        if ($IncludeKey) {
            $key = $null
            if ($node.PSObject.Properties.Name -contains 'Key') {
                $key = [string]$node.Key
            }
            $item.key = $key
        }

        [pscustomobject]$item
    }

    $items = @($candidates)
    if ($allowedClientTypes.Count -gt 0) {
        $filtered = @(
            $items |
            Where-Object { $_.clientType -and ($allowedClientTypes -contains $_.clientType.ToUpperInvariant()) }
        )
        if ($filtered.Count -gt 0) {
            $items = $filtered
        }
    }

    return @($items | Sort-Object -Descending -Property @(
            @{ Expression = { $_.windowsRelease } },
            @{ Expression = { $_.buildMajor } },
            @{ Expression = { $_.buildUbr } },
            @{ Expression = { $_.architecture } },
            @{ Expression = { $_.languageCode } },
            @{ Expression = { $_.edition } },
            @{ Expression = { $_.fileName } },
            @{ Expression = { $_.sha1 } },
            @{ Expression = { $_.sha256 } }
        ))
}

function Get-LocalProductsSource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$CatalogRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string[]]$ClientTypes,

        [Parameter()]
        [System.Management.Automation.SwitchParameter]$IncludeKey
    )

    $fileInfo = Get-Item -Path $Path -ErrorAction Stop
    $sourceDefinition = Get-SourceDefinitionFromFileName -FileName $fileInfo.Name

    [xml]$productsXml = Get-Content -Path $fileInfo.FullName -Raw
    $sourceItems = @(
        ConvertFrom-ProductsXml -ProductsXml $productsXml -SourceId $sourceDefinition.id -ClientTypes $ClientTypes -IncludeKey:$IncludeKey
    )

    if ($sourceItems.Count -lt 1) {
        throw ("Source XML '{0}' yielded no matching ESD item after normalization/filtering." -f $fileInfo.Name)
    }

    $matchingBuildItems = @($sourceItems | Where-Object { $_.buildMajor -eq $sourceDefinition.buildMajor })
    if ($matchingBuildItems.Count -lt 1) {
        throw ("Source XML '{0}' does not contain items for build major {1} declared by its file name." -f $fileInfo.Name, $sourceDefinition.buildMajor)
    }

    $representativeItem = $matchingBuildItems |
        Sort-Object -Descending -Property @(
            @{ Expression = { if ($null -eq $_.buildUbr) { -1 } else { $_.buildUbr } } },
            @{ Expression = { $_.build } },
            @{ Expression = { $_.fileName } }
        ) |
        Select-Object -First 1

    $releaseId = if ($representativeItem.releaseId) { [string]$representativeItem.releaseId } else { [string]$sourceDefinition.releaseId }
    if ($releaseId -ne $sourceDefinition.releaseId) {
        throw ("Source XML '{0}' declares release {1} in its file name but contains items for {2}." -f $fileInfo.Name, $sourceDefinition.releaseId, $releaseId)
    }

    $windowsRelease = @($matchingBuildItems | Where-Object { $_.windowsRelease -ne $null } | Select-Object -ExpandProperty windowsRelease -Unique)
    if ($windowsRelease.Count -gt 1) {
        throw ("Source XML '{0}' contains mixed Windows releases." -f $fileInfo.Name)
    }

    $windowsMajor = if ($windowsRelease.Count -eq 1) { [string]$windowsRelease[0] } else { [string]$sourceDefinition.windowsMajor }
    if ($windowsMajor -ne [string]$sourceDefinition.windowsMajor) {
        throw ("Source XML '{0}' declares Windows {1} in its file name but contains Windows {2} items." -f $fileInfo.Name, $sourceDefinition.windowsMajor, $windowsMajor)
    }

    $sourceRelativePath = [System.IO.Path]::GetRelativePath($CatalogRoot, $fileInfo.FullName).Replace('\', '/')
    $sourceXmlSha256 = (Get-FileHash -Path $fileInfo.FullName -Algorithm SHA256).Hash.ToLowerInvariant()

    return [pscustomobject]([ordered]@{
            Source = [pscustomobject]([ordered]@{
                    id = $sourceDefinition.id
                    windowsMajor = $windowsMajor
                    releaseId = $releaseId
                    build = if ($representativeItem.build) { [string]$representativeItem.build } else { [string]$sourceDefinition.buildMajor }
                    buildMajor = $sourceDefinition.buildMajor
                    buildUbr = $representativeItem.buildUbr
                    sourceFile = $sourceRelativePath
                    sourceXmlSha256 = $sourceXmlSha256
                    itemCount = $sourceItems.Count
                })
            Items = @($sourceItems)
        })
}

function Write-OperatingSystemXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Catalog
    )

    $writer = New-CatalogXmlWriter -Path $Path
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('OperatingSystemCatalog')
        $writer.WriteAttributeString('schemaVersion', [string]$Catalog.schemaVersion)
        $writer.WriteAttributeString('generatedAtUtc', [string]$Catalog.generatedAtUtc)

        $writer.WriteStartElement('Source')
        foreach ($property in $Catalog.source.PSObject.Properties) {
            if ($null -eq $property.Value) {
                continue
            }

            $writer.WriteAttributeString([string]$property.Name, [string]$property.Value)
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('Sources')
        foreach ($source in $Catalog.sources) {
            $writer.WriteStartElement('Source')
            foreach ($property in $source.PSObject.Properties) {
                if ($null -eq $property.Value) {
                    continue
                }

                $writer.WriteAttributeString([string]$property.Name, [string]$property.Value)
            }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        $writer.WriteStartElement('Items')
        foreach ($item in $Catalog.items) {
            $writer.WriteStartElement('Item')
            foreach ($property in $item.PSObject.Properties) {
                if ($null -eq $property.Value) {
                    continue
                }

                $writer.WriteStartElement([string]$property.Name)
                $writer.WriteString([string]$property.Value)
                $writer.WriteEndElement()
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

#endregion OS-Specific Functions

#region Main Execution

$startedAt = Get-Date
$generatedAt = (Get-Date).ToUniversalTime()
$generatedAtUtc = $generatedAt.ToString('yyyy-MM-ddTHH:mm:ssZ')
$generatedAtDisplay = $generatedAt.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC'
$resolvedOutputDirectory = Resolve-DirectoryPath -Path $OutputDirectory
$resolvedSourceOutputDirectory = Resolve-DirectoryPath -Path $SourceOutputDirectory
$sevenZipPath = Get-SevenZipCommandPath
$sourceFiles = @(
    Get-ProductsSourceFiles -SourceOutputDirectory $resolvedSourceOutputDirectory -TargetReleases $TargetReleases -ClientTypes $ClientTypes -SevenZipPath $sevenZipPath -IncludeKey:$IncludeKey
)

$sources = @()
$itemsAll = @()

foreach ($sourceFile in $sourceFiles) {
    $sourceResult = Get-LocalProductsSource -Path $sourceFile.FullName -CatalogRoot $resolvedOutputDirectory -ClientTypes $ClientTypes -IncludeKey:$IncludeKey
    $sources += $sourceResult.Source
    $itemsAll += $sourceResult.Items
}

if (-not $itemsAll -or $itemsAll.Count -lt $MinimumItemCount) {
    throw ("Catalog looks unexpectedly small (items={0}, minimum={1})." -f @($itemsAll).Count, $MinimumItemCount)
}

$dedupMap = [ordered]@{}
foreach ($item in $itemsAll) {
    $key = $null
    if ($item.sha256) {
        $key = 'sha256:' + [string]$item.sha256
    }
    elseif ($item.sha1) {
        $key = 'sha1:' + [string]$item.sha1
    }
    else {
        $key = 'url:' + [string]$item.url + '|fn:' + [string]$item.fileName
    }

    if (-not $dedupMap.Contains($key)) {
        $dedupMap[$key] = $item
    }
}

$itemsSorted = @($dedupMap.Values | Sort-Object -Descending -Property @(
        @{ Expression = { $_.windowsRelease } },
        @{ Expression = { $_.buildMajor } },
        @{ Expression = { $_.buildUbr } },
        @{ Expression = { $_.architecture } },
        @{ Expression = { $_.languageCode } },
        @{ Expression = { $_.edition } },
        @{ Expression = { $_.fileName } },
        @{ Expression = { $_.sha1 } },
        @{ Expression = { $_.sha256 } }
    ))

$sourcesSorted = @($sources | Sort-Object -Descending -Property @(
        @{ Expression = { [int]$_.windowsMajor } },
        @{ Expression = { $_.buildMajor } },
        @{ Expression = { if ($null -eq $_.buildUbr) { -1 } else { $_.buildUbr } } },
        @{ Expression = { $_.releaseId } },
        @{ Expression = { $_.id } }
    ))

$relativeSourceDirectory = [System.IO.Path]::GetRelativePath($resolvedOutputDirectory, $resolvedSourceOutputDirectory).Replace('\', '/')

$catalog = [pscustomobject]([ordered]@{
        schemaVersion = 3
        generatedAtUtc = $generatedAtUtc
        source = [pscustomobject]([ordered]@{
                name = 'Foundry Automated OS Catalog Generation'
                directory = $relativeSourceDirectory
            })
        sources = $sourcesSorted
        items = $itemsSorted
    })

$xmlPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'OperatingSystem.xml'
$mdPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'README.md'

Write-OperatingSystemXml -Path $xmlPath -Catalog $catalog

$xmlHash = (Get-FileHash -Path $xmlPath -Algorithm SHA256).Hash.ToLowerInvariant()
$status = if ($itemsSorted.Count -gt 0) { 'SUCCESS' } else { 'WARNING' }
$durationSeconds = [int][Math]::Round(((Get-Date) - $startedAt).TotalSeconds)

$summaryLines = @(
    '# OperatingSystem Summary',
    '',
    '| Metric | Value |',
    '| --- | --- |',
    "| Executed At (UTC) | $generatedAtDisplay |",
    "| Status | $status |",
    "| Source Directory | $relativeSourceDirectory |",
    "| Source Files | $($sourceFiles.Count) |",
    "| Sources Processed | $($sourcesSorted.Count) |",
    "| Items | $($itemsSorted.Count) |",
    "| Duration (Seconds) | $durationSeconds |",
    "| SHA256 XML | $xmlHash |"
)

Write-Utf8NoBomFile -Path $mdPath -Content ($summaryLines -join "`r`n")

[pscustomobject]@{
    XmlPath = $xmlPath
    MarkdownPath = $mdPath
    SourceFiles = $sourceFiles.Count
    Sources = $sourcesSorted.Count
    Items = $itemsSorted.Count
    Status = $status
}

#endregion Main Execution
