[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\Cache\OS'),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$VersionsUri = 'https://worproject.com/dldserv/esd/getversions.php',

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

function Get-FwlinkIdFromUrl {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Url
    )

    if (-not $Url) {
        return $null
    }

    try {
        $match = [regex]::Match($Url, '(?i)(?:\?|&)LinkId=(\d+)')
        if ($match.Success) {
            return $match.Groups[1].Value
        }
    }
    catch {
        Write-Verbose -Message ("Unable to parse fwlink id from URL '{0}': {1}" -f $Url, $_.Exception.Message)
    }

    return $null
}

function ConvertTo-BuildParts {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Build
    )

    $major = $null
    $ubr = $null

    if ($Build) {
        $match = [regex]::Match($Build, '^(\d{5})(?:\.(\d+))?$')
        if ($match.Success) {
            [int]$parsedMajor = 0
            if ([int]::TryParse($match.Groups[1].Value, [ref]$parsedMajor)) {
                $major = $parsedMajor
            }

            if ($match.Groups[2].Success) {
                [int]$parsedUbr = 0
                if ([int]::TryParse($match.Groups[2].Value, [ref]$parsedUbr)) {
                    $ubr = $parsedUbr
                }
            }
        }
    }

    return [pscustomobject]@{
        Major = $major
        Ubr = $ubr
    }
}

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

function Get-WorProjectCatalogSources {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $response = Invoke-RestMethod -Uri $Uri -ErrorAction Stop
    $versionNodes = @($response.productsDb.versions.version)

    $latestFwlinksMap = @{}
    $sources = @()

    foreach ($versionNode in $versionNodes) {
        $windowsMajor = [string]$versionNode.number
        if (-not $windowsMajor) {
            continue
        }

        $latestFwlinkUrl = [string]$versionNode.latestCabLink
        if ($latestFwlinkUrl) {
            $latestFwlinkId = Get-FwlinkIdFromUrl -Url $latestFwlinkUrl
            if ($latestFwlinkId) {
                $latestFwlinksMap[$windowsMajor] = [pscustomobject]([ordered]@{
                        windowsMajor = $windowsMajor
                        fwlinkId = $latestFwlinkId
                        fwlinkUrl = $latestFwlinkUrl
                    })
            }
        }

        $bestByBuildMajor = @{}
        foreach ($releaseNode in @($versionNode.releases.release)) {
            $build = [string]$releaseNode.build
            $date = [string]$releaseNode.date
            $cabUrl = [string]$releaseNode.cabLink

            if (-not $build -or -not $cabUrl) {
                continue
            }

            $parts = ConvertTo-BuildParts -Build $build
            if ($null -eq $parts.Major) {
                continue
            }

            $buildKey = [string]$parts.Major
            $current = $bestByBuildMajor[$buildKey]
            $take = $false

            if (-not $current) {
                $take = $true
            }
            else {
                $currentUbr = $current.buildUbr
                $newUbr = $parts.Ubr
                if ($null -eq $currentUbr) { $currentUbr = -1 }
                if ($null -eq $newUbr) { $newUbr = -1 }

                if ($newUbr -gt $currentUbr) {
                    $take = $true
                }
                elseif ($newUbr -eq $currentUbr -and $date -and ($date -gt $current.date)) {
                    $take = $true
                }
            }

            if ($take) {
                $bestByBuildMajor[$buildKey] = [pscustomobject]([ordered]@{
                        windowsMajor = $windowsMajor
                        build = $build
                        buildMajor = $parts.Major
                        buildUbr = $parts.Ubr
                        date = if ($date) { $date } else { $null }
                        cabUrl = $cabUrl
                    })
            }
        }

        $sources += @($bestByBuildMajor.Values)
    }

    $latestFwlinks = @($latestFwlinksMap.Values | Sort-Object -Descending -Property @(
            @{ Expression = { [int]$_.windowsMajor } },
            @{ Expression = { $_.fwlinkId } },
            @{ Expression = { $_.fwlinkUrl } }
        ))

    $sourcesSorted = @($sources | Sort-Object -Descending -Property @(
            @{ Expression = { [int]$_.windowsMajor } },
            @{ Expression = { $_.buildMajor } },
            @{ Expression = { $_.buildUbr } },
            @{ Expression = { $_.date } },
            @{ Expression = { $_.cabUrl } }
        ))

    return [pscustomobject]@{
        LatestFwlinks = $latestFwlinks
        Sources = $sourcesSorted
    }
}

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

    $clientTypeRegex = ($ClientTypes | ForEach-Object { [regex]::Escape($_) }) -join '|'

    $candidates = foreach ($node in $fileNodes) {
        if (-not $node) {
            continue
        }

        $fileName = Get-XmlNodePropertyValue -Node $node -Name 'FileName'
        if (-not $fileName -or -not ($fileName -like '*.esd')) {
            continue
        }

        $filePath = Get-XmlNodePropertyValue -Node $node -Name 'FilePath'
        if (-not $filePath) {
            continue
        }

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

        $clientType = $null
        if ($clientTypeRegex) {
            $clientTypeMatch = [regex]::Match($filePath, $clientTypeRegex)
            if ($clientTypeMatch.Success) {
                $clientType = $clientTypeMatch.Value
            }
        }

        $mctId = $null
        if ($node.PSObject.Properties.Name -contains 'id') {
            $mctId = [string]$node.id
        }

        $isRetailOnly = $null
        $isRetailOnlyText = Get-XmlNodePropertyValue -Node $node -Name 'IsRetailOnly'
        if ($isRetailOnlyText) {
            $isRetailOnly = ($isRetailOnlyText -match '^(?i:true)$')
        }

        $architecture = Get-XmlNodePropertyValue -Node $node -Name 'Architecture'
        $languageCode = Get-XmlNodePropertyValue -Node $node -Name 'LanguageCode'
        $language = Get-XmlNodePropertyValue -Node $node -Name 'Language'
        $edition = Get-XmlNodePropertyValue -Node $node -Name 'Edition'
        $sha1 = Get-XmlNodePropertyValue -Node $node -Name 'Sha1'

        $item = [ordered]@{
            sourceId = $SourceId
            mctId = $mctId
            clientType = $clientType
            windowsRelease = $windowsRelease
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
            isRetailOnly = $isRetailOnly
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
    if ($clientTypeRegex) {
        $filtered = @($items | Where-Object { $_.url -and [regex]::IsMatch([string]$_.url, $clientTypeRegex) })
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
            @{ Expression = { $_.sha1 } }
        ))
}

function Write-OperatingSystemXml {
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
        $writer.WriteStartElement('OperatingSystemCatalog')
        $writer.WriteAttributeString('schemaVersion', [string]$Catalog.schemaVersion)
        $writer.WriteAttributeString('generatedAtUtc', [string]$Catalog.generatedAtUtc)

        $writer.WriteStartElement('Source')
        $writer.WriteAttributeString('name', [string]$Catalog.source.name)
        $writer.WriteAttributeString('versionsUri', [string]$Catalog.source.versionsUri)
        $writer.WriteEndElement()

        $writer.WriteStartElement('LatestFwlinks')
        foreach ($fwlink in $Catalog.latestFwlinks) {
            $writer.WriteStartElement('LatestFwlink')
            $writer.WriteAttributeString('windowsMajor', [string]$fwlink.windowsMajor)
            $writer.WriteAttributeString('fwlinkId', [string]$fwlink.fwlinkId)
            $writer.WriteAttributeString('fwlinkUrl', [string]$fwlink.fwlinkUrl)
            $writer.WriteEndElement()
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

$startedAt = Get-Date
$generatedAt = (Get-Date).ToUniversalTime()
$generatedAtUtc = $generatedAt.ToString('yyyy-MM-ddTHH:mm:ssZ')
$generatedAtDisplay = $generatedAt.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC'

$expandExe = Join-Path -Path $env:SystemRoot -ChildPath 'System32\expand.exe'
if (-not (Test-Path -Path $expandExe)) {
    throw "Required tool not found: $expandExe"
}

if (-not (Test-Path -Path $OutputDirectory)) {
    $null = New-Item -Path $OutputDirectory -ItemType Directory -Force
}

$catalogSources = Get-WorProjectCatalogSources -Uri $VersionsUri
$latestFwlinks = @($catalogSources.LatestFwlinks)
$sourceInputs = @($catalogSources.Sources)

$sources = @()
$itemsAll = @()
$skippedSources = 0
$skippedSourceDetails = @()

$tempRoot = Join-Path -Path $env:TEMP -ChildPath ('foundry-os-catalog-' + [guid]::NewGuid())
$null = New-Item -Path $tempRoot -ItemType Directory -Force

try {
    foreach ($sourceInput in $sourceInputs) {
        $sourceId = 'Win{0}_{1}' -f [string]$sourceInput.windowsMajor, [string]$sourceInput.build
        $cabUrl = [string]$sourceInput.cabUrl

        if (-not $cabUrl) {
            $reason = 'cabUrl is missing'
            Write-Warning -Message ("Skipping source '{0}' because {1}." -f $sourceId, $reason)
            $skippedSources += 1
            $skippedSourceDetails += [pscustomobject]@{
                sourceId = $sourceId
                reason = $reason
            }
            continue
        }

        $cabPath = Join-Path -Path $tempRoot -ChildPath ("products_{0}.cab" -f $sourceId)
        $xmlPath = Join-Path -Path $tempRoot -ChildPath ("products_{0}.xml" -f $sourceId)

        try {
            Invoke-WebRequest -Uri $cabUrl -OutFile $cabPath -ErrorAction Stop | Out-Null
        }
        catch {
            $reason = "CAB download failed: {0}" -f $_.Exception.Message
            Write-Warning -Message ("Skipping source '{0}' because {1}" -f $sourceId, $reason)
            $skippedSources += 1
            $skippedSourceDetails += [pscustomobject]@{
                sourceId = $sourceId
                reason = $reason
            }
            continue
        }

        $cabSha256 = (Get-FileHash -Path $cabPath -Algorithm SHA256).Hash.ToLowerInvariant()

        & $expandExe '-F:products.xml' $cabPath $xmlPath | Out-Null
        if (-not (Test-Path -Path $xmlPath)) {
            & $expandExe '-F:*.xml' $cabPath $tempRoot | Out-Null
            $xmlCandidates = @(Get-ChildItem -Path $tempRoot -Filter '*.xml' -File | Where-Object { $_.Name -notlike ("products_{0}.xml" -f $sourceId) })
            if ($xmlCandidates.Count -eq 1) {
                Copy-Item -Path $xmlCandidates[0].FullName -Destination $xmlPath -Force
            }
        }

        if (-not (Test-Path -Path $xmlPath)) {
            $reason = 'products.xml was not found after CAB extraction'
            Write-Warning -Message ("Skipping source '{0}' because {1}." -f $sourceId, $reason)
            $skippedSources += 1
            $skippedSourceDetails += [pscustomobject]@{
                sourceId = $sourceId
                reason = $reason
            }
            continue
        }

        $productsXmlSha256 = (Get-FileHash -Path $xmlPath -Algorithm SHA256).Hash.ToLowerInvariant()

        try {
            [xml]$productsXml = Get-Content -Path $xmlPath -Raw
        }
        catch {
            $reason = "products.xml could not be parsed: {0}" -f $_.Exception.Message
            Write-Warning -Message ("Skipping source '{0}' because {1}" -f $sourceId, $reason)
            $skippedSources += 1
            $skippedSourceDetails += [pscustomobject]@{
                sourceId = $sourceId
                reason = $reason
            }
            continue
        }

        $sourceItems = ConvertFrom-ProductsXml -ProductsXml $productsXml -SourceId $sourceId -ClientTypes $ClientTypes -IncludeKey:$IncludeKey
        $itemsAll += @($sourceItems)

        $sources += [pscustomobject]([ordered]@{
                id = $sourceId
                windowsMajor = [string]$sourceInput.windowsMajor
                build = [string]$sourceInput.build
                buildMajor = $sourceInput.buildMajor
                buildUbr = $sourceInput.buildUbr
                date = [string]$sourceInput.date
                cabUrl = $cabUrl
                cabSha256 = $cabSha256
                productsXmlSha256 = $productsXmlSha256
                itemCount = @($sourceItems).Count
            })
    }
}
finally {
    Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not $itemsAll -or $itemsAll.Count -lt $MinimumItemCount) {
    throw ("Catalog looks unexpectedly small (items={0}, minimum={1})." -f @($itemsAll).Count, $MinimumItemCount)
}

$dedupMap = [ordered]@{}
foreach ($item in $itemsAll) {
    $key = $null
    if ($item.sha1) {
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
        @{ Expression = { $_.sha1 } }
    ))

$sourcesSorted = @($sources | Sort-Object -Descending -Property @(
        @{ Expression = { [int]$_.windowsMajor } },
        @{ Expression = { $_.buildMajor } },
        @{ Expression = { $_.buildUbr } },
        @{ Expression = { $_.date } },
        @{ Expression = { $_.id } }
    ))

$catalog = [ordered]@{
    schemaVersion = 2
    generatedAtUtc = $generatedAtUtc
    source = [ordered]@{
        name = 'WORProject MCT Catalogs API'
        versionsUri = $VersionsUri
    }
    latestFwlinks = $latestFwlinks
    sources = $sourcesSorted
    items = $itemsSorted
}

$jsonPath = Join-Path -Path $OutputDirectory -ChildPath 'OperatingSystem.json'
$xmlPath = Join-Path -Path $OutputDirectory -ChildPath 'OperatingSystem.xml'
$mdPath = Join-Path -Path $OutputDirectory -ChildPath 'OperatingSystem.md'

$json = ConvertTo-DeterministicJson -Object $catalog
Write-Utf8NoBomFile -Path $jsonPath -Content $json
Write-OperatingSystemXml -Path $xmlPath -Catalog $catalog

$jsonHash = (Get-FileHash -Path $jsonPath -Algorithm SHA256).Hash.ToLowerInvariant()
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
    "| Sources Processed | $($sourcesSorted.Count) |",
    "| Sources Skipped | $skippedSources |",
    "| Items | $($itemsSorted.Count) |",
    "| Duration (Seconds) | $durationSeconds |",
    "| SHA256 JSON | $jsonHash |",
    "| SHA256 XML | $xmlHash |"
)

if ($skippedSourceDetails.Count -gt 0) {
    $summaryLines += @(
        '',
        '## Sources Not Processed',
        '',
        '| Source | Reason |',
        '| --- | --- |'
    )

    foreach ($skipped in ($skippedSourceDetails | Sort-Object -Property sourceId)) {
        $reasonEscaped = ([string]$skipped.reason) -replace '\|', '\|'
        $summaryLines += "| $([string]$skipped.sourceId) | $reasonEscaped |"
    }
}

Write-Utf8NoBomFile -Path $mdPath -Content ($summaryLines -join "`r`n")

[pscustomobject]@{
    JsonPath = $jsonPath
    XmlPath = $xmlPath
    MarkdownPath = $mdPath
    LatestFwlinks = $latestFwlinks.Count
    Sources = $sourcesSorted.Count
    Items = $itemsSorted.Count
    Status = $status
}
