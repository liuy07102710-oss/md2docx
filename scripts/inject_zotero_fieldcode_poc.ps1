<# 
依赖说明

1. 本脚本依赖本机正在运行的 Zotero + Better BibTeX。
2. 本脚本调用的是 Better BibTeX 提供的本地 JSON-RPC 接口：
   http://127.0.0.1:23119/better-bibtex/json-rpc
3. 本脚本不依赖 Zotero 官方 Web API，也不需要 Zotero Web API key。
4. 因此，即使没有云端 API key，只要本地 Zotero 已启动、Better BibTeX 已安装且 citekey 可查询，脚本就可以工作。
#>

param(
  [string]$InputDocx = "zotero-fieldcode/sample-zotero-full-plain.docx",
  [string]$OutputDocx = "zotero-fieldcode/sample-zotero-full-fieldcoded.docx",
  [string]$InputMarkdown = "",
  [string]$ReferenceDocx = "",
  [switch]$GenerateDocxFirst
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:MissingCitekeys = New-Object System.Collections.Generic.HashSet[string]

function Invoke-BbtJsonRpc {
  param(
    [string]$Method,
    [object[]]$Params
  )

  $endpoint = "http://127.0.0.1:23119/better-bibtex/json-rpc"
  $tmp = Join-Path $env:TEMP ("bbt-" + [guid]::NewGuid().ToString() + ".json")
  try {
    $request = @{
      jsonrpc = "2.0"
      method = $Method
      params = $Params
      id = 1
    } | ConvertTo-Json -Depth 20 -Compress

    Set-Content -LiteralPath $tmp -Value $request -NoNewline -Encoding utf8
    $responseText = & curl.exe --noproxy "*" -sS $endpoint `
      -H "Content-Type: application/json" `
      -H "Accept: application/json" `
      --data-binary "@$tmp" 2>&1
    $curlExitCode = $LASTEXITCODE

    if ($curlExitCode -ne 0) {
      throw @(
        "Failed to connect to Better BibTeX local JSON-RPC.",
        "Endpoint: $endpoint",
        "Check that Zotero is running, Better BibTeX is installed, and the local endpoint is available.",
        "curl: $responseText"
      ) -join [Environment]::NewLine
    }

    if (-not $responseText) {
      throw @(
        "Better BibTeX local JSON-RPC returned an empty response.",
        "Endpoint: $endpoint"
      ) -join [Environment]::NewLine
    }

    try {
      $response = $responseText | ConvertFrom-Json
    }
    catch {
      throw @(
        "Better BibTeX local JSON-RPC returned invalid JSON.",
        "Endpoint: $endpoint",
        "Response: $responseText"
      ) -join [Environment]::NewLine
    }

    if ($null -eq $response) {
      throw @(
        "Better BibTeX local JSON-RPC returned no usable response object.",
        "Endpoint: $endpoint"
      ) -join [Environment]::NewLine
    }

    if ($response.PSObject.Properties.Name -contains "error" -and $null -ne $response.error) {
      throw "Better BibTeX JSON-RPC error: $($response.error.message)"
    }

    return $response.result
  }
  finally {
    if (Test-Path -LiteralPath $tmp) {
      Remove-Item -LiteralPath $tmp -Force
    }
  }
}

function Escape-XmlText {
  param([string]$Text)

  return [System.Security.SecurityElement]::Escape($Text)
}

function Get-CreatorDisplayName {
  param([object]$Creator)

  if ($null -eq $Creator) {
    return "Anon"
  }

  $props = $Creator.PSObject.Properties.Name
  if ($props -contains "literal" -and $Creator.literal) {
    return [string]$Creator.literal
  }
  if ($props -contains "name" -and $Creator.name) {
    return [string]$Creator.name
  }
  if ($props -contains "family" -and $Creator.family) {
    return [string]$Creator.family
  }
  if ($props -contains "lastName" -and $Creator.lastName) {
    return [string]$Creator.lastName
  }
  if ($props -contains "given" -and $Creator.given) {
    return [string]$Creator.given
  }

  return "Anon"
}

function Get-IssuedYear {
  param([pscustomobject]$Item)

  $props = $Item.PSObject.Properties.Name
  if ($props -contains "issued" -and $null -ne $Item.issued -and $Item.issued.PSObject.Properties.Name -contains "date-parts") {
    $parts = @($Item.issued.'date-parts')
    if ($parts.Count -gt 0) {
      $firstPart = @($parts[0])
      if ($firstPart.Count -gt 0) {
        return [string]$firstPart[0]
      }
    }
  }

  if ($props -contains "date" -and $Item.date) {
    if ($Item.date -match '\d{4}') {
      return $matches[0]
    }
  }

  return "n.d."
}

function Get-ItemAuthors {
  param([pscustomobject]$Item)

  $props = $Item.PSObject.Properties.Name
  if ($props -contains "author" -and $null -ne $Item.author) {
    return @($Item.author)
  }
  if ($props -contains "creators" -and $null -ne $Item.creators) {
    return @($Item.creators | Where-Object {
      $creatorProps = $_.PSObject.Properties.Name
      -not ($creatorProps -contains "creatorType") -or $_.creatorType -eq "author"
    })
  }

  return @()
}

function Get-CitationDisplayText {
  param([pscustomobject]$Item)

  $core = Get-ParentheticalCitationCoreText -Item $Item
  return "($core)"
}

function Get-ParentheticalCitationCoreText {
  param([pscustomobject]$Item)

  $authors = @(Get-ItemAuthors -Item $Item)
  $year = Get-IssuedYear -Item $Item

  if ($authors.Count -eq 0) {
    return $year
  }

  $first = Get-CreatorDisplayName -Creator $authors[0]
  $suffix = if ($authors.Count -gt 1) { " et al." } else { "" }
  return "$first$suffix, $year"
}

function Get-NarrativeAuthorDisplayText {
  param([pscustomobject]$Item)

  $authors = @(Get-ItemAuthors -Item $Item)

  if ($authors.Count -eq 0) {
    return ""
  }

  $first = Get-CreatorDisplayName -Creator $authors[0]
  $suffix = if ($authors.Count -gt 1) { " et al." } else { "" }
  return "$first$suffix"
}

function Get-NarrativeCitationDisplayText {
  param([pscustomobject]$Item)

  $authorText = Get-NarrativeAuthorDisplayText -Item $Item
  $year = Get-IssuedYear -Item $Item

  if (-not $authorText) {
    return $year
  }

  return "$authorText ($year)"
}

function Parse-CitationSegment {
  param([string]$Segment)

  $piece = $Segment.Trim()
  if (-not $piece) {
    return $null
  }

  if ($piece -notmatch '^(?<prefix>.*?)@(?<key>[A-Za-z0-9_.:\-]+)(?<suffix>.*)$') {
    throw "Unsupported citation segment: $Segment"
  }

  return [pscustomobject]@{
    Prefix = $matches['prefix'].Trim()
    Citekey = $matches['key']
    Suffix = $matches['suffix'].Trim()
  }
}

function Parse-BracketCitationParts {
  param([string]$InnerText)

  $parts = New-Object System.Collections.Generic.List[object]
  foreach ($part in ($InnerText -split ';')) {
    $parsed = Parse-CitationSegment -Segment $part
    if ($null -ne $parsed) {
      $parts.Add($parsed)
    }
  }

  return $parts.ToArray()
}

function Get-NormalCitationPartDisplayText {
  param(
    [pscustomobject]$Item,
    [string]$Prefix = "",
    [string]$Suffix = ""
  )

  $text = Get-ParentheticalCitationCoreText -Item $Item
  if ($Prefix) {
    $text = "$Prefix $text"
  }
  if ($Suffix) {
    $text = "$text$Suffix"
  }

  return $text
}

function New-CitationItemPayload {
  param(
    [pscustomobject]$Item,
    [string]$Prefix = "",
    [string]$Suffix = ""
  )

  $payload = [ordered]@{
    uris = @($Item.uri)
    uri = $Item.uri
    itemData = Get-ItemData -Item $Item
  }

  if ($Prefix) {
    $payload.prefix = $Prefix
  }
  if ($Suffix) {
    $payload.suffix = $Suffix
  }

  return $payload
}

function Get-ItemData {
  param([pscustomobject]$Item)

  $props = $Item.PSObject.Properties.Name

  return [ordered]@{
    id = if ($props -contains "uri") { $Item.uri } elseif ($props -contains "id") { $Item.id } else { $null }
    type = if ($props -contains "type") { $Item.type } elseif ($props -contains "itemType") { $Item.itemType } else { $null }
    title = if ($props -contains "title") { $Item.title } else { $null }
    author = if ($props -contains "author") { $Item.author } else { $null }
    issued = if ($props -contains "issued") { $Item.issued } else { $null }
    "container-title" = if ($props -contains "container-title") { $Item.'container-title' } elseif ($props -contains "publicationTitle") { $Item.publicationTitle } else { $null }
    volume = if ($props -contains "volume") { $Item.volume } else { $null }
    issue = if ($props -contains "issue") { $Item.issue } else { $null }
    page = if ($props -contains "page") { $Item.page } elseif ($props -contains "pages") { $Item.pages } else { $null }
    DOI = if ($props -contains "DOI") { $Item.DOI } else { $null }
    URL = if ($props -contains "URL") { $Item.URL } elseif ($props -contains "url") { $Item.url } else { $null }
  }
}

function Get-CitationKeysFromXml {
  param([string]$DocumentXml)

  $keys = New-Object System.Collections.Generic.HashSet[string]

  foreach ($match in [regex]::Matches($DocumentXml, '@([A-Za-z0-9_.:\-]+)')) {
    [void]$keys.Add($match.Groups[1].Value)
  }

  return @($keys)
}

function Get-TextRunMatches {
  param([string]$Xml)

  return @([regex]::Matches($Xml, '<w:r><w:t(?: xml:space="preserve")?>(.*?)</w:t></w:r>'))
}

function Get-CombinedRunText {
  param(
    [object[]]$RunMatches,
    [int]$StartIndex,
    [int]$RunCount
  )

  $builder = New-Object System.Text.StringBuilder
  for ($offset = 0; $offset -lt $RunCount; $offset++) {
    [void]$builder.Append($RunMatches[$StartIndex + $offset].Groups[1].Value)
  }

  return $builder.ToString()
}

function Test-RunSpanIsAdjacent {
  param(
    [string]$Xml,
    [object[]]$RunMatches,
    [int]$StartIndex,
    [int]$RunCount
  )

  for ($offset = 0; $offset -lt ($RunCount - 1); $offset++) {
    $current = $RunMatches[$StartIndex + $offset]
    $next = $RunMatches[$StartIndex + $offset + 1]
    $betweenStart = $current.Index + $current.Length
    $betweenLength = $next.Index - $betweenStart
    if ($betweenLength -lt 0) {
      return $false
    }

    $between = $Xml.Substring($betweenStart, $betweenLength)
    if ($between.Length -gt 0) {
      return $false
    }
  }

  return $true
}

function Get-ReplacementFieldXmlForMarker {
  param(
    [string]$Marker,
    [hashtable]$CitationMap
  )

  try {
    if ($Marker -match '^@([A-Za-z0-9_.:\-]+)\s+\[(.+)\]$') {
      return Convert-MixedCitationMarkerToFieldXml -Marker $Marker -CitationMap $CitationMap
    }

    if ($Marker -match '^\[(?:.*@.*)\]$') {
      return Convert-CitationMarkerToFieldXml -Marker $Marker -CitationMap $CitationMap
    }

    if ($Marker -match '^@([A-Za-z0-9_.:\-]+)$') {
      return Convert-NarrativeCitationToFieldXml -Citekey $matches[1] -CitationMap $CitationMap
    }
  }
  catch {
    $message = $_.Exception.Message
    if ($message -match '^Citekey not found in Better BibTeX response: (.+)$') {
      [void]$script:MissingCitekeys.Add($matches[1])
      return $null
    }

    throw
  }

  return $null
}

function Replace-CitationRunsInParagraph {
  param(
    [string]$ParagraphXml,
    [hashtable]$CitationMap
  )

  $runMatches = @(Get-TextRunMatches -Xml $ParagraphXml)
  if ($runMatches.Count -eq 0) {
    return $ParagraphXml
  }

  $result = New-Object System.Text.StringBuilder
  $cursor = 0
  $index = 0

  while ($index -lt $runMatches.Count) {
    $matched = $false

    $maxSpanLength = [Math]::Min(12, $runMatches.Count - $index)
    foreach ($spanLength in ($maxSpanLength..1)) {
      if (($index + $spanLength) -gt $runMatches.Count) {
        continue
      }
      if (-not (Test-RunSpanIsAdjacent -Xml $ParagraphXml -RunMatches $runMatches -StartIndex $index -RunCount $spanLength)) {
        continue
      }

      $combinedText = Get-CombinedRunText -RunMatches $runMatches -StartIndex $index -RunCount $spanLength
      $fieldXml = Get-ReplacementFieldXmlForMarker -Marker $combinedText -CitationMap $CitationMap
      if ($null -eq $fieldXml) {
        continue
      }

      $start = $runMatches[$index].Index
      $end = $runMatches[$index + $spanLength - 1].Index + $runMatches[$index + $spanLength - 1].Length
      [void]$result.Append($ParagraphXml.Substring($cursor, $start - $cursor))
      [void]$result.Append($fieldXml)
      $cursor = $end
      $index += $spanLength
      $matched = $true
      break
    }

    if (-not $matched) {
      $run = $runMatches[$index]
      $end = $run.Index + $run.Length
      [void]$result.Append($ParagraphXml.Substring($cursor, $end - $cursor))
      $cursor = $end
      $index += 1
    }
  }

  if ($cursor -lt $ParagraphXml.Length) {
    [void]$result.Append($ParagraphXml.Substring($cursor))
  }

  return $result.ToString()
}

function Replace-CitationRunsInDocumentXml {
  param(
    [string]$DocumentXml,
    [hashtable]$CitationMap
  )

  return [regex]::Replace(
    $DocumentXml,
    '<w:p\b[^>]*>.*?</w:p>',
    [System.Text.RegularExpressions.MatchEvaluator]{
      param($match)
      return Replace-CitationRunsInParagraph -ParagraphXml $match.Value -CitationMap $CitationMap
    },
    [System.Text.RegularExpressions.RegexOptions]::Singleline
  )
}

function New-CitationFieldXml {
  param(
    [string]$VisibleText,
    [object[]]$CitationItems
  )

  $payload = [ordered]@{
    citationID = "codex-" + [guid]::NewGuid().ToString("N")
    properties = [ordered]@{
      formattedCitation = $VisibleText
      plainCitation = $VisibleText
      noteIndex = 0
    }
    citationItems = $CitationItems
    schema = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
  }

  $fieldCode = " ADDIN ZOTERO_ITEM CSL_CITATION " + ($payload | ConvertTo-Json -Depth 50 -Compress)
  return @(
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    ('<w:r><w:instrText xml:space="preserve">' + (Escape-XmlText -Text $fieldCode) + '</w:instrText></w:r>'),
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    ('<w:r><w:t>' + (Escape-XmlText -Text $VisibleText) + '</w:t></w:r>'),
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  ) -join ""
}

function Get-CitationMap {
  param([string[]]$Citekeys)

  if ($Citekeys.Count -eq 0) {
    return @{}
  }

  $result = Invoke-BbtJsonRpc -Method "item.pandoc_filter" -Params @($Citekeys, $false)
  if ($null -eq $result -or $null -eq $result.items) {
    throw "Better BibTeX returned no items for the requested citekeys"
  }

  $map = @{}
  foreach ($prop in $result.items.PSObject.Properties) {
    $map[$prop.Name] = $prop.Value
  }

  return $map
}

function Convert-CitationMarkerToFieldXml {
  param(
    [string]$Marker,
    [hashtable]$CitationMap
  )

  $inner = $Marker.Trim('[', ']')
  $parts = @(Parse-BracketCitationParts -InnerText $inner)

  if ($parts.Count -eq 0) {
    return $null
  }

  $citationItems = @()
  $displayParts = New-Object System.Collections.Generic.List[string]

  foreach ($part in $parts) {
    $citekey = $part.Citekey
    if (-not $CitationMap.ContainsKey($citekey)) {
      throw "Citekey not found in Better BibTeX response: $citekey"
    }

    $item = $CitationMap[$citekey]
    $displayParts.Add((Get-NormalCitationPartDisplayText -Item $item -Prefix $part.Prefix -Suffix $part.Suffix))
    $citationItems += (New-CitationItemPayload -Item $item -Prefix $part.Prefix -Suffix $part.Suffix)
  }

  $plainCitation = "(" + ($displayParts -join "; ") + ")"
  return New-CitationFieldXml -VisibleText $plainCitation -CitationItems $citationItems
}

function Convert-NarrativeCitationToFieldXml {
  param(
    [string]$Citekey,
    [hashtable]$CitationMap
  )

  if (-not $CitationMap.ContainsKey($Citekey)) {
    throw "Citekey not found in Better BibTeX response: $Citekey"
  }

  $item = $CitationMap[$Citekey]
  $displayText = Get-NarrativeCitationDisplayText -Item $item
  return New-CitationFieldXml -VisibleText $displayText -CitationItems @((New-CitationItemPayload -Item $item))
}

function Convert-MixedCitationMarkerToFieldXml {
  param(
    [string]$Marker,
    [hashtable]$CitationMap
  )

  if ($Marker -notmatch '^@([A-Za-z0-9_.:\-]+)\s+\[(.+)\]$') {
    throw "Unsupported mixed citation marker: $Marker"
  }

  $narrativeKey = $matches[1]
  $inner = $matches[2]

  if (-not $CitationMap.ContainsKey($narrativeKey)) {
    throw "Citekey not found in Better BibTeX response: $narrativeKey"
  }

  $narrativeItem = $CitationMap[$narrativeKey]
  $narrativeAuthorText = Get-NarrativeAuthorDisplayText -Item $narrativeItem
  $parentheticalParts = New-Object System.Collections.Generic.List[string]
  $citationItems = @()

  $parentheticalParts.Add((Get-IssuedYear -Item $narrativeItem))
  $citationItems += (New-CitationItemPayload -Item $narrativeItem)

  foreach ($part in (Parse-BracketCitationParts -InnerText $inner)) {
    if (-not $CitationMap.ContainsKey($part.Citekey)) {
      throw "Citekey not found in Better BibTeX response: $($part.Citekey)"
    }

    $item = $CitationMap[$part.Citekey]
    $parentheticalParts.Add((Get-NormalCitationPartDisplayText -Item $item -Prefix $part.Prefix -Suffix $part.Suffix))
    $citationItems += (New-CitationItemPayload -Item $item -Prefix $part.Prefix -Suffix $part.Suffix)
  }

  $visibleText = if ($narrativeAuthorText) {
    "$narrativeAuthorText (" + ($parentheticalParts -join "; ") + ")"
  }
  else {
    "(" + ($parentheticalParts -join "; ") + ")"
  }

  return New-CitationFieldXml -VisibleText $visibleText -CitationItems $citationItems
}

function Ensure-ParentDirectory {
  param([string]$Path)

  $parent = Split-Path -Path $Path -Parent
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
}

function Remove-ExistingOutputDocx {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    return
  }

  try {
    Remove-Item -LiteralPath $Path -Force
  }
  catch {
    throw @(
      "Cannot overwrite output docx because it is in use or locked.",
      "Output: $Path",
      "Close the file in Word or any preview tool, then run the script again."
    ) -join [Environment]::NewLine
  }
}

$cwd = Get-Location
$inputDocxPath = if ([System.IO.Path]::IsPathRooted($InputDocx)) { $InputDocx } else { Join-Path $cwd $InputDocx }
$outputDocxPath = if ([System.IO.Path]::IsPathRooted($OutputDocx)) { $OutputDocx } else { Join-Path $cwd $OutputDocx }
$workDir = Join-Path $env:TEMP ("docx-zotero-" + [guid]::NewGuid().ToString())
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($GenerateDocxFirst) {
  if (-not $InputMarkdown) {
    throw "InputMarkdown is required when -GenerateDocxFirst is used"
  }

  Ensure-ParentDirectory -Path $inputDocxPath
  $pandocArgs = @($InputMarkdown, "-o", $InputDocx)
  if ($ReferenceDocx) {
    $pandocArgs += @("--reference-doc", $ReferenceDocx)
  }

  & pandoc @pandocArgs
  if ($LASTEXITCODE -ne 0) {
    throw "pandoc failed while generating the intermediate docx"
  }
}

if (-not (Test-Path -LiteralPath $inputDocxPath)) {
  throw "Input docx not found: $inputDocxPath"
}

if (Test-Path -LiteralPath $workDir) {
  Remove-Item -LiteralPath $workDir -Recurse -Force
}

try {
  Copy-Item -LiteralPath $inputDocxPath -Destination ($inputDocxPath + ".zip") -Force
  Expand-Archive -LiteralPath ($inputDocxPath + ".zip") -DestinationPath $workDir -Force
  Remove-Item -LiteralPath ($inputDocxPath + ".zip") -Force

  $documentXmlPath = Join-Path $workDir "word/document.xml"
  $documentXml = [System.IO.File]::ReadAllText((Resolve-Path $documentXmlPath), $utf8NoBom)
  $citekeys = @(Get-CitationKeysFromXml -DocumentXml $documentXml)

  if ($citekeys.Count -eq 0) {
    throw "No citation markers like [@citekey] were found in word/document.xml"
  }

  $citationMap = Get-CitationMap -Citekeys $citekeys

  $updatedXml = Replace-CitationRunsInDocumentXml -DocumentXml $documentXml -CitationMap $citationMap
  $missingCitekeys = @($script:MissingCitekeys | Sort-Object)
  $hasMissingCitekeys = $missingCitekeys.Count -gt 0

  if ($updatedXml -eq $documentXml -and -not $hasMissingCitekeys) {
    throw "Citation markers were found but no replacements were applied"
  }

  [System.IO.File]::WriteAllText((Resolve-Path $documentXmlPath), $updatedXml, $utf8NoBom)

  Remove-ExistingOutputDocx -Path $outputDocxPath

  Compress-Archive -Path (Join-Path $workDir "*") -DestinationPath ($outputDocxPath + ".zip") -Force
  Move-Item -LiteralPath ($outputDocxPath + ".zip") -Destination $outputDocxPath -Force

  if ($hasMissingCitekeys) {
    $warningMessage = @(
      "Some citekeys were not found in Better BibTeX and were left unchanged in the output document.",
      "Missing citekeys:",
      ($missingCitekeys | ForEach-Object { " - $_" })
    ) -join [Environment]::NewLine
    Write-Warning $warningMessage
  }
}
finally {
  if (Test-Path -LiteralPath ($inputDocxPath + ".zip")) {
    Remove-Item -LiteralPath ($inputDocxPath + ".zip") -Force
  }
  if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
  }
}
