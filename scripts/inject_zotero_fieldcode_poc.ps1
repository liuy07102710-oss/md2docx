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

$script:BbtEndpoint = if ($env:MD2DOCX_BBT_ENDPOINT) { $env:MD2DOCX_BBT_ENDPOINT } else { "http://127.0.0.1:23119/better-bibtex/json-rpc" }
$script:MissingCitekeys = New-Object System.Collections.Generic.HashSet[string]
$script:ConvertedCitationKeyOccurrences = 0
$script:ConvertedReferenceKeys = New-Object System.Collections.Generic.HashSet[string]

function Invoke-BbtJsonRpc {
  param(
    [string]$Method,
    [object[]]$Params
  )

  $endpoint = $script:BbtEndpoint
  $tmp = Join-Path $env:TEMP ("bbt-" + [guid]::NewGuid().ToString() + ".json")
  $stdoutPath = Join-Path $env:TEMP ("bbt-stdout-" + [guid]::NewGuid().ToString() + ".log")
  $stderrPath = Join-Path $env:TEMP ("bbt-stderr-" + [guid]::NewGuid().ToString() + ".log")
  try {
    $request = @{
      jsonrpc = "2.0"
      method = $Method
      params = $Params
      id = 1
    } | ConvertTo-Json -Depth 20 -Compress

    [System.IO.File]::WriteAllText($tmp, $request, (New-Object System.Text.UTF8Encoding($false)))
    $curlProcess = Start-Process `
      -FilePath "curl.exe" `
      -ArgumentList @(
        "--noproxy", "*",
        "-sS", $endpoint,
        "-H", '"Content-Type: application/json"',
        "-H", '"Accept: application/json"',
        "--data-binary", ('"@' + $tmp + '"')
      ) `
      -NoNewWindow `
      -Wait `
      -PassThru `
      -RedirectStandardOutput $stdoutPath `
      -RedirectStandardError $stderrPath
    $curlExitCode = $curlProcess.ExitCode
    $responseText = if (Test-Path -LiteralPath $stdoutPath) {
      [System.IO.File]::ReadAllText($stdoutPath, [System.Text.Encoding]::UTF8)
    }
    else {
      ""
    }

    if ($curlExitCode -ne 0) {
      $stderrText = if (Test-Path -LiteralPath $stderrPath) {
        [System.IO.File]::ReadAllText($stderrPath, [System.Text.Encoding]::UTF8)
      }
      else {
        ""
      }
      throw @(
        "Step 2 failed: could not connect to Better BibTeX local JSON-RPC.",
        "Endpoint: $endpoint",
        "Check that Zotero is running, Better BibTeX is installed, and the local endpoint is available.",
        "Request error: $stderrText"
      ) -join [Environment]::NewLine
    }
  }
  finally {
    if (Test-Path -LiteralPath $tmp) {
      Remove-Item -LiteralPath $tmp -Force
    }
    if (Test-Path -LiteralPath $stdoutPath) {
      Remove-Item -LiteralPath $stdoutPath -Force
    }
    if (Test-Path -LiteralPath $stderrPath) {
      Remove-Item -LiteralPath $stderrPath -Force
    }
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
    $responsePreviewLength = [Math]::Min(200, $responseText.Length)
    $responsePreview = $responseText.Substring(0, $responsePreviewLength)
    throw @(
      "Better BibTeX local JSON-RPC returned invalid JSON.",
      "Endpoint: $endpoint",
      "Response preview: $responsePreview"
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

function Register-ConvertedCitekeys {
  param([string[]]$Citekeys)

  foreach ($citekey in $Citekeys) {
    if (-not $citekey) {
      continue
    }

    $script:ConvertedCitationKeyOccurrences += 1
    [void]$script:ConvertedReferenceKeys.Add($citekey)
  }
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

function New-WordXmlDocument {
  param([string]$Xml)

  $document = New-Object System.Xml.XmlDocument
  $document.PreserveWhitespace = $true
  $document.LoadXml($Xml)
  return $document
}

function Get-WordNamespaceManager {
  param([System.Xml.XmlDocument]$Document)

  $manager = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
  $manager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
  Write-Output -NoEnumerate $manager
}

function Get-EligibleTextEntries {
  param(
    [System.Xml.XmlNode]$Paragraph,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $entries = New-Object System.Collections.Generic.List[object]
  $offset = 0
  $fieldDepth = 0
  foreach ($run in @($Paragraph.SelectNodes(".//w:r", $NamespaceManager))) {
    $beginCount = @($run.SelectNodes('./w:fldChar[@w:fldCharType="begin"]', $NamespaceManager)).Count
    $endCount = @($run.SelectNodes('./w:fldChar[@w:fldCharType="end"]', $NamespaceManager)).Count
    $fieldDepth += $beginCount

    $insideSimpleField = $null -ne $run.SelectSingleNode("ancestor::w:fldSimple", $NamespaceManager)
    if ($fieldDepth -eq 0 -and -not $insideSimpleField) {
      foreach ($textNode in @($run.SelectNodes("./w:t", $NamespaceManager))) {
        $text = [string]$textNode.InnerText
        if ($text.Length -eq 0) {
          continue
        }
        $entries.Add([pscustomobject]@{
          Run = $run
          TextNode = $textNode
          Start = $offset
          Length = $text.Length
          Text = $text
        })
        $offset += $text.Length
      }
    }

    $fieldDepth = [Math]::Max(0, $fieldDepth - $endCount)
  }

  return $entries.ToArray()
}

function Get-CombinedEntryText {
  param([object[]]$Entries)

  return (($Entries | ForEach-Object { $_.Text }) -join "")
}

function Get-CitationMarkerMatches {
  param([string]$Text)

  $pattern = '@[A-Za-z0-9_.:\-]+\s+\[[^\]]*@[^\]]*\]|\[[^\]]*@[^\]]*\]|@[A-Za-z0-9_.:\-]+'
  return @([regex]::Matches($Text, $pattern))
}

function Get-CitationKeysFromDocument {
  param(
    [System.Xml.XmlDocument]$Document,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $keys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($paragraph in @($Document.SelectNodes("//w:p", $NamespaceManager))) {
    $text = Get-CombinedEntryText -Entries @(Get-EligibleTextEntries -Paragraph $paragraph -NamespaceManager $NamespaceManager)
    foreach ($match in [regex]::Matches($text, '@([A-Za-z0-9_.:\-]+)')) {
      [void]$keys.Add($match.Groups[1].Value)
    }
  }
  return @($keys)
}

function Get-StructurallyUnresolvedMarkers {
  param(
    [System.Xml.XmlDocument]$Document,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $unresolved = New-Object System.Collections.Generic.HashSet[string]
  foreach ($paragraph in @($Document.SelectNodes("//w:p", $NamespaceManager))) {
    $text = Get-CombinedEntryText -Entries @(Get-EligibleTextEntries -Paragraph $paragraph -NamespaceManager $NamespaceManager)
    foreach ($marker in @(Get-CitationMarkerMatches -Text $text)) {
      $markerKeys = @([regex]::Matches($marker.Value, '@([A-Za-z0-9_.:\-]+)') | ForEach-Object { $_.Groups[1].Value })
      $containsMissingKey = $false
      foreach ($key in $markerKeys) {
        if ($script:MissingCitekeys.Contains($key)) {
          $containsMissingKey = $true
          break
        }
      }
      if (-not $containsMissingKey) {
        [void]$unresolved.Add($marker.Value)
      }
    }
  }
  return @($unresolved)
}

function Get-RunPropertiesXml {
  param(
    [System.Xml.XmlNode]$Run,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $properties = $Run.SelectSingleNode("./w:rPr", $NamespaceManager)
  if ($null -eq $properties) {
    return ""
  }
  return $properties.OuterXml
}

function New-RunWithText {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$Text,
    [string]$RunPropertiesXml = ""
  )

  $spaceAttribute = if ($Text -match '^\s|\s$|  ') { ' xml:space="preserve"' } else { "" }
  $fragment = New-Object System.Xml.XmlDocument
  $fragment.PreserveWhitespace = $true
  $fragment.LoadXml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + $RunPropertiesXml + '<w:t' + $spaceAttribute + '>' + (Escape-XmlText -Text $Text) + '</w:t></w:r>')
  return $Document.ImportNode($fragment.DocumentElement, $true)
}

function Convert-FieldXmlToNodes {
  param(
    [System.Xml.XmlDocument]$Document,
    [string]$FieldXml
  )

  $fragment = New-Object System.Xml.XmlDocument
  $fragment.PreserveWhitespace = $true
  $fragment.LoadXml('<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + $FieldXml + '</root>')
  return @($fragment.DocumentElement.ChildNodes | ForEach-Object { $Document.ImportNode($_, $true) })
}

function Get-ReplacementFieldXmlForMarker {
  param(
    [string]$Marker,
    [hashtable]$CitationMap,
    [string]$RunPropertiesXml = ""
  )

  try {
    if ($Marker -match '^@([A-Za-z0-9_.:\-]+)\s+\[(.+)\]$') {
      return Convert-MixedCitationMarkerToFieldXml -Marker $Marker -CitationMap $CitationMap -RunPropertiesXml $RunPropertiesXml
    }

    if ($Marker -match '^\[(?:.*@.*)\]$') {
      return Convert-CitationMarkerToFieldXml -Marker $Marker -CitationMap $CitationMap -RunPropertiesXml $RunPropertiesXml
    }

    if ($Marker -match '^@([A-Za-z0-9_.:\-]+)$') {
      return Convert-NarrativeCitationToFieldXml -Citekey $matches[1] -CitationMap $CitationMap -RunPropertiesXml $RunPropertiesXml
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

function Replace-CitationMarkerInParagraph {
  param(
    [System.Xml.XmlNode]$Paragraph,
    [System.Text.RegularExpressions.Match]$MarkerMatch,
    [hashtable]$CitationMap
  )

  $document = $Paragraph.OwnerDocument
  $namespaceManager = Get-WordNamespaceManager -Document $document
  $entries = @(Get-EligibleTextEntries -Paragraph $Paragraph -NamespaceManager $namespaceManager)
  $matchStart = $MarkerMatch.Index
  $matchEnd = $MarkerMatch.Index + $MarkerMatch.Length
  $affected = @($entries | Where-Object { $_.Start -lt $matchEnd -and ($_.Start + $_.Length) -gt $matchStart })
  if ($affected.Count -eq 0) {
    return $false
  }

  $first = $affected[0]
  $last = $affected[-1]
  $parent = $first.Run.ParentNode
  if (@($affected | Where-Object { -not [object]::ReferenceEquals($_.Run.ParentNode, $parent) }).Count -gt 0) {
    return $false
  }

  $firstRunEntries = @($entries | Where-Object { [object]::ReferenceEquals($_.Run, $first.Run) })
  $lastRunEntries = @($entries | Where-Object { [object]::ReferenceEquals($_.Run, $last.Run) })
  $firstRunStart = $firstRunEntries[0].Start
  $lastRunStart = $lastRunEntries[0].Start
  $firstRunText = Get-CombinedEntryText -Entries $firstRunEntries
  $lastRunText = Get-CombinedEntryText -Entries $lastRunEntries
  $prefixLength = $matchStart - $firstRunStart
  $suffixStart = $matchEnd - $lastRunStart
  $prefix = if ($prefixLength -gt 0) { $firstRunText.Substring(0, $prefixLength) } else { "" }
  $suffix = if ($suffixStart -lt $lastRunText.Length) { $lastRunText.Substring($suffixStart) } else { "" }
  $firstProperties = Get-RunPropertiesXml -Run $first.Run -NamespaceManager $namespaceManager
  $lastProperties = Get-RunPropertiesXml -Run $last.Run -NamespaceManager $namespaceManager
  $fieldXml = Get-ReplacementFieldXmlForMarker -Marker $MarkerMatch.Value -CitationMap $CitationMap -RunPropertiesXml $firstProperties
  if ($null -eq $fieldXml) {
    return $false
  }

  if ($prefix) {
    [void]$parent.InsertBefore((New-RunWithText -Document $document -Text $prefix -RunPropertiesXml $firstProperties), $first.Run)
  }
  foreach ($node in @(Convert-FieldXmlToNodes -Document $document -FieldXml $fieldXml)) {
    [void]$parent.InsertBefore($node, $first.Run)
  }
  if ($suffix) {
    [void]$parent.InsertBefore((New-RunWithText -Document $document -Text $suffix -RunPropertiesXml $lastProperties), $first.Run)
  }

  $runs = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
  foreach ($entry in $affected) {
    if (-not $runs.Contains($entry.Run)) {
      $runs.Add($entry.Run)
    }
  }
  foreach ($run in $runs) {
    [void]$run.ParentNode.RemoveChild($run)
  }
  return $true
}

function Replace-CitationRunsInDocument {
  param(
    [System.Xml.XmlDocument]$Document,
    [hashtable]$CitationMap
  )

  $namespaceManager = Get-WordNamespaceManager -Document $Document
  foreach ($paragraph in @($Document.SelectNodes("//w:p", $namespaceManager))) {
    $entries = @(Get-EligibleTextEntries -Paragraph $paragraph -NamespaceManager $namespaceManager)
    $text = Get-CombinedEntryText -Entries $entries
    $markerMatches = @(Get-CitationMarkerMatches -Text $text)
    for ($index = $markerMatches.Count - 1; $index -ge 0; $index--) {
      [void](Replace-CitationMarkerInParagraph -Paragraph $paragraph -MarkerMatch $markerMatches[$index] -CitationMap $CitationMap)
    }
  }
}

function New-CitationFieldXml {
  param(
    [string]$VisibleText,
    [object[]]$CitationItems,
    [string]$RunPropertiesXml = ""
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
    ('<w:r>' + $RunPropertiesXml + '<w:t>' + (Escape-XmlText -Text $VisibleText) + '</w:t></w:r>'),
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
    [hashtable]$CitationMap,
    [string]$RunPropertiesXml = ""
  )

  $inner = $Marker.Trim('[', ']')
  $parts = @(Parse-BracketCitationParts -InnerText $inner)

  if ($parts.Count -eq 0) {
    return $null
  }

  $citationItems = @()
  $displayParts = New-Object System.Collections.Generic.List[string]
  $convertedCitekeys = New-Object System.Collections.Generic.List[string]

  foreach ($part in $parts) {
    $citekey = $part.Citekey
    if (-not $CitationMap.ContainsKey($citekey)) {
      throw "Citekey not found in Better BibTeX response: $citekey"
    }

    $item = $CitationMap[$citekey]
    $displayParts.Add((Get-NormalCitationPartDisplayText -Item $item -Prefix $part.Prefix -Suffix $part.Suffix))
    $citationItems += (New-CitationItemPayload -Item $item -Prefix $part.Prefix -Suffix $part.Suffix)
    $convertedCitekeys.Add($citekey)
  }

  $plainCitation = "(" + ($displayParts -join "; ") + ")"
  Register-ConvertedCitekeys -Citekeys $convertedCitekeys.ToArray()
  return New-CitationFieldXml -VisibleText $plainCitation -CitationItems $citationItems -RunPropertiesXml $RunPropertiesXml
}

function Convert-NarrativeCitationToFieldXml {
  param(
    [string]$Citekey,
    [hashtable]$CitationMap,
    [string]$RunPropertiesXml = ""
  )

  if (-not $CitationMap.ContainsKey($Citekey)) {
    throw "Citekey not found in Better BibTeX response: $Citekey"
  }

  $item = $CitationMap[$Citekey]
  $displayText = Get-NarrativeCitationDisplayText -Item $item
  Register-ConvertedCitekeys -Citekeys @($Citekey)
  return New-CitationFieldXml -VisibleText $displayText -CitationItems @((New-CitationItemPayload -Item $item)) -RunPropertiesXml $RunPropertiesXml
}

function Convert-MixedCitationMarkerToFieldXml {
  param(
    [string]$Marker,
    [hashtable]$CitationMap,
    [string]$RunPropertiesXml = ""
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
  $convertedCitekeys = New-Object System.Collections.Generic.List[string]

  $parentheticalParts.Add((Get-IssuedYear -Item $narrativeItem))
  $citationItems += (New-CitationItemPayload -Item $narrativeItem)
  $convertedCitekeys.Add($narrativeKey)

  foreach ($part in (Parse-BracketCitationParts -InnerText $inner)) {
    if (-not $CitationMap.ContainsKey($part.Citekey)) {
      throw "Citekey not found in Better BibTeX response: $($part.Citekey)"
    }

    $item = $CitationMap[$part.Citekey]
    $parentheticalParts.Add((Get-NormalCitationPartDisplayText -Item $item -Prefix $part.Prefix -Suffix $part.Suffix))
    $citationItems += (New-CitationItemPayload -Item $item -Prefix $part.Prefix -Suffix $part.Suffix)
    $convertedCitekeys.Add($part.Citekey)
  }

  $visibleText = if ($narrativeAuthorText) {
    "$narrativeAuthorText (" + ($parentheticalParts -join "; ") + ")"
  }
  else {
    "(" + ($parentheticalParts -join "; ") + ")"
  }

  Register-ConvertedCitekeys -Citekeys $convertedCitekeys.ToArray()
  return New-CitationFieldXml -VisibleText $visibleText -CitationItems $citationItems -RunPropertiesXml $RunPropertiesXml
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

try {
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
    $converterScript = Join-Path $PSScriptRoot "md2docx.py"
    if (-not (Test-Path -LiteralPath $converterScript)) {
      throw "md2docx.py not found next to inject_zotero_fieldcode_poc.ps1"
    }

    $pandocArgs = @($converterScript, $InputMarkdown, "-o", $InputDocx)
    if ($ReferenceDocx) {
      $pandocArgs += @("--reference", $ReferenceDocx)
    }

    & python @pandocArgs
    if ($LASTEXITCODE -ne 0) {
      throw "md2docx.py failed while generating the intermediate docx"
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
    $wordDocument = New-WordXmlDocument -Xml $documentXml
    $namespaceManager = Get-WordNamespaceManager -Document $wordDocument
    $citekeys = @(Get-CitationKeysFromDocument -Document $wordDocument -NamespaceManager $namespaceManager)

    if ($citekeys.Count -eq 0) {
      throw "No citation markers like [@citekey] were found in word/document.xml"
    }

    $citationMap = Get-CitationMap -Citekeys $citekeys

    Replace-CitationRunsInDocument -Document $wordDocument -CitationMap $citationMap
    $missingCitekeys = @($script:MissingCitekeys | Sort-Object)
    $hasMissingCitekeys = $missingCitekeys.Count -gt 0
    $structurallyUnresolved = @(Get-StructurallyUnresolvedMarkers -Document $wordDocument -NamespaceManager $namespaceManager)

    if ($script:ConvertedCitationKeyOccurrences -eq 0 -and -not $hasMissingCitekeys) {
      throw "Citation markers were found but no replacements were applied"
    }
    if ($structurallyUnresolved.Count -gt 0) {
      throw "Citation markers remain unconverted: $($structurallyUnresolved -join ', ')"
    }

    [System.IO.File]::WriteAllText((Resolve-Path $documentXmlPath), $wordDocument.OuterXml, $utf8NoBom)

    Remove-ExistingOutputDocx -Path $outputDocxPath

    Compress-Archive -Path (Join-Path $workDir "*") -DestinationPath ($outputDocxPath + ".zip") -Force
    Move-Item -LiteralPath ($outputDocxPath + ".zip") -Destination $outputDocxPath -Force

    if ($hasMissingCitekeys) {
      $countLabel = if ($missingCitekeys.Count -eq 1) { "citekey" } else { "citekeys" }
      $missingSummary = $missingCitekeys -join ", "
      Write-Output "Step 2 warning: $($missingCitekeys.Count) $countLabel not found in Better BibTeX -> $missingSummary"
    }

    Write-Output "Step 2 complete: Zotero field codes injected -> $outputDocxPath"
    Write-Output "Converted $($script:ConvertedCitationKeyOccurrences) citation key occurrences."
    Write-Output "Resolved $($script:ConvertedReferenceKeys.Count) unique references."
  }
  finally {
    if (Test-Path -LiteralPath ($inputDocxPath + ".zip")) {
      Remove-Item -LiteralPath ($inputDocxPath + ".zip") -Force
    }
    if (Test-Path -LiteralPath $workDir) {
      Remove-Item -LiteralPath $workDir -Recurse -Force
    }
  }
}
catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
