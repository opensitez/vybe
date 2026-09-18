# vybe-test: powershell/xml_adapter_dom_navigation/select_xml_pattern_and_node_metadata
$xml = [xml]'<Library name="Central">
  <Book isbn="978-01">PowerShell In Action</Book>
</Library>'

# Select-Xml emits SelectXmlInfo objects exposing .Node, .Pattern, and document context
$result = Select-Xml -Xml $xml -XPath '//Book'

if ($null -eq $result) {
    Write-Host "FAIL: Select-Xml returned `$null"
    exit 1
}

if ($result.Pattern -ne "//Book") {
    Write-Host "FAIL: expected Pattern '//Book', got '$($result.Pattern)'"
    exit 1
}

if ($result.Node.InnerText -ne "PowerShell In Action") {
    Write-Host "FAIL: expected InnerText 'PowerShell In Action', got '$($result.Node.InnerText)'"
    exit 1
}

if ($result.Node.Attributes["isbn"].Value -ne "978-01") {
    Write-Host "FAIL: expected isbn '978-01', got '$($result.Node.Attributes['isbn'].Value)'"
    exit 1
}

Write-Host "PASS"
exit 0
