# vybe-test: powershell/xml_adapter_dom_navigation/preserve_whitespace_toggle
$rawXml = "<Doc>`n    <IndentedTag>Value</IndentedTag>`n</Doc>"

# Configuring PreserveWhitespace = $true before LoadXml preserves exact newline and spacing
$doc = [System.Xml.XmlDocument]::new()
$doc.PreserveWhitespace = $true
$doc.LoadXml($rawXml)

$outer = $doc.OuterXml

if (-not ($outer -match "`n    <IndentedTag>")) {
    Write-Host "FAIL: PreserveWhitespace did not retain indentation formatting, got: $outer"
    exit 1
}

if ($doc.Doc.IndentedTag -ne "Value") {
    Write-Host "FAIL: property readback failed, got: $($doc.Doc.IndentedTag)"
    exit 1
}

Write-Host "PASS"
exit 0
