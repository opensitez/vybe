# vybe-test: powershell/xml_adapter_dom_navigation/set_attribute_method_on_document_element
$xml = [xml]'<Service name="AuthService"/>'

# SetAttribute method dynamically assigns new attributes
$xml.DocumentElement.SetAttribute("version", "3.2.1")
$xml.DocumentElement.SetAttribute("enabled", "true")

$ver = $xml.DocumentElement.version
$enabled = $xml.DocumentElement.enabled
$outer = $xml.OuterXml

if ($ver -ne "3.2.1" -or $enabled -ne "true") {
    Write-Host "FAIL: SetAttribute failed to expose attribute via dot syntax, got ver='$ver', enabled='$enabled'"
    exit 1
}

if (-not ($outer -match 'version="3.2.1"' -and $outer -match 'enabled="true"')) {
    Write-Host "FAIL: attributes missing from OuterXml: $outer"
    exit 1
}

Write-Host "PASS"
exit 0
