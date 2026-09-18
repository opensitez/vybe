# vybe-test: powershell/xml_adapter_dom_navigation/modify_attribute_via_property_assignment
$xml = [xml]'<Server name="srv1" status="offline"/>'

# Modifying attribute via property assignment updates the attribute in the underlying DOM
$xml.Server.status = "online"

$newStatus = $xml.Server.status
$outer = $xml.OuterXml

if ($newStatus -ne "online") {
    Write-Host "FAIL: attribute readback expected 'online', got '$newStatus'"
    exit 1
}

if (-not ($outer -match 'status="online"')) {
    Write-Host "FAIL: OuterXml did not reflect updated attribute value, got: $outer"
    exit 1
}

Write-Host "PASS"
exit 0
