# vybe-test: powershell/xml_adapter_dom_navigation/attribute_access_via_dot_notation
$xml = [xml]'<Root><Element id="E101" status="Active" priority="1"/></Root>'

# Attributes on XML elements are exposed directly as properties through the adapter
$elem = $xml.Root.Element
$id = $elem.id
$status = $elem.status
$priority = $elem.priority

if ($id -ne "E101" -or $status -ne "Active" -or $priority -ne "1") {
    Write-Host "FAIL: attribute dot access failed, got id='$id', status='$status', priority='$priority'"
    exit 1
}

Write-Host "PASS"
exit 0
