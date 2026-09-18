# vybe-test: powershell/xml_adapter_dom_navigation/text_only_leaf_element_unrolls_to_string
$xml = [xml]'<Configuration><Port>8080</Port></Configuration>'

# In PowerShell XML adapter, an element containing only text (no attributes or child tags)
# directly evaluates to [string] when accessed via dot notation
$port = $xml.Configuration.Port

if (-not ($port -is [string])) {
    Write-Host "FAIL: leaf element was not evaluated as [string], type is: $($port.GetType().FullName)"
    exit 1
}

if ($port -ne "8080") {
    Write-Host "FAIL: expected '8080', got '$port'"
    exit 1
}

Write-Host "PASS"
exit 0
