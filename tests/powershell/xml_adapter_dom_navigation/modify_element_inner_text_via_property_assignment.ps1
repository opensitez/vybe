# vybe-test: powershell/xml_adapter_dom_navigation/modify_element_inner_text_via_property_assignment
$xml = [xml]'<Settings><Theme>Light</Theme></Settings>'

# Modifying text via property assignment updates the DOM node
$xml.Settings.Theme = "Dark"

$newVal = $xml.Settings.Theme
$outer = $xml.OuterXml

if ($newVal -ne "Dark") {
    Write-Host "FAIL: property readback expected 'Dark', got '$newVal'"
    exit 1
}

if (-not ($outer -match '<Theme>Dark</Theme>')) {
    Write-Host "FAIL: OuterXml did not reflect updated inner text, got: $outer"
    exit 1
}

Write-Host "PASS"
exit 0
