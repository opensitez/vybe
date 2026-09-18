# vybe-test: powershell/xml_adapter_dom_navigation/repeated_child_elements_unrolled_as_array
$xml = [xml]'<Catalog><Item>First</Item><Item>Second</Item><Item>Third</Item></Catalog>'

# Repeated sibling elements sharing the same tag are automatically adapted as an array
$items = $xml.Catalog.Item

if (-not ($items -is [array])) {
    Write-Host "FAIL: repeated child elements were not unrolled as an array"
    exit 1
}

if ($items.Count -ne 3) {
    Write-Host "FAIL: expected 3 items, got $($items.Count)"
    exit 1
}

if ($items[0] -ne "First" -or $items[1] -ne "Second" -or $items[2] -ne "Third") {
    Write-Host "FAIL: expected values @('First', 'Second', 'Third'), got @($($items -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
