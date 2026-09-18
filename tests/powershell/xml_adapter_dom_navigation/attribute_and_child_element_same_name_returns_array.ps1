# vybe-test: powershell/xml_adapter_dom_navigation/attribute_and_child_element_same_name_returns_array
# When an element has both an attribute 'category' and a child element '<category>',
# PowerShell's XML adapter bundles both into an array [System.Object[]]
$xml = [xml]'<Item category="AttrCategory"><category>ElemCategory</category></Item>'

$cat = $xml.Item.category

if (-not ($cat -is [array])) {
    Write-Host "FAIL: expected array when attribute and child share name, type is: $($cat.GetType().FullName)"
    exit 1
}

if ($cat.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements in bundled array, got $($cat.Count)"
    exit 1
}

if ($cat[0] -ne "AttrCategory" -or $cat[1] -ne "ElemCategory") {
    Write-Host "FAIL: expected @('AttrCategory', 'ElemCategory'), got @($($cat -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
