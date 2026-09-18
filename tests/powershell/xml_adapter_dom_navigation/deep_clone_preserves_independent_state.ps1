# vybe-test: powershell/xml_adapter_dom_navigation/deep_clone_preserves_independent_state
$xml = [xml]'<Inventory><Product id="P1"><Qty>50</Qty></Product></Inventory>'

# CloneNode($true) performs a deep clone including all child elements and attributes
$clone = $xml.CloneNode($true)

# Mutating the clone must not alter the original
$clone.Inventory.Product.Qty = "999"

if ($xml.Inventory.Product.Qty -ne "50") {
    Write-Host "FAIL: mutating clone altered original document, expected original Qty 50, got: $($xml.Inventory.Product.Qty)"
    exit 1
}

if ($clone.Inventory.Product.Qty -ne "999") {
    Write-Host "FAIL: clone did not retain its mutated value, got: $($clone.Inventory.Product.Qty)"
    exit 1
}

Write-Host "PASS"
exit 0
