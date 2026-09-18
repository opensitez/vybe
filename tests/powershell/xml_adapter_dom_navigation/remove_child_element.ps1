# vybe-test: powershell/xml_adapter_dom_navigation/remove_child_element
$xml = [xml]'<Basket><Fruit id="1">Apple</Fruit><Fruit id="2">Banana</Fruit></Basket>'

# Remove the first child element
$appleNode = $xml.SelectSingleNode('//Fruit[@id="1"]')
$xml.DocumentElement.RemoveChild($appleNode) | Out-Null

# Only Banana should remain in Fruit
$remainingFruit = $xml.Basket.Fruit.InnerText
$remainingId = $xml.Basket.Fruit.id

if ($remainingFruit -ne "Banana" -or $remainingId -ne "2") {
    Write-Host "FAIL: expected only Banana (id 2) to remain, got fruit='$remainingFruit', id='$remainingId'"
    exit 1
}

Write-Host "PASS"
exit 0
