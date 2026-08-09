# vybe-test: powershell/type_accelerators/type_accelerator_xml
$xml = [xml]"<root><item id='1'>Hello</item></root>"
if ($xml.root.item.id -ne "1") {
    Write-Host "FAIL: expected '1', got $($xml.root.item.id)"
    exit 1
}
if ($xml.root.item.InnerText -ne "Hello") {
    Write-Host "FAIL: expected 'Hello', got $($xml.root.item.InnerText)"
    exit 1
}
Write-Host "PASS"
exit 0
