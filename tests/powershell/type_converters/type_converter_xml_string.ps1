# vybe-test: powershell/type_converters/type_converter_xml_string
$xml = [xml]"<node>content</node>"
if ($xml.node -ne "content") {
    Write-Host "FAIL: string to [xml] conversion expected content, got $($xml.node)"
    exit 1
}
Write-Host "PASS"
exit 0
