# vybe-test: powershell/type_guid_parsing_and_generation/parse_braced_format
$str = "{d3b07384-d113-40a2-b94f-a496c152a3b1}"
$g = [guid]::Parse($str)
if ($g.ToString("B") -ne $str) {
    Write-Host "FAIL: braced parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
