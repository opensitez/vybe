# vybe-test: powershell/type_guid_parsing_and_generation/parse_digits_only_32_chars
$str = "d3b07384d11340a2b94fa496c152a3b1"
$g = [guid]::Parse($str)
if ($g.ToString("N") -ne $str) {
    Write-Host "FAIL: digits only parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
