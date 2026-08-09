# vybe-test: powershell/type_converters/type_converter_char_int
$c = [char]65
if ($c -ne 'A') {
    Write-Host "FAIL: int 65 to char conversion expected 'A', got '$c'"
    exit 1
}
Write-Host "PASS"
exit 0
