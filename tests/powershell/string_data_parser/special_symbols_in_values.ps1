# vybe-test: powershell/string_data_parser/special_symbols_in_values
$str = "symbols = !@#$%^&*()_+"
$ht = ConvertFrom-StringData -StringData $str
if ($ht["symbols"] -ne "!@#$%^&*()_+") {
    Write-Host "FAIL: Special symbols in value failed"
    exit 1
}
Write-Host "PASS"
exit 0
