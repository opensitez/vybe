# vybe-test: powershell/string_data_parser/whitespace_trimming_around_keys_and_values
$str = "   spacedKey   =   spacedValue   "
$ht = ConvertFrom-StringData -StringData $str
if ($ht["spacedKey"] -ne "spacedValue") {
    Write-Host "FAIL: Whitespace trimming around keys/values failed"
    exit 1
}
Write-Host "PASS"
exit 0
