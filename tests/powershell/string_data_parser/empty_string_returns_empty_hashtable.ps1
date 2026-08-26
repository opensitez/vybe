# vybe-test: powershell/string_data_parser/empty_string_returns_empty_hashtable
$ht = ConvertFrom-StringData -StringData ""
if ($ht.Count -ne 0) {
    Write-Host "FAIL: Empty string input must return 0-count hashtable"
    exit 1
}
Write-Host "PASS"
exit 0
