# vybe-test: powershell/string_normalization_forms/empty_string_normalization
$empty = ""
$norm = $empty.Normalize()
if ($norm -ne "") {
    Write-Host "FAIL: Empty string normalization failed"
    exit 1
}
Write-Host "PASS"
exit 0
