# vybe-test: powershell/strings/string_comparison_ordinal
$result = [string]::Compare("apple", "banana", "Ordinal")
if ($result -ge 0) {
    Write-Host "FAIL: expected negative comparison result"
    exit 1
}
Write-Host "PASS"
exit 0
