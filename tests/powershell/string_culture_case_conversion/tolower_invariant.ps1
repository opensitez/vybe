# vybe-test: powershell/string_culture_case_conversion/tolower_invariant
$str = "HELLO WORLD"
$lower = $str.ToLowerInvariant()
if ($lower -ne "hello world") {
    Write-Host "FAIL: ToLowerInvariant failed"
    exit 1
}
Write-Host "PASS"
exit 0
