# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_with_empty_string_is_not_null
$str = ""
# Empty string is NOT null, so ?? must return empty string
$res = $str ?? "Default"
if ($res -ne "") {
    Write-Host "FAIL: Empty string should not be coalesced, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
