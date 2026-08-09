# vybe-test: powershell/null_coalescing_assignment/null_assignment_empty_string_not_null
$str = ""
$str ??= "FallbackString"
if ($str -ne "") {
    Write-Host "FAIL: empty string should NOT be treated as null by ??=, got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
