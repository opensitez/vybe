# vybe-test: powershell/regex_lookaround_assertions/lookbehind_fixed_width_requirement_check
$str = "file_test.txt"
$matched = $str -match "(?<=file_)\w+"
if (-not $matched -or $Matches[0] -ne "test") {
    Write-Host "FAIL: Fixed-width lookbehind match failed"
    exit 1
}
Write-Host "PASS"
exit 0
