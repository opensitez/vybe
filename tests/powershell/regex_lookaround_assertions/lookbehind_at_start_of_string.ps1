# vybe-test: powershell/regex_lookaround_assertions/lookbehind_at_start_of_string
$str = "start:data"
$matched = $str -match "(?<=^start:)data"
if (-not $matched -or $Matches[0] -ne "data") {
    Write-Host "FAIL: Lookbehind at start of string failed"
    exit 1
}
Write-Host "PASS"
exit 0
