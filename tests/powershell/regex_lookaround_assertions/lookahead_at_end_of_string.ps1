# vybe-test: powershell/regex_lookaround_assertions/lookahead_at_end_of_string
$str = "data;end"
$matched = $str -match "data(?=;end$)"
if (-not $matched -or $Matches[0] -ne "data") {
    Write-Host "FAIL: Lookahead at end of string failed"
    exit 1
}
Write-Host "PASS"
exit 0
