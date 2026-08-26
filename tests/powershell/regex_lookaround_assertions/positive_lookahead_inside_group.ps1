# vybe-test: powershell/regex_lookaround_assertions/positive_lookahead_inside_group
$str = "item123end"
$matched = $str -match "item(\d+(?=end))"
if (-not $matched -or $Matches[1] -ne "123") {
    Write-Host "FAIL: Lookahead inside capture group failed"
    exit 1
}
Write-Host "PASS"
exit 0
