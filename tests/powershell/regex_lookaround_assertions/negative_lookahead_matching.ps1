# vybe-test: powershell/regex_lookaround_assertions/negative_lookahead_matching
$m = [regex]::Match("100EUR", "\d+(?!USD)")
if (-not $m.Success -or $m.Value -ne "100") {
    Write-Host "FAIL: Negative lookahead failed"
    exit 1
}
Write-Host "PASS"
exit 0
