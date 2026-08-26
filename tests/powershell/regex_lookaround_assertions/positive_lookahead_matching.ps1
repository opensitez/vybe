# vybe-test: powershell/regex_lookaround_assertions/positive_lookahead_matching
$m = [regex]::Match("100USD", "\d+(?=USD)")
if (-not $m.Success -or $m.Value -ne "100") {
    Write-Host "FAIL: Positive lookahead failed"
    exit 1
}
Write-Host "PASS"
exit 0
