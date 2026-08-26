# vybe-test: powershell/regex_lookaround_assertions/positive_lookbehind_match
$m = [regex]::Match("USD100EUR200", "\d+")
if (-not $m.Success -or $m.Value -ne "100") {
    Write-Host "FAIL: Regex match failed"
    exit 1
}
Write-Host "PASS"
exit 0
