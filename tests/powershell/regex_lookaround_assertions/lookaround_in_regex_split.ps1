# vybe-test: powershell/regex_lookaround_assertions/lookaround_in_regex_split
$m = [regex]::Match("USD100EUR200", "\d+")
if (-not $m.Success -or $m.Value -ne "100") {
    Write-Host "FAIL: Regex match failed"
    exit 1
}
Write-Host "PASS"
exit 0
