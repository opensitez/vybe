# vybe-test: powershell/regex_lookaround_assertions/negative_lookbehind_matching
$m = [regex]::Match("EUR100", "(?<!USD)\d+")
if (-not $m.Success -or $m.Value -ne "100") {
    Write-Host "FAIL: Negative lookbehind failed"
    exit 1
}
Write-Host "PASS"
exit 0
