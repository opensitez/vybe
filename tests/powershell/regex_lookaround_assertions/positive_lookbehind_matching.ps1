# vybe-test: powershell/regex_lookaround_assertions/positive_lookbehind_matching
$m = [regex]::Match("USD100", "(?<=USD)\d+")
if (-not $m.Success -or $m.Value -ne "100") {
    Write-Host "FAIL: Positive lookbehind failed"
    exit 1
}
Write-Host "PASS"
exit 0
