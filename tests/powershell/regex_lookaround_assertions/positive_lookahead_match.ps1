# vybe-test: powershell/regex_lookaround_assertions/positive_lookahead_match
$str = "100 USD and 200 EUR"
$matched = $str -match "\d+(?=\s*USD)"
if (-not $matched -or $Matches[0] -ne "100") {
    Write-Host "FAIL: Positive lookahead (?=USD) failed, got $($Matches[0])"
    exit 1
}
Write-Host "PASS"
exit 0
