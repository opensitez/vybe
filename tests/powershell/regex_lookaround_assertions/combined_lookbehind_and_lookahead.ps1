# vybe-test: powershell/regex_lookaround_assertions/combined_lookbehind_and_lookahead
$str = "<title>My Webpage</title>"
$matched = $str -match "(?<=<title>).*?(?=</title>)"
if (-not $matched -or $Matches[0] -ne "My Webpage") {
    Write-Host "FAIL: Combined lookbehind and lookahead failed, got $($Matches[0])"
    exit 1
}
Write-Host "PASS"
exit 0
