# vybe-test: powershell/regex_lookaround_assertions/lookahead_zero_width_replace
$str = "abc"
$res = $str -replace "(?=[b])", "-"
if ($res -ne "a-bc") {
    Write-Host "FAIL: Zero-width lookahead insert replace failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
