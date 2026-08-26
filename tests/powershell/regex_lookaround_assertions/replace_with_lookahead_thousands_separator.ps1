# vybe-test: powershell/regex_lookaround_assertions/replace_with_lookahead_thousands_separator
$num = "1234567890"
$formatted = $num -replace "(?<=\d)(?=(\d{3})+$)", ","
if ($formatted -ne "1,234,567,890") {
    Write-Host "FAIL: Thousands separator lookahead replace failed, got '$formatted'"
    exit 1
}
Write-Host "PASS"
exit 0
