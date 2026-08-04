# vybe-test: powershell/ternary/ternary_with_strings
$result = ('a' -eq 'b') ? 'match' : 'nomatch'
if ($result -ne 'nomatch') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
