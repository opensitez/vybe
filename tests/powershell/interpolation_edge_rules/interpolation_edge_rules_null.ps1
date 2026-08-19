# vybe-test: powershell/interpolation_edge_rules/null
$word = $null
if ("$word" -ne '') {
    Write-Host "FAIL: null variable should interpolate to empty"
    exit 1
}
Write-Host 'PASS'
exit 0
