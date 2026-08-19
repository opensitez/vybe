# vybe-test: powershell/interpolation_edge_rules/empty
$word = ''
if ("$word" -ne '') {
    Write-Host "FAIL: empty variable should interpolate to empty"
    exit 1
}
Write-Host 'PASS'
exit 0
