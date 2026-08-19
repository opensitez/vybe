# vybe-test: powershell/interpolation_edge_rules/invalid
$result = "$missingVar"
if ($result -ne '') {
    Write-Host "FAIL: unknown variable should interpolate to empty string"
    exit 1
}
Write-Host 'PASS'
exit 0
