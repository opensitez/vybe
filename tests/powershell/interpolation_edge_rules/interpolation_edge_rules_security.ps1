# vybe-test: powershell/interpolation_edge_rules/security
$unsafe = '$HOME'
$result = "`$unsafe"
if ($result -ne '`$unsafe') {
    Write-Host "FAIL: escaped interpolation should prevent variable expansion"
    exit 1
}
Write-Host 'PASS'
exit 0
