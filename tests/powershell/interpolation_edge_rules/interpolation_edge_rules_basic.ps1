# vybe-test: powershell/interpolation_edge_rules/basic
$value = 'ok'
if ("$value" -ne 'ok') {
    Write-Host "FAIL: basic interpolation expected ok"
    exit 1
}
Write-Host 'PASS'
exit 0
