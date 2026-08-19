# vybe-test: powershell/interpolation_edge_rules/binding
function test-binding {
    param($x)
    return "$x"
}
if ((test-binding -x 99) -ne '99') {
    Write-Host 'FAIL: bound parameter interpolation expected 99'
    exit 1
}
Write-Host 'PASS'
exit 0
