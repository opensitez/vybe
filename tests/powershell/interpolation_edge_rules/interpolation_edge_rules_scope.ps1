# vybe-test: powershell/interpolation_edge_rules/scope
$global:scopeRoot = 'root'
function interp-scope-check {
    $scopeRoot = 'local'
    return "$scopeRoot"
}
if ((interp-scope-check) -ne 'local') {
    Write-Host "FAIL: function-scoped interpolation expected local value"
    exit 1
}
if ($global:scopeRoot -ne 'root') {
    Write-Host "FAIL: global scope should remain root, got '$global:scopeRoot'"
    exit 1
}
Write-Host 'PASS'
exit 0
