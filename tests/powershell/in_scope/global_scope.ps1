# vybe-test: powershell/in_scope/global_scope
function Test-Func {
    $global:x = 1
}
Test-Func
if ($x -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
