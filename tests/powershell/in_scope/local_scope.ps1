# vybe-test: powershell/in_scope/local_scope
function Test-Func {
    $x = 1
}
Test-Func
if ($x -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
