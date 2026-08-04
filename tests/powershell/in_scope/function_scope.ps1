# vybe-test: powershell/in_scope/function_scope
function Test-Func {
    $x = 2
    return $x
}
if ((Test-Func) -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
