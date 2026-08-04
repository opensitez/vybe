# vybe-test: powershell/in_scope/variable_shadowing
$x = 1
function Test-Func {
    $x = 2
    Write-Output $x
}
if ((Test-Func) -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
