# vybe-test: powershell/statement_terminators/function_definition
function Test-Func {
    Write-Output 'PASS'
}
if ((Test-Func) -eq 'PASS') { exit 0 }
exit 1
