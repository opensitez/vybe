# vybe-test: powershell/function_scope/nested_function_scope
function Outer { function Inner { return 'PASS' }; return (Inner) }
if ((Outer) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
