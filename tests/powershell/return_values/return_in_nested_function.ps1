# vybe-test: powershell/return_values/return_in_nested_function
function Outer { function Inner { return 'PASS' }; return (Inner) }
if ((Outer) -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
