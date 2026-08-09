# vybe-test: powershell/readonly_variables/readonly_variable_curly_braces
New-Variable -Name "RO WITH SPACE" -Value "SpacedData" -Option ReadOnly
if (${RO WITH SPACE} -ne "SpacedData") {
    Write-Host "FAIL: curly brace ReadOnly variable expected SpacedData, got ${RO WITH SPACE}"
    exit 1
}
Write-Host "PASS"
exit 0
