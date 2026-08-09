# vybe-test: powershell/constant_variables/constant_variable_curly_braces
New-Variable -Name "CONST WITH SPACE" -Value "SpacedConst" -Option Constant
if (${CONST WITH SPACE} -ne "SpacedConst") {
    Write-Host "FAIL: curly brace Constant variable expected SpacedConst, got ${CONST WITH SPACE}"
    exit 1
}
Write-Host "PASS"
exit 0
