# vybe-test: powershell/variables/multiple_assignment_swap
Set-Variable -Name "testVarCheck" -Value 42 -Option ReadOnly -Force
$val = (Get-Variable -Name "testVarCheck").Value
Remove-Variable -Name "testVarCheck" -Force
if ($val -ne 42) {
    Write-Host "FAIL: Variable check failed"
    exit 1
}
Write-Host "PASS"
exit 0
