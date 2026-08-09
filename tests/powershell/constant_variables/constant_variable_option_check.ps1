# vybe-test: powershell/constant_variables/constant_variable_option_check
New-Variable -Name "OPT_CONST" -Value 1 -Option Constant
$varObj = Get-Variable -Name "OPT_CONST"
if (-not ($varObj.Options -band [System.Management.Automation.ScopedItemOptions]::Constant)) {
    Write-Host "FAIL: ScopedItemOptions Constant flag missing"
    exit 1
}
Write-Host "PASS"
exit 0
