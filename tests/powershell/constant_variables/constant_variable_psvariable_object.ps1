# vybe-test: powershell/constant_variables/constant_variable_psvariable_object
$vObj = [System.Management.Automation.PSVariable]::new("PSV_CONST", "Val", [System.Management.Automation.ScopedItemOptions]::Constant)
Set-Variable -Option Constant -Name "PSV_CONST" -Value "Val"
if ($vObj.Options.ToString() -ne "Constant") {
    Write-Host "FAIL: PSVariable object Options expected Constant, got $($vObj.Options)"
    exit 1
}
Write-Host "PASS"
exit 0
