# vybe-test: powershell/readonly_variables/readonly_variable_psvariable_object
$vObj = [System.Management.Automation.PSVariable]::new("RO_OBJ", "ObjVal", [System.Management.Automation.ScopedItemOptions]::ReadOnly)
Set-Variable -Option ReadOnly -Name "RO_OBJ" -Value "ObjVal"
if ($vObj.Options.ToString() -ne "ReadOnly") {
    Write-Host "FAIL: PSVariable object Options expected ReadOnly, got $($vObj.Options)"
    exit 1
}
Write-Host "PASS"
exit 0
