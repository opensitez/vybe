# vybe-test: powershell/psvariable_objects/psvariable_remove_variable_cmdlet
$ToRemove = "Temporary"
Remove-Variable -Name "ToRemove"
if (Get-Variable -Name "ToRemove" -ErrorAction SilentlyContinue) {
    Write-Host "FAIL: Remove-Variable failed, variable still exists"
    exit 1
}
Write-Host "PASS"
exit 0
