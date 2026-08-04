# vybe-test: powershell/modules/module_export_function
function Get-ModuleValue {
    return 42
}
Export-ModuleMember -Function Get-ModuleValue
$result = Get-ModuleValue
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
