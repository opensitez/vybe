# vybe-test: powershell/modules/simple_module_export
function Get-ModuleValue {
    return 7
}
Export-ModuleMember -Function Get-ModuleValue
$result = Get-ModuleValue
if ($result -ne 7) {
    Write-Host "FAIL: expected 7, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
