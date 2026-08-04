# vybe-test: powershell/modules/module_scoped_variable
$script:Value = 10
function Get-Value {
    return $script:Value
}
Export-ModuleMember -Function Get-Value
$result = Get-Value
if ($result -ne 10) {
    Write-Host "FAIL: expected 10, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
