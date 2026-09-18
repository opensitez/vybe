# vybe-test: powershell/modules/alias_export
function Show-Value {
    return 'ok'
}
Export-ModuleMember -Function Show-Value -Alias show
$result = show
if ($result -ne 'ok') {
    Write-Host "FAIL: expected ok, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
