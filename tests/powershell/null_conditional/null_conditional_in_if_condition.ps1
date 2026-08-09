# vybe-test: powershell/null_conditional/null_conditional_in_if_condition
$user = $null
if (${user}?.IsAdmin) {
    Write-Host "FAIL: null conditional evaluated as truthy"
    exit 1
}
Write-Host "PASS"
exit 0
