# vybe-test: powershell/should_process/should_process_confirm_impact_low
function Get-LowImpact {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
    param()
    return "LowExecuted"
}
$res = Get-LowImpact
if ($res -ne "LowExecuted") {
    Write-Host "FAIL: ConfirmImpact Low function execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
