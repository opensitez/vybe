# vybe-test: powershell/should_process/should_process_confirm_impact_high
function Get-HighImpact {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    param()
    return "HighExecuted"
}
$res = Get-HighImpact
if ($res -ne "HighExecuted") {
    Write-Host "FAIL: ConfirmImpact High function execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
