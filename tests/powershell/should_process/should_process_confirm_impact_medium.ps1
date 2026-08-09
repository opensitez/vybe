# vybe-test: powershell/should_process/should_process_confirm_impact_medium
function Get-MedImpact {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param()
    return "MedExecuted"
}
$res = Get-MedImpact
if ($res -ne "MedExecuted") {
    Write-Host "FAIL: ConfirmImpact Medium function execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
