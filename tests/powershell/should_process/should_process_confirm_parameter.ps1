# vybe-test: powershell/should_process/should_process_confirm_parameter
function Stop-Daemon {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    param([string]$DaemonName)
    if ($PSCmdlet.ShouldProcess($DaemonName, "Stop")) {
        return "Stopped"
    }
    return "NotStopped"
}
$cmd = Get-Command Stop-Daemon
if ($cmd.Parameters.ContainsKey("Confirm") -and $cmd.Parameters.ContainsKey("WhatIf")) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL: SupportsShouldProcess cmdlet missing -Confirm / -WhatIf parameters"
exit 1
