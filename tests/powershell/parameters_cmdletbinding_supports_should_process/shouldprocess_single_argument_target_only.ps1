# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/shouldprocess_single_argument_target_only
function Reset-ItemState {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Target)
    if ($PSCmdlet.ShouldProcess($Target)) {
        return "ResetDone"
    }
    return "Skipped"
}
$res = Reset-ItemState -Target "Database"
if ($res -ne "ResetDone") {
    Write-Host "FAIL: ShouldProcess target only call failed"
    exit 1
}
Write-Host "PASS"
exit 0
