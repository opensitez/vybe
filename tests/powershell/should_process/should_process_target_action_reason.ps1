# vybe-test: powershell/should_process/should_process_target_action_reason
function Rebuild-Index {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Index)
    $reason = [System.Management.Automation.ShouldProcessReason]::None
    if ($PSCmdlet.ShouldProcess($Index, "Rebuild", [ref]$reason)) {
        return "Rebuilt"
    }
}
$res = Rebuild-Index "MainIndex"
if ($res -ne "Rebuilt") {
    Write-Host "FAIL: ShouldProcess with reason ref parameter expected Rebuilt"
    exit 1
}
Write-Host "PASS"
exit 0
