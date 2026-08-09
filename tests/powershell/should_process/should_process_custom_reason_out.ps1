# vybe-test: powershell/should_process/should_process_custom_reason_out
function Test-ReasonOut {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Name)
    $reason = [System.Management.Automation.ShouldProcessReason]::None
    $shouldRun = $PSCmdlet.ShouldProcess($Name, "Action", [ref]$reason)
    return [pscustomobject]@{ Allowed = $shouldRun; Reason = $reason }
}
$res = Test-ReasonOut "Resource1"
if (-not $res.Allowed) {
    Write-Host "FAIL: Reason out expected Allowed=true"
    exit 1
}
Write-Host "PASS"
exit 0
