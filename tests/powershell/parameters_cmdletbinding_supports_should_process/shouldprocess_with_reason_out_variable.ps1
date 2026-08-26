# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/shouldprocess_with_reason_out_variable
function Check-ReasonCall {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Item)
    [System.Management.Automation.ShouldProcessReason]$reason = [System.Management.Automation.ShouldProcessReason]::None
    $ok = $PSCmdlet.ShouldProcess("Description", "Warning", "Caption", [ref]$reason)
    return $ok
}
$res = Check-ReasonCall -Item "Test"
if ($res -ne $true) {
    Write-Host "FAIL: ShouldProcess with reason ref parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
