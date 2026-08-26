# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/pscmdlet_shouldprocess_target_and_action
function Invoke-SafeAction {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Item)
    if ($PSCmdlet.ShouldProcess($Item, "Delete")) {
        return "Deleted:$Item"
    }
    return "Skipped:$Item"
}
$res = Invoke-SafeAction -Item "file.txt"
if ($res -ne "Deleted:file.txt") {
    Write-Host "FAIL: ShouldProcess standard call failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
