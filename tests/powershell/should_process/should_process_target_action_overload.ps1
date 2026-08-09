# vybe-test: powershell/should_process/should_process_target_action_overload
function Update-Cache {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Key)
    if ($PSCmdlet.ShouldProcess($Key, "UpdateCache")) {
        return "Updated"
    }
}
$res = Update-Cache "SessionKey"
if ($res -ne "Updated") {
    Write-Host "FAIL: ShouldProcess target/action overload expected Updated, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
