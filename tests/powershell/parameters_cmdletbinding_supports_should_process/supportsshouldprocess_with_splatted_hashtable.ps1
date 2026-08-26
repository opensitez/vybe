# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/supportsshouldprocess_with_splatted_hashtable
function Lock-Account {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$User)
    if ($PSCmdlet.ShouldProcess($User, "Lock")) {
        return "Locked:$User"
    }
    return "Skipped"
}
$p = @{ User = "eve" }
$res = Lock-Account @p
if ($res -ne "Locked:eve") {
    Write-Host "FAIL: SupportsShouldProcess splatting failed"
    exit 1
}
Write-Host "PASS"
exit 0
