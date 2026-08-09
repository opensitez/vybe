# vybe-test: powershell/should_process/should_process_basic_true
function Remove-Data {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Key)
    if ($PSCmdlet.ShouldProcess($Key)) {
        return "Removed"
    }
}
$res = Remove-Data "Key1"
if ($res -ne "Removed") {
    Write-Host "FAIL: ShouldProcess single target overload expected 'Removed', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
