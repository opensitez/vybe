# vybe-test: powershell/should_process/should_process_supports_should_process
function Invoke-ItemAction {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$Target)
    process {
        if ($PSCmdlet.ShouldProcess($Target, "Action")) {
            return "Executed"
        }
        return "Skipped"
    }
}
$res = Invoke-ItemAction -Target "TestResource"
if ($res -ne "Executed") {
    Write-Host "FAIL: ShouldProcess basic invocation expected 'Executed', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
