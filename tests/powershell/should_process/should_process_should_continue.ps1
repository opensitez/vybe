# vybe-test: powershell/should_process/should_process_should_continue
function Prompt-Continue {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    # ShouldContinue is available on $PSCmdlet
    if ($PSCmdlet.GetType().GetMethod("ShouldContinue") -ne $null) {
        return "ShouldContinueAvailable"
    }
    return "MissingMethod"
}
$res = Prompt-Continue
if ($res -ne "ShouldContinueAvailable") {
    Write-Host "FAIL: PSCmdlet.ShouldContinue method check failed"
    exit 1
}
Write-Host "PASS"
exit 0
