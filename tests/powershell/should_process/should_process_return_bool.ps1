# vybe-test: powershell/should_process/should_process_return_bool
function Test-BoolReturn {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    return $PSCmdlet.ShouldProcess("Target")
}
$res = Test-BoolReturn
if ($res -ne $true) {
    Write-Host "FAIL: ShouldProcess boolean return expected true, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
