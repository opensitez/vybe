# vybe-test: powershell/should_process/should_process_subexpression
function Get-Status {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    if ($PSCmdlet.ShouldProcess("Res")) { return "OK" }
}
$msg = "Result: $( Get-Status )"
if ($msg -ne "Result: OK") {
    Write-Host "FAIL: ShouldProcess in subexpression expected 'Result: OK', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
