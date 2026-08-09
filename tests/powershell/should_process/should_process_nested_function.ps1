# vybe-test: powershell/should_process/should_process_nested_function
function Outer-Cmd {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    function Inner-Cmd {
        [CmdletBinding(SupportsShouldProcess=$true)]
        param()
        if ($PSCmdlet.ShouldProcess("InnerTarget", "InnerAction")) {
            return "InnerPass"
        }
    }
    return Inner-Cmd
}
$res = Outer-Cmd
if ($res -ne "InnerPass") {
    Write-Host "FAIL: nested function ShouldProcess expected InnerPass, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
