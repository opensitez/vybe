# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/supportsshouldprocess_nested_function_call
function Inner-Action {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$T)
    if ($PSCmdlet.ShouldProcess($T, "Inner")) { return "InnerRan" }
    return "InnerSkipped"
}
function Outer-Action {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param([string]$T)
    if ($PSCmdlet.ShouldProcess($T, "Outer")) {
        return Inner-Action -T $T
    }
    return "OuterSkipped"
}
$res = Outer-Action -T "TargetX"
if ($res -ne "InnerRan") {
    Write-Host "FAIL: Nested SupportsShouldProcess call failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
