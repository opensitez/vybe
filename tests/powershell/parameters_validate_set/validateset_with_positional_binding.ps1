# vybe-test: powershell/parameters_validate_set/validateset_with_positional_binding
function Invoke-Action {
    param(
        [Parameter(Position=0)]
        [ValidateSet("Start", "Stop", "Restart")]
        [string]$Action
    )
    return "Action:$Action"
}
$res = Invoke-Action "Restart"
if ($res -ne "Action:Restart") {
    Write-Host "FAIL: Positional ValidateSet argument failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
