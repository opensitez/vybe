# vybe-test: powershell/parameters_validate_set/validateset_with_default_parameter_value
function Get-Mode {
    param([ValidateSet("Fast", "Slow")][string]$Mode = "Fast")
    return $Mode
}
$res = Get-Mode
if ($res -ne "Fast") {
    Write-Host "FAIL: ValidateSet with default value failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
