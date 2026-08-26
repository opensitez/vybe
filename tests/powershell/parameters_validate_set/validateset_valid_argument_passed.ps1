# vybe-test: powershell/parameters_validate_set/validateset_valid_argument_passed
function Set-Env {
    param([ValidateSet("Dev", "Test", "Prod")][string]$EnvName)
    return "Target:$EnvName"
}
$res = Set-Env -EnvName "Test"
if ($res -ne "Target:Test") {
    Write-Host "FAIL: ValidateSet valid argument failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
