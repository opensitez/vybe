# vybe-test: powershell/parameters_validate_set/validateset_with_splatted_arguments
function Set-Theme {
    param([ValidateSet("Light", "Dark", "HighContrast")][string]$Theme)
    return "Theme:$Theme"
}
$params = @{ Theme = "Dark" }
$res = Set-Theme @params
if ($res -ne "Theme:Dark") {
    Write-Host "FAIL: Splatted ValidateSet parameter failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
