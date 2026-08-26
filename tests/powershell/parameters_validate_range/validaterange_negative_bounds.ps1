# vybe-test: powershell/parameters_validate_range/validaterange_negative_bounds
function Set-Temp {
    param([ValidateRange(-50, 50)][int]$Temp)
    return $Temp
}
$res = Set-Temp -Temp -25
if ($res -ne -25) {
    Write-Host "FAIL: ValidateRange negative bounds failed"
    exit 1
}
Write-Host "PASS"
exit 0
