# vybe-test: powershell/parameters_validate_range/validaterange_splatting_argument
function Set-Level {
    param([ValidateRange(1, 5)][int]$Level)
    return "Level:$Level"
}
$params = @{ Level = 3 }
$res = Set-Level @params
if ($res -ne "Level:3") {
    Write-Host "FAIL: ValidateRange splatting failed"
    exit 1
}
Write-Host "PASS"
exit 0
