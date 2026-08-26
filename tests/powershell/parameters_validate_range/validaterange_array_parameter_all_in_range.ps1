# vybe-test: powershell/parameters_validate_range/validaterange_array_parameter_all_in_range
function Process-Scores {
    param([ValidateRange(1, 10)][int[]]$Scores)
    return ($Scores | Measure-Object -Sum).Sum
}
$sum = Process-Scores -Scores 2, 4, 6, 8
if ($sum -ne 20) {
    Write-Host "FAIL: ValidateRange array sum failed, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
