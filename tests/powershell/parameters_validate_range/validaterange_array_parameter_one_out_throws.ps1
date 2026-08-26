# vybe-test: powershell/parameters_validate_range/validaterange_array_parameter_one_out_throws
function Process-Scores2 {
    param([ValidateRange(1, 10)][int[]]$Scores)
    return $Scores
}
$caught = $false
try {
    $x = Process-Scores2 -Scores 5, 15, 8
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected error when array item out of ValidateRange"
    exit 1
}
Write-Host "PASS"
exit 0
