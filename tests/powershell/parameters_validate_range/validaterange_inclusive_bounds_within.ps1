# vybe-test: powershell/parameters_validate_range/validaterange_inclusive_bounds_within
function Set-Percentage {
    param([ValidateRange(0, 100)][int]$Pct)
    return "Pct:$Pct"
}
$r1 = Set-Percentage -Pct 0
$r2 = Set-Percentage -Pct 50
$r3 = Set-Percentage -Pct 100
if ($r1 -ne "Pct:0" -or $r2 -ne "Pct:50" -or $r3 -ne "Pct:100") {
    Write-Host "FAIL: ValidateRange inclusive bounds failed"
    exit 1
}
Write-Host "PASS"
exit 0
