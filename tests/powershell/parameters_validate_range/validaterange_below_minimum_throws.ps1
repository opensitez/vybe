# vybe-test: powershell/parameters_validate_range/validaterange_below_minimum_throws
function Set-Age {
    param([ValidateRange(18, 120)][int]$Age)
    return $Age
}
$caught = $false
try {
    $x = Set-Age -Age 17
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected error when value below ValidateRange min"
    exit 1
}
Write-Host "PASS"
exit 0
