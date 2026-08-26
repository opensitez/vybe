# vybe-test: powershell/parameters_validate_range/validaterange_above_maximum_throws
function Set-Age2 {
    param([ValidateRange(18, 120)][int]$Age)
    return $Age
}
$caught = $false
try {
    $x = Set-Age2 -Age 121
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected error when value above ValidateRange max"
    exit 1
}
Write-Host "PASS"
exit 0
