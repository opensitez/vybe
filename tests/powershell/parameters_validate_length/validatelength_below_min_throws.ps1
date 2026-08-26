# vybe-test: powershell/parameters_validate_length/validatelength_below_min_throws
function Set-Name {
    param([ValidateLength(3, 10)][string]$Name)
    return $Name
}
$caught = $false
try {
    $x = Set-Name -Name "Al" # length 2 < 3
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when string shorter than min length"
    exit 1
}
Write-Host "PASS"
exit 0
