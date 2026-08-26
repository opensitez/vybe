# vybe-test: powershell/parameters_validate_length/validatelength_above_max_throws
function Set-Name2 {
    param([ValidateLength(3, 5)][string]$Name)
    return $Name
}
$caught = $false
try {
    $x = Set-Name2 -Name "Alexander" # length 9 > 5
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when string longer than max length"
    exit 1
}
Write-Host "PASS"
exit 0
