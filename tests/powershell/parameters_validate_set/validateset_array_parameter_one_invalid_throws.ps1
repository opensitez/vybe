# vybe-test: powershell/parameters_validate_set/validateset_array_parameter_one_invalid_throws
function Select-Colors2 {
    param([ValidateSet("Red", "Green", "Blue")][string[]]$Colors)
    return $Colors
}
$caught = $false
try {
    $x = Select-Colors2 -Colors "Red", "Yellow"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when one element in array fails ValidateSet"
    exit 1
}
Write-Host "PASS"
exit 0
