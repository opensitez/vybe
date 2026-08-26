# vybe-test: powershell/parameters_validate_count/validatecount_empty_array_rejected_if_min_count_greater_than_zero
function Set-NonEmptyArr {
    param([ValidateCount(1, 5)][string[]]$Items)
    return $Items.Length
}
$caught = $false
try {
    $x = Set-NonEmptyArr -Items @()
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Empty array should be rejected when min count > 0"
    exit 1
}
Write-Host "PASS"
exit 0
