# vybe-test: powershell/parameters_validate_count/validatecount_below_min_throws
function Select-Items {
    param([ValidateCount(2, 5)][string[]]$Items)
    return $Items.Length
}
$caught = $false
try {
    $x = Select-Items -Items "single"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when array count below ValidateCount min"
    exit 1
}
Write-Host "PASS"
exit 0
