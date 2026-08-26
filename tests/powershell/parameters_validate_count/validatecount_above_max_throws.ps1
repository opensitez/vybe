# vybe-test: powershell/parameters_validate_count/validatecount_above_max_throws
function Select-Items2 {
    param([ValidateCount(1, 2)][string[]]$Items)
    return $Items.Length
}
$caught = $false
try {
    $x = Select-Items2 -Items "a", "b", "c" # count 3 > 2
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when array count above ValidateCount max"
    exit 1
}
Write-Host "PASS"
exit 0
