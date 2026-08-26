# vybe-test: powershell/parameters_validate_count/validatecount_zero_min_count_allows_empty_array
function Set-OptionalArr {
    param([ValidateCount(0, 3)][string[]]$Items)
    return $Items.Length
}
$res = Set-OptionalArr -Items @()
if ($res -ne 0) {
    Write-Host "FAIL: ValidateCount(0, 3) on empty array failed"
    exit 1
}
Write-Host "PASS"
exit 0
