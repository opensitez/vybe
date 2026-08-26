# vybe-test: powershell/parameters_validate_range/validaterange_equal_min_and_max
function Enforce-ExactTen {
    param([ValidateRange(10, 10)][int]$Val)
    return $Val
}
$res = Enforce-ExactTen -Val 10
if ($res -ne 10) {
    Write-Host "FAIL: ValidateRange equal min and max failed"
    exit 1
}
Write-Host "PASS"
exit 0
