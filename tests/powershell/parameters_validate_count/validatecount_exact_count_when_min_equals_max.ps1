# vybe-test: powershell/parameters_validate_count/validatecount_exact_count_when_min_equals_max
function Require-Pair {
    param([ValidateCount(2, 2)][string[]]$Pair)
    return "$($Pair[0])-$($Pair[1])"
}
$res = Require-Pair -Pair "left", "right"
if ($res -ne "left-right") {
    Write-Host "FAIL: Exact count ValidateCount(2, 2) failed"
    exit 1
}
Write-Host "PASS"
exit 0
