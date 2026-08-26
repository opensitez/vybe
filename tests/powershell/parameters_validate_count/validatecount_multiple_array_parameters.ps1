# vybe-test: powershell/parameters_validate_count/validatecount_multiple_array_parameters
function Pair-Arrays {
    param(
        [ValidateCount(1, 2)][string[]]$A,
        [ValidateCount(1, 2)][int[]]$B
    )
    return "$($A.Length)-$($B.Length)"
}
$res = Pair-Arrays -A "x" -B 10, 20
if ($res -ne "1-2") {
    Write-Host "FAIL: Multiple ValidateCount array parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
