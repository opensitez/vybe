# vybe-test: powershell/parameters_validate_count/validatecount_with_integer_array
function Sum-Numbers {
    param([ValidateCount(2, 4)][int[]]$Nums)
    $sum = 0
    foreach ($n in $Nums) { $sum += $n }
    return $sum
}
$res = Sum-Numbers -Nums 10, 20, 30
if ($res -ne 60) {
    Write-Host "FAIL: ValidateCount int array sum failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
