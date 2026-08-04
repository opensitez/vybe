# vybe-test: powershell/operators/subexpression_operator
$arr = @(1,2,3,4,5)
$sum = $(foreach ($n in $arr) { $n }) | Measure-Object -Sum
if ($sum.Sum -ne 15) { Write-Host "FAIL: sum $($sum.Sum)"; exit 1 }
# $() forces single-value context on a multi-item expression
$msg = "Count: $($arr.Count)"
if ($msg -ne "Count: 5") { Write-Host "FAIL: '$msg'"; exit 1 }
Write-Host "PASS"
exit 0
