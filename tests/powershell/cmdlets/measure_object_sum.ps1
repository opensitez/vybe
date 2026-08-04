# vybe-test: powershell/cmdlets/measure_object_sum
$numbers = 1, 2, 3, 4, 5
$result = $numbers | Measure-Object -Sum
$sum = $result.Sum
if ($sum -ne 15) {
    Write-Host "FAIL: expected 15, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
