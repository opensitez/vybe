# vybe-test: powershell/loops/foreach_array
$sum = 0
$numbers = @(1, 2, 3, 4, 5)
foreach ($num in $numbers) {
    $sum += $num
}
if ($sum -ne 15) {
    Write-Host "FAIL: expected 15, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
