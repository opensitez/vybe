# vybe-test: powershell/loops/for_loop_sum
$sum = 0
for ($i = 1; $i -le 10; $i++) {
    $sum += $i
}
if ($sum -ne 55) {
    Write-Host "FAIL: expected 55, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
