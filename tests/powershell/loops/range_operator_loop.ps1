# vybe-test: powershell/loops/range_operator_loop
$sum = 0
foreach ($n in 1..100) { $sum += $n }
# Gauss: 100*101/2 = 5050
if ($sum -ne 5050) {
    Write-Host "FAIL: expected 5050, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
