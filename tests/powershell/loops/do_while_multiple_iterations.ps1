# vybe-test: powershell/loops/do_while_multiple_iterations
$i = 0
$sum = 0
do {
    $sum += $i
    $i++
} while ($i -lt 5)
# 0+1+2+3+4 = 10
if ($sum -ne 10) {
    Write-Host "FAIL: expected 10, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
