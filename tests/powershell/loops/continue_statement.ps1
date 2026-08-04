# vybe-test: powershell/loops/continue_statement
$sum = 0
for ($i = 0; $i -lt 5; $i++) {
    if ($i -eq 2) {
        continue
    }
    $sum += $i
}
if ($sum -ne 8) {
    Write-Host "FAIL: expected 8, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
