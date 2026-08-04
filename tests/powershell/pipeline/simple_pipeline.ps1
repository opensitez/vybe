# vybe-test: powershell/pipeline/simple_pipeline
$result = 1..3 | ForEach-Object { $_ * 2 }
$sum = 0
foreach ($val in $result) {
    $sum += $val
}
if ($sum -ne 12) {
    Write-Host "FAIL: expected 12, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
