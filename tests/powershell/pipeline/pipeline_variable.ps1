# vybe-test: powershell/pipeline/pipeline_variable
$result = 1..3 | ForEach-Object { $_ + 10 }
$sum = 0
foreach ($val in $result) {
    $sum += $val
}
if ($sum -ne 36) {
    Write-Host "FAIL: expected 36, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
