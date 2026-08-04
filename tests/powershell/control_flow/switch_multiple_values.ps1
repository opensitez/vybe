# vybe-test: powershell/control_flow/switch_multiple_values
$values = @(1, 2, 3)
$sum = 0
switch ($values) {
    1 { $sum += 10 }
    2 { $sum += 20 }
    3 { $sum += 30 }
}
if ($sum -ne 60) {
    Write-Host "FAIL: expected 60, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
