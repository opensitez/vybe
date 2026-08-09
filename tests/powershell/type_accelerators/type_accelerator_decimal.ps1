# vybe-test: powershell/type_accelerators/type_accelerator_decimal
$m = [decimal]99.95m
if ($m -ne 99.95m) {
    Write-Host "FAIL: decimal expected 99.95m, got $m"
    exit 1
}
Write-Host "PASS"
exit 0
