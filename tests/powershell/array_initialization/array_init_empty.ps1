# vybe-test: powershell/array_initialization/array_init_empty
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
