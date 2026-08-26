# vybe-test: powershell/data_conversion/to_int_to_string
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
