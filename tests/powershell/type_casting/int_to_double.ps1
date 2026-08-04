# vybe-test: powershell/type_casting/int_to_double
$value = [double]5
if ($value -ne 5.0) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
