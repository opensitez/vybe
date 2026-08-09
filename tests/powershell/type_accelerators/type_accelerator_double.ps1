# vybe-test: powershell/type_accelerators/type_accelerator_double
$d = [double]"3.14159"
if ($d -ne 3.14159) {
    Write-Host "FAIL: double expected 3.14159, got $d"
    exit 1
}
Write-Host "PASS"
exit 0
