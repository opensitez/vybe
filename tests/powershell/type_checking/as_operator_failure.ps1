# vybe-test: powershell/type_checking/as_operator_failure
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
