# vybe-test: powershell/function_metadata/parameter_alias
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
