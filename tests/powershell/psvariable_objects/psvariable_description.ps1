# vybe-test: powershell/psvariable_objects/psvariable_description
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
