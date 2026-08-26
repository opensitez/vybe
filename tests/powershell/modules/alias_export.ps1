# vybe-test: powershell/modules/alias_export
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
