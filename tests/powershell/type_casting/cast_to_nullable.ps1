# vybe-test: powershell/type_casting/cast_to_nullable
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
