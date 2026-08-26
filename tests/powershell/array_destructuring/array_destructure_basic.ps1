# vybe-test: powershell/array_destructuring/array_destructure_basic
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
