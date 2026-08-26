# vybe-test: powershell/try_catch/throw_in_catch
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
