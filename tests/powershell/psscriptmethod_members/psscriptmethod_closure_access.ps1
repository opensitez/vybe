# vybe-test: powershell/psscriptmethod_members/psscriptmethod_closure_access
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
