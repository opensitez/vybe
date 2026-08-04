# vybe-test: powershell/variables/assign_boolean
$flag = $true
if ($flag -ne $true) {
    Write-Host "FAIL: expected True, got $flag"
    exit 1
}
Write-Host "PASS"
exit 0
