# vybe-test: powershell/control_flow/if_else_branching
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
