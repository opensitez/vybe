# vybe-test: powershell/operators/pipeline_chain_operators
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
