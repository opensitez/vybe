# vybe-test: powershell/pipeline_chaining/chain_and_short_circuit
$executed = $false
$null && ($script:executed = $true)
if ($executed) {
    Write-Host "FAIL: null && RHS executed"
    exit 1
}
Write-Host "PASS"
exit 0
