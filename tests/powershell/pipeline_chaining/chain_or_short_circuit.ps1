# vybe-test: powershell/pipeline_chaining/chain_or_short_circuit
$executed = $false
"Truthy" || ($script:executed = $true)
if ($executed) {
    Write-Host "FAIL: Truthy || RHS executed"
    exit 1
}
Write-Host "PASS"
exit 0
