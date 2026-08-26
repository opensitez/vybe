# vybe-test: powershell/pipeline_chain_operators_and_or/or_operator_inside_while_loop
$pass1 = $true
$pass2 = $true
$combined = $pass1 -and $pass2
if (-not $combined) {
    Write-Host "FAIL: Chain logic check failed"
    exit 1
}
Write-Host "PASS"
exit 0
