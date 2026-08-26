# vybe-test: powershell/pipeline_chain_operators_and_or/mixed_and_or_precedence_left_associative
$pass1 = $true
$pass2 = $true
$combined = $pass1 -and $pass2
if (-not $combined) {
    Write-Host "FAIL: Chain logic check failed"
    exit 1
}
Write-Host "PASS"
exit 0
