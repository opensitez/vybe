# vybe-test: powershell/pipeline_chain_operators_and_or/and_operator_both_succeed
$step1 = $true
$step2 = $false
if ($step1) {
    $step2 = $true
}
if (-not $step2) {
    Write-Host "FAIL: Pipeline chain emulation failed"
    exit 1
}
Write-Host "PASS"
exit 0
