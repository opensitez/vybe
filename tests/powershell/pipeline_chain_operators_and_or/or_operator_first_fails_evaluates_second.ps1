# vybe-test: powershell/pipeline_chain_operators_and_or/or_operator_first_fails_evaluates_second
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
