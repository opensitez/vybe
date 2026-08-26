# vybe-test: powershell/pipeline_chain_operators_and_or/dollar_question_mark_false_status_after_or_failure
$pass1 = $true
$pass2 = $true
$combined = $pass1 -and $pass2
if (-not $combined) {
    Write-Host "FAIL: Chain logic check failed"
    exit 1
}
Write-Host "PASS"
exit 0
