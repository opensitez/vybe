# vybe-test: powershell/pipeline_chaining/chain_subexpression
$msg = "Result: $( ($true) && 'OK' )"
if ($msg -ne "Result: OK") {
    Write-Host "FAIL: chain in subexpression expected 'Result: OK', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
