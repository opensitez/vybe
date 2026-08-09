# vybe-test: powershell/pipeline_chaining/chain_side_effects
$count = 0
($false) && ($script:count++)
if ($count -ne 0) {
    Write-Host "FAIL: side effect executed on short-circuited &&"
    exit 1
}
Write-Host "PASS"
exit 0
