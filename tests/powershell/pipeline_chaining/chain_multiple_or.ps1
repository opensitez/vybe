# vybe-test: powershell/pipeline_chaining/chain_multiple_or
$step = 0
($script:step = 1; $false) || ($script:step = 2; $false) || ($script:step = 3; $true)
if ($step -ne 3) {
    Write-Host "FAIL: multiple || operators step expected 3, got $step"
    exit 1
}
Write-Host "PASS"
exit 0
