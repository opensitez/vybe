# vybe-test: powershell/pipeline_chaining/chain_multiple_and
$step = 0
($script:step = 1; $true) && ($script:step = 2; $true) && ($script:step = 3; $true)
if ($step -ne 3) {
    Write-Host "FAIL: multiple && operators step expected 3, got $step"
    exit 1
}
Write-Host "PASS"
exit 0
