# vybe-test: powershell/pipeline_chaining/chain_and_both_true
$log = @()
($script:log += "First"; $true) && ($script:log += "Second"; $true)
if ($log.Count -ne 2 -or $log[0] -ne "First" -or $log[1] -ne "Second") {
    Write-Host "FAIL: && with both true expected First, Second logged"
    exit 1
}
Write-Host "PASS"
exit 0
