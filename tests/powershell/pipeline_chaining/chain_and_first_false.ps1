# vybe-test: powershell/pipeline_chaining/chain_and_first_false
$log = @()
($script:log += "First"; $false) && ($script:log += "Second"; $true)
if ($log.Count -ne 1 -or $log[0] -ne "First") {
    Write-Host "FAIL: && short circuit failed, second executed"
    exit 1
}
Write-Host "PASS"
exit 0
