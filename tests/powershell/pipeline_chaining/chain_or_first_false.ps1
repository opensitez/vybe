# vybe-test: powershell/pipeline_chaining/chain_or_first_false
$log = @()
($script:log += "First"; $false) || ($script:log += "Second"; $true)
if ($log.Count -ne 2 -or $log[0] -ne "First" -or $log[1] -ne "Second") {
    Write-Host "FAIL: || with first false expected both First and Second logged"
    exit 1
}
Write-Host "PASS"
exit 0
