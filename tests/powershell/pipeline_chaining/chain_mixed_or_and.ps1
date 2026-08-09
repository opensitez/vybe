# vybe-test: powershell/pipeline_chaining/chain_mixed_or_and
$log = @()
($script:log += "1"; $false) || ($script:log += "2"; $true) && ($script:log += "3"; $true)
if ($log.Count -ne 3 -or $log[0] -ne "1" -or $log[1] -ne "2" -or $log[2] -ne "3") {
    Write-Host "FAIL: 1 || 2 && 3 expected 1, 2, 3 logged, got $($log -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
