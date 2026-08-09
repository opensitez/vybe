# vybe-test: powershell/pipeline_chaining/chain_mixed_and_or
$log = @()
($script:log += "A"; $false) && ($script:log += "B"; $true) || ($script:log += "C"; $true)
if ($log.Count -ne 2 -or $log[0] -ne "A" -or $log[1] -ne "C") {
    Write-Host "FAIL: A && B || C expected A, C logged, got $($log -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
