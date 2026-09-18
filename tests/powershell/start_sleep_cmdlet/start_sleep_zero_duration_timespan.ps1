# vybe-test: powershell/start_sleep_cmdlet/start_sleep_zero_duration_timespan
# Passing [TimeSpan]::Zero to -Duration returns immediately without error
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Duration ([TimeSpan]::Zero)
$sw.Stop()

if ($sw.ElapsedMilliseconds -gt 250) {
    Write-Host "FAIL: Start-Sleep -Duration Zero took unexpectedly long: $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
