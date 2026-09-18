# vybe-test: powershell/start_sleep_cmdlet/start_sleep_zero_seconds_immediate_return
# Calling Start-Sleep with 0 seconds returns immediately without error
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Seconds 0
$sw.Stop()

if ($sw.ElapsedMilliseconds -gt 250) {
    Write-Host "FAIL: Start-Sleep 0 took unexpectedly long: $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
