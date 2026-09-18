# vybe-test: powershell/start_sleep_cmdlet/start_sleep_fractional_seconds_precision
# Start-Sleep supports high precision sub-second delays (e.g. 0.025s = 25ms)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Seconds 0.025
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 15) {
    Write-Host "FAIL: fractional sub-second sleep insufficient, got $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
