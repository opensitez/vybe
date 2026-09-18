# vybe-test: powershell/start_sleep_cmdlet/start_sleep_positional_seconds_parameter
# Start-Sleep binds a positional numeric argument to the -Seconds parameter
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep 0.05
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 30) {
    Write-Host "FAIL: positional sleep duration insufficient, expected >= 30ms, got $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
