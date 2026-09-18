# vybe-test: powershell/start_sleep_cmdlet/start_sleep_duration_timespan_parameter
# The -Duration parameter accepts System.TimeSpan instances
$duration = [TimeSpan]::FromMilliseconds(50)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Duration $duration
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 30) {
    Write-Host "FAIL: -Duration sleep insufficient, expected >= 30ms, got $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
