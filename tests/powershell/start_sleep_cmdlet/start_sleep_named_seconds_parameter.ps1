# vybe-test: powershell/start_sleep_cmdlet/start_sleep_named_seconds_parameter
# The -Seconds parameter accepts fractional floating point values (System.Double)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Seconds 0.05
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 30) {
    Write-Host "FAIL: named -Seconds sleep insufficient, got $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
