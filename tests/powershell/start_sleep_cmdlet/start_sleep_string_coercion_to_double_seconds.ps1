# vybe-test: powershell/start_sleep_cmdlet/start_sleep_string_coercion_to_double_seconds
# Passing a string representation of a float to -Seconds coerces into System.Double
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Seconds "0.04"
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 25) {
    Write-Host "FAIL: string coercion to double sleep duration insufficient: $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
