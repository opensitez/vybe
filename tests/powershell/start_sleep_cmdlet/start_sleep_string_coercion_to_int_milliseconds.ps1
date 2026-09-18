# vybe-test: powershell/start_sleep_cmdlet/start_sleep_string_coercion_to_int_milliseconds
# Passing a string representation of an integer to -Milliseconds coerces into System.Int32
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Milliseconds "40"
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 25) {
    Write-Host "FAIL: string coercion to int sleep duration insufficient: $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
