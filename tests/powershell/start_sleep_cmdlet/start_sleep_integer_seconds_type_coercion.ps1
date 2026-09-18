# vybe-test: powershell/start_sleep_cmdlet/start_sleep_integer_seconds_type_coercion
# Integer instances passed to -Seconds successfully coerce to double
$secInt = [int]0
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Seconds $secInt
$sw.Stop()

if ($sw.ElapsedMilliseconds -gt 250) {
    Write-Host "FAIL: integer seconds took unexpectedly long: $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
