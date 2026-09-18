# vybe-test: powershell/cmdlets/start_sleep_milliseconds_delay
# Start-Sleep with -Milliseconds pauses execution for at least the specified millisecond duration
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Start-Sleep -Milliseconds 40

$stopwatch.Stop()
$elapsedMs = $stopwatch.ElapsedMilliseconds

# Allow reasonable scheduling tolerance (at least 25ms elapsed)
if ($elapsedMs -lt 25) {
    Write-Host "FAIL: sleep finished too quickly: $elapsedMs ms"
    exit 1
}

Write-Host "PASS"
exit 0
