# vybe-test: powershell/start_sleep_cmdlet/start_sleep_loop_iteration_accumulation
# Start-Sleep calls across loop iterations accumulate elapsed execution time
$sw = [System.Diagnostics.Stopwatch]::StartNew()
for ($i = 0; $i -lt 3; $i++) {
    Start-Sleep -Milliseconds 20
}
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 45) {
    Write-Host "FAIL: accumulated loop sleep duration insufficient, expected >= 45ms, got $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
