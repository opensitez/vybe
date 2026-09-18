# vybe-test: powershell/start_sleep_cmdlet/start_sleep_measure_command_integration
# Measure-Command accurately captures execution duration of Start-Sleep
$elapsed = Measure-Command {
    Start-Sleep -Milliseconds 40
}

if ($elapsed.TotalMilliseconds -lt 25) {
    Write-Host "FAIL: measured duration insufficient, expected >= 25ms, got $($elapsed.TotalMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
