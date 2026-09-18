# vybe-test: powershell/start_sleep_cmdlet/start_sleep_named_milliseconds_parameter
# The -Milliseconds parameter accepts integer values (System.Int32)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Start-Sleep -Milliseconds 50
$sw.Stop()

if ($sw.ElapsedMilliseconds -lt 30) {
    Write-Host "FAIL: named -Milliseconds sleep insufficient, got $($sw.ElapsedMilliseconds)ms"
    exit 1
}

Write-Host "PASS"
exit 0
