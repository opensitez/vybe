# vybe-test: powershell/start_sleep_cmdlet/start_sleep_exclusive_seconds_and_milliseconds_throws
# Specifying both -Seconds and -Milliseconds violates parameter set exclusivity and throws
$threwError = $false
try {
    Start-Sleep -Seconds 1 -Milliseconds 100 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: combining -Seconds and -Milliseconds did not throw parameter binding error"
    exit 1
}

Write-Host "PASS"
exit 0
