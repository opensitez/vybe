# vybe-test: powershell/start_sleep_cmdlet/start_sleep_exclusive_milliseconds_and_duration_throws
# Specifying both -Milliseconds and -Duration violates parameter set exclusivity and throws
$threwError = $false
try {
    Start-Sleep -Milliseconds 100 -Duration ([TimeSpan]::FromSeconds(1)) -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: combining -Milliseconds and -Duration did not throw parameter binding error"
    exit 1
}

Write-Host "PASS"
exit 0
