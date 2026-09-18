# vybe-test: powershell/start_sleep_cmdlet/start_sleep_negative_duration_timespan_throws
# Passing a negative TimeSpan to -Duration throws an ArgumentOutOfRangeException
$threwError = $false
try {
    Start-Sleep -Duration ([TimeSpan]::FromSeconds(-1)) -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: negative TimeSpan duration did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
