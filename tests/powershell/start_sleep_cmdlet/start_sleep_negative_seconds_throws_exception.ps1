# vybe-test: powershell/start_sleep_cmdlet/start_sleep_negative_seconds_throws_exception
# Specifying a negative value for -Seconds throws an ArgumentOutOfRangeException
$threwError = $false
try {
    Start-Sleep -Seconds -1 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: negative -Seconds value did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
