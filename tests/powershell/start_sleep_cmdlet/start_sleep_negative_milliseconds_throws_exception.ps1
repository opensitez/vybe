# vybe-test: powershell/start_sleep_cmdlet/start_sleep_negative_milliseconds_throws_exception
# Specifying a negative value for -Milliseconds throws an ArgumentOutOfRangeException
$threwError = $false
try {
    Start-Sleep -Milliseconds -100 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: negative -Milliseconds value did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
