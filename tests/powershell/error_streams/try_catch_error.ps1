# vybe-test: powershell/error_streams/try_catch_error
try {
    throw 'boom'
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: expected catch block"
    exit 1
}
Write-Host 'PASS'
exit 0
