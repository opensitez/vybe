# vybe-test: powershell/error_record_exception_chaining/rethrow_wrapping_in_try_catch
$caught = $false
try {
    try {
        1 / 0
    } catch {
        throw [System.InvalidOperationException]::new("Wrapped", $_.Exception)
    }
} catch {
    $caught = ($_.Exception.InnerException -ne $null)
}
if (-not $caught) {
    Write-Host "FAIL: Wrapped exception catch failed"
    exit 1
}
Write-Host "PASS"
exit 0
