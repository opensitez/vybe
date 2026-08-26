# vybe-test: powershell/error_record_category_info/write_error_with_errorid_parameter
function Test-ErrorIdEmit {
    [CmdletBinding()]
    param()
    Write-Error -Message "Validation failed" -ErrorId "SchemaValidationError_101"
}
$errRecord = $null
try {
    Test-ErrorIdEmit -ErrorAction Stop
} catch {
    $errRecord = $_
}
if ($errRecord.FullyQualifiedErrorId -notmatch "SchemaValidationError_101") {
    Write-Host "FAIL: ErrorId check failed, got '$($errRecord.FullyQualifiedErrorId)'"
    exit 1
}
Write-Host "PASS"
exit 0
