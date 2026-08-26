# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/dollar_under_contains_caught_error_record
$msg = ""
try {
    throw [System.InvalidOperationException]::new("DetailedFailure")
} catch [System.InvalidOperationException] {
    $msg = $_.Exception.Message
}
if ($msg -ne "DetailedFailure") {
    Write-Host "FAIL: `$_ in typed catch block failed, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
