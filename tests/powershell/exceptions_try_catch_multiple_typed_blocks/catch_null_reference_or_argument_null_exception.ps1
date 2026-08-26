# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_null_reference_or_argument_null_exception
$caught = $false
try {
    throw [System.ArgumentNullException]::new("param")
} catch [System.ArgumentNullException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Catching ArgumentNullException failed"
    exit 1
}
Write-Host "PASS"
exit 0
