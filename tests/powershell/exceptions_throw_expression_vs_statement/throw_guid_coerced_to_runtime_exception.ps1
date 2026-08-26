# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_guid_coerced_to_runtime_exception
$g = [guid]::NewGuid()
$err = $null
try {
    throw $g
} catch {
    $err = $_
}
if ($err.Exception.Message -ne $g.ToString()) {
    Write-Host "FAIL: Throw GUID coerced to message failed"
    exit 1
}
Write-Host "PASS"
exit 0
