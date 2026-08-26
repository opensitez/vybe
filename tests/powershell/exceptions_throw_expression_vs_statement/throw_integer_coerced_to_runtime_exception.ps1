# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_integer_coerced_to_runtime_exception
$err = $null
try {
    throw 404
} catch {
    $err = $_
}
if ($err.Exception.Message -ne "404") {
    Write-Host "FAIL: Throw integer coerced to string message failed, got '$($err.Exception.Message)'"
    exit 1
}
Write-Host "PASS"
exit 0
