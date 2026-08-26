# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_string_creates_runtime_exception
$err = $null
try {
    throw "Simple error string"
} catch {
    $err = $_
}
if ($err.Exception.GetType().Name -ne "RuntimeException" -or $err.Exception.Message -ne "Simple error string") {
    Write-Host "FAIL: Throw string should produce RuntimeException, got $($err.Exception.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
