# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_empty_string
$err = $null
try {
    throw ""
} catch {
    $err = $_
}
if ($err.Exception -eq $null -or $err.Exception.Message -ne "") {
    Write-Host "FAIL: Throw empty string failed"
    exit 1
}
Write-Host "PASS"
exit 0
