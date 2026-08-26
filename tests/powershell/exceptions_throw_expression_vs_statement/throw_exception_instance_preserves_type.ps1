# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_exception_instance_preserves_type
$err = $null
try {
    throw [System.IO.FileNotFoundException]::new("data.csv")
} catch {
    $err = $_
}
if ($err.Exception -isnot [System.IO.FileNotFoundException]) {
    Write-Host "FAIL: Throw exception instance type preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
