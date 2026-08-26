# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_formatted_string_interpolation
$item = "file.txt"
$err = $null
try {
    throw "Could not find file: $item"
} catch {
    $err = $_
}
if ($err.Exception.Message -ne "Could not find file: file.txt") {
    Write-Host "FAIL: Throw formatted string interpolation failed, got '$($err.Exception.Message)'"
    exit 1
}
Write-Host "PASS"
exit 0
