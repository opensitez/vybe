# vybe-test: powershell/exceptions_throw_expression_vs_statement/rethrow_naked_throw_inside_catch_block
$caught = $null
try {
    try {
        throw [System.TimeoutException]::new("Network timeout")
    } catch {
        throw # naked rethrow
    }
} catch {
    $caught = $_
}
if ($caught.Exception -isnot [System.TimeoutException] -or $caught.Exception.Message -ne "Network timeout") {
    Write-Host "FAIL: Naked rethrow inside catch failed"
    exit 1
}
Write-Host "PASS"
exit 0
