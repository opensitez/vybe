# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_in_subexpression
$caught = $false
try {
    throw "TargetError"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Throw failed"
    exit 1
}
Write-Host "PASS"
exit 0
