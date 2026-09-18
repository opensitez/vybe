# vybe-test: powershell/invoke_expression_cmdlet/iex_throw_statement_rethrows_in_caller
# A 'throw' embedded in an Invoke-Expression string is caught by the calling try/catch block
$caught = $false
try {
    Invoke-Expression "throw 'dynamic error'" -ErrorAction Stop
} catch {
    $caught = $true
}

if (-not $caught) {
    Write-Host "FAIL: throw inside Invoke-Expression was not rethrown to caller"
    exit 1
}

Write-Host "PASS"
exit 0
