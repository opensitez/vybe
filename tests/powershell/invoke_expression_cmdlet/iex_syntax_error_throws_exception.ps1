# vybe-test: powershell/invoke_expression_cmdlet/iex_syntax_error_throws_exception
# Invoke-Expression throws a parse exception when given syntactically invalid PowerShell
$threwError = $false
try {
    Invoke-Expression "1 + (2 *" -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: syntax error string did not throw an exception"
    exit 1
}

Write-Host "PASS"
exit 0
