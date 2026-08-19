# vybe-test: powershell/token_and_grammar_system/identifier_dash_invalid
$code = '$bad-name = 1'
$errors = @()
[System.Management.Automation.PSParser]::Tokenize($code, [ref]$errors) | Out-Null
if ($errors.Count -eq 0) {
    Write-Host 'FAIL: expected parser to reject dashed identifier token'
    exit 1
}

Write-Host 'PASS'
exit 0
