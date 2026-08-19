# vybe-test: powershell/lexical_escape_rules/escape_backtick_dollar
$value = "\$HOME"
if ($value -ne '$HOME') {
    Write-Host "FAIL: expected literal $HOME text, got $value"
    exit 1
}

Write-Host 'PASS'
exit 0
