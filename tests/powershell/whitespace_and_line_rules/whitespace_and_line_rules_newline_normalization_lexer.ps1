# vybe-test: powershell/whitespace_and_line_rules/newline_normalization_lexer
$script = "3 + 1`r`n+ 2"
$errors = @()
[void][System.Management.Automation.PSParser]::Tokenize($script, [ref]$errors)
if ($errors.Count -ne 0) {
    Write-Host 'FAIL: mixed newline sequence broke tokenization'
    exit 1
}

if ((Invoke-Expression $script) -ne 6) {
    Write-Host 'FAIL: newline normalization changed expression value'
    exit 1
}

Write-Host 'PASS'
exit 0
