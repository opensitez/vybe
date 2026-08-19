# vybe-test: powershell/whitespace_and_line_rules/backtick_continuation_whitespace
$expr = "1 + `t`n2"
$errors = @()
$tokens = [System.Management.Automation.PSParser]::Tokenize($expr, [ref]$errors)
if ($errors.Count -ne 0) {
    Write-Host 'FAIL: tab-whitespace line continuation with backtick did not tokenize'
    exit 1
}

if (Invoke-Expression $expr -ne 3) {
    Write-Host 'FAIL: backtick + whitespace continuation did not evaluate as 3'
    exit 1
}

Write-Host 'PASS'
exit 0
