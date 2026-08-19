# vybe-test: powershell/whitespace_and_line_rules/backtick_continuation_comment
$expr = "2 + ` # continues across comment`n1"
$errors = @()
[void][System.Management.Automation.PSParser]::Tokenize($expr, [ref]$errors)
if ($errors.Count -ne 0) {
    Write-Host 'FAIL: backtick continuation with comment did not tokenize'
    exit 1
}

if (Invoke-Expression $expr -ne 3) {
    Write-Host 'FAIL: backtick+comment continuation miscomputed'
    exit 1
}

Write-Host 'PASS'
exit 0
