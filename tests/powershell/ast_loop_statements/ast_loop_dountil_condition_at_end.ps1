# vybe-test: powershell/ast_loop_statements/ast_loop_dountil_condition_at_end
# DoUntilStatementAst cleanly encapsulates the until termination condition
$code = "do { `$attempts++ } until (`$attempts -ge 10)"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$doUntilAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.DoUntilStatementAst] }, $true)

if ($doUntilAst -eq $null) {
    Write-Host "FAIL: DoUntilStatementAst not found"
    exit 1
}

if ($doUntilAst.Condition.Extent.Text -ne "`$attempts -ge 10") {
    Write-Host "FAIL: until condition mismatch: '$($doUntilAst.Condition.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
