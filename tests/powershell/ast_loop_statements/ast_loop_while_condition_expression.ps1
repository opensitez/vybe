# vybe-test: powershell/ast_loop_statements/ast_loop_while_condition_expression
# WhileStatementAst.Condition captures the test condition expression
$code = "while (`$activeWorkers -gt 0) { Start-Sleep -Milliseconds 50 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$whileAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.WhileStatementAst] }, $true)

if ($whileAst.Condition.Extent.Text -ne "`$activeWorkers -gt 0") {
    Write-Host "FAIL: While condition mismatch: '$($whileAst.Condition.Extent.Text)'"
    exit 1
}

Write-Host "PASS"
exit 0
