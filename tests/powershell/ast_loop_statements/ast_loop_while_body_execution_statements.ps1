# vybe-test: powershell/ast_loop_statements/ast_loop_while_body_execution_statements
# WhileStatementAst.Body contains the body statements of the while loop
$code = @"
while (`$pending) {
    `$task = Get-NextTask
    `$task.Run()
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$whileAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.WhileStatementAst] }, $true)

if ($whileAst.Body.Statements.Count -ne 2) {
    Write-Host "FAIL: expected 2 statements in while body, got $($whileAst.Body.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
