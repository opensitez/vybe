# vybe-test: powershell/ast_error_handling_ast/ast_error_try_body_statement_block
# TryStatementAst.Body contains the statement block enclosed by the try block
$code = @"
try {
    `$x = 10
    `$y = 20
} catch {
    `$z = 30
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$tryAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($tryAst.Body.Statements.Count -ne 2) {
    Write-Host "FAIL: expected 2 statements in try body, got $($tryAst.Body.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
