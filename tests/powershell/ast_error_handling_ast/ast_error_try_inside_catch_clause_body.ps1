# vybe-test: powershell/ast_error_handling_ast/ast_error_try_inside_catch_clause_body
# A TryStatementAst nested inside a CatchClauseAst.Body is properly located in the AST hierarchy
$code = @"
try {
    Write-Host "primary"
} catch {
    try {
        Write-Host "fallback"
    } catch {
        Write-Host "fatal"
    }
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$outerTry = $ast.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $false)
$catchClause = $outerTry.CatchClauses[0]

$fallbackTry = $catchClause.Body.Find({ $args[0] -is [System.Management.Automation.Language.TryStatementAst] }, $true)

if ($fallbackTry -eq $null) {
    Write-Host "FAIL: fallback TryStatementAst not found in catch body"
    exit 1
}

Write-Host "PASS"
exit 0
