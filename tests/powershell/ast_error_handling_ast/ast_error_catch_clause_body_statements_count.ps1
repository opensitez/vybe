# vybe-test: powershell/ast_error_handling_ast/ast_error_catch_clause_body_statements_count
# Statements declared inside catch {} populate CatchClauseAst.Body.Statements
$code = @"
try {
    `$x = 1
} catch {
    `$err = `$_.Exception.Message
    Write-Warning `$err
    `$recovered = `$true
}
"@

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$catchAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.CatchClauseAst] }, $true)

if ($catchAst.Body.Statements.Count -ne 3) {
    Write-Host "FAIL: expected 3 statements in catch body, got $($catchAst.Body.Statements.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
