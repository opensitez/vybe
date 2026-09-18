# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_right_hashtable_literal
# Assigning a hashtable literal evaluates Right.Expression as a HashtableAst
$code = "`$serverConfig = @{ Host = '127.0.0.1'; Port = 443 }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$cmdExpr = $assignAst.Right
$expr = $cmdExpr.Expression

if ($expr -isnot [System.Management.Automation.Language.HashtableAst]) {
    Write-Host "FAIL: expected HashtableAst, got '$($expr.GetType().Name)'"
    exit 1
}

if ($expr.KeyValuePairs.Count -ne 2) {
    Write-Host "FAIL: expected 2 key-value pairs, got $($expr.KeyValuePairs.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
