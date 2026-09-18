# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_right_scriptblock_literal
# Assigning a scriptblock literal evaluates Right.Expression as a ScriptBlockExpressionAst
$code = "`$action = { Write-Host 'Executing block' }"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$cmdExpr = $assignAst.Right
$expr = $cmdExpr.Expression

if ($expr -isnot [System.Management.Automation.Language.ScriptBlockExpressionAst]) {
    Write-Host "FAIL: expected ScriptBlockExpressionAst, got '$($expr.GetType().Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
