# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_right_binary_expression
# Assigning a calculated expression evaluates Right.Expression as a BinaryExpressionAst
$code = "`$total = `$subtotal * `$tax"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

$cmdExpr = $assignAst.Right
$expr = $cmdExpr.Expression

if ($expr -isnot [System.Management.Automation.Language.BinaryExpressionAst]) {
    Write-Host "FAIL: expected BinaryExpressionAst, got '$($expr.GetType().Name)'"
    exit 1
}

if ($expr.Operator -ne [System.Management.Automation.Language.TokenKind]::Multiply) {
    Write-Host "FAIL: operator mismatch: '$($expr.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
