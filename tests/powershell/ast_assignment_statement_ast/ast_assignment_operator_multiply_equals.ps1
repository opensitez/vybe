# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_operator_multiply_equals
# Multiplicative assignment $x *= 4 sets Operator to TokenKind.MultiplyEquals
$code = "`$factor *= 4"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Operator -ne [System.Management.Automation.Language.TokenKind]::MultiplyEquals) {
    Write-Host "FAIL: expected TokenKind.MultiplyEquals, got '$($assignAst.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
