# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_operator_remainder_equals
# Modulo assignment $x %= 3 sets Operator to TokenKind.RemainderEquals
$code = "`$value %= 3"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Operator -ne [System.Management.Automation.Language.TokenKind]::RemainderEquals) {
    Write-Host "FAIL: expected TokenKind.RemainderEquals, got '$($assignAst.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
