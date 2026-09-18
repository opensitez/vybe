# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_operator_divide_equals
# Division assignment $x /= 2 sets Operator to TokenKind.DivideEquals
$code = "`$quotient /= 2"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Operator -ne [System.Management.Automation.Language.TokenKind]::DivideEquals) {
    Write-Host "FAIL: expected TokenKind.DivideEquals, got '$($assignAst.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
