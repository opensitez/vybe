# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_operator_equals
# Standard assignment $x = 10 sets Operator to TokenKind.Equals
$code = "`$x = 10"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Operator -ne [System.Management.Automation.Language.TokenKind]::Equals) {
    Write-Host "FAIL: expected TokenKind.Equals, got '$($assignAst.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
