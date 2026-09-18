# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_operator_plus_equals
# Additive assignment $x += 5 sets Operator to TokenKind.PlusEquals
$code = "`$counter += 5"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Operator -ne [System.Management.Automation.Language.TokenKind]::PlusEquals) {
    Write-Host "FAIL: expected TokenKind.PlusEquals, got '$($assignAst.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
