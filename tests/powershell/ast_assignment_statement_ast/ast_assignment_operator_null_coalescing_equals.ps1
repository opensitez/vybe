# vybe-test: powershell/ast_assignment_statement_ast/ast_assignment_operator_null_coalescing_equals
# Null-coalescing assignment $x ??= 'fallback' sets Operator to TokenKind.QuestionQuestionEquals
$code = "`$configValue ??= 'default_setting'"

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($code, [ref]$tokens, [ref]$errors)

$assignAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)

if ($assignAst.Operator -ne [System.Management.Automation.Language.TokenKind]::QuestionQuestionEquals) {
    Write-Host "FAIL: expected TokenKind.QuestionQuestionEquals, got '$($assignAst.Operator)'"
    exit 1
}

Write-Host "PASS"
exit 0
